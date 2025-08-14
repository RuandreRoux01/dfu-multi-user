const express = require('express');
const http = require('http');
const socketIo = require('socket.io');
const cors = require('cors');
const multer = require('multer');
const { MongoClient } = require('mongodb');
const XLSX = require('xlsx');
const path = require('path');
require('dotenv').config();

const app = express();
const server = http.createServer(app);
const io = socketIo(server, {
    cors: {
        origin: "*",
        methods: ["GET", "POST"]
    }
});

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' })); // Increase limit for large data
app.use(express.static('public'));

// Debug middleware for API routes
app.use('/api', (req, res, next) => {
    console.log(`[API] ${req.method} ${req.path}`);
    console.log(`[API] Full URL: ${req.originalUrl}`);
    next();
});

// MongoDB connection
let db;
const mongoUri = process.env.MONGODB_URI;

// Add connection status flag
let isDbConnected = false;

MongoClient.connect(mongoUri)
    .then(client => {
        console.log('[DB] Connected to MongoDB Atlas!');
        db = client.db('dfu_transfer_db');
        isDbConnected = true;
        initializeCollections();
    })
    .catch(error => {
        console.error('[DB] MongoDB connection error:', error);
        console.log('Please check your connection string in .env file');
    });

// Initialize database collections
async function initializeCollections() {
    try {
        await db.createCollection('sessions');
        await db.createCollection('transfers');
        console.log('[DB] Database collections ready!');
    } catch (error) {
        console.log('[DB] Collections might already exist, that\'s okay!');
    }
}

// File upload setup
const storage = multer.memoryStorage();
const upload = multer({ 
    storage: storage,
    limits: { fileSize: 50 * 1024 * 1024 } // 50MB limit
});

// Store active sessions in memory
const activeSessions = new Map();

// Database connection check middleware
const checkDbConnection = (req, res, next) => {
    if (!isDbConnected || !db) {
        console.error('[DB] Database not connected - request blocked');
        return res.status(503).json({ 
            error: 'Database connection not ready. Please try again in a moment.' 
        });
    }
    next();
};

// ============= REST API ENDPOINTS =============

// Health check endpoint
app.get('/api/health', (req, res) => {
    res.json({ 
        status: 'Server is running!', 
        database: isDbConnected ? 'Connected' : 'Not connected',
        timestamp: new Date().toISOString()
    });
});

// Create or join a session - with DB check
app.post('/api/session/join', checkDbConnection, async (req, res) => {
    try {
        const { sessionName, userName } = req.body;
        
        console.log(`[SESSION] Join request - Session: ${sessionName}, User: ${userName}`);
        
        if (!sessionName || !userName) {
            return res.status(400).json({ error: 'Session name and user name are required' });
        }
        
        // Find or create session by name
        let session = await db.collection('sessions').findOne({ name: sessionName });
        
        if (!session) {
            // Create new session with a simple string ID
            const sessionId = `session_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
            session = {
                _id: sessionId,
                name: sessionName,
                createdAt: new Date(),
                users: [],
                dataUploaded: false,
                rawData: null,
                variantCycleData: null,
                hasVariantCycleData: false,
                completedTransfers: {},
                status: 'active'
            };
            await db.collection('sessions').insertOne(session);
            console.log(`[SESSION] New session created: ${sessionName} with ID: ${sessionId}`);
        } else {
            console.log(`[SESSION] Existing session found: ${session._id}`);
        }
        
        // Add user to session if not already present
        const existingUser = session.users ? session.users.find(u => u.name === userName) : null;
        if (!existingUser) {
            const userId = `user_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
            await db.collection('sessions').updateOne(
                { _id: session._id },
                { 
                    $push: { users: { id: userId, name: userName, joinedAt: new Date() } }
                }
            );
        }
        
        console.log(`[SESSION] ${userName} joined session: ${sessionName} (${session._id})`);
        
        res.json({ 
            success: true,
            sessionId: session._id,
            sessionName: session.name,
            userId: existingUser ? existingUser.id : `user_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`
        });
    } catch (error) {
        console.error('[SESSION] Error joining session:', error);
        res.status(500).json({ error: 'Failed to join session: ' + error.message });
    }
});

// Upload Excel file - with better error handling
app.post('/api/upload/:sessionId', checkDbConnection, upload.single('file'), async (req, res) => {
    try {
        const { sessionId } = req.params;
        const file = req.file;
        
        console.log(`[UPLOAD] Processing file for session: ${sessionId}`);
        console.log(`[UPLOAD] File info: ${file ? `${file.originalname} (${file.size} bytes)` : 'No file'}`);
        
        if (!file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }
        
        // Check if session exists
        const sessionExists = await db.collection('sessions').findOne({ _id: sessionId });
        if (!sessionExists) {
            console.error(`[UPLOAD] Session not found: ${sessionId}`);
            return res.status(404).json({ error: `Session not found: ${sessionId}` });
        }
        
        console.log(`[UPLOAD] Session found: ${sessionExists.name}`);
        
        // Parse Excel file
        const workbook = XLSX.read(file.buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet);
        
        console.log(`[UPLOAD] Parsed ${data.length} rows from Excel file`);
        
        if (data.length === 0) {
            return res.status(400).json({ error: 'Excel file contains no data' });
        }
        
        // Store raw data in session
        const updateResult = await db.collection('sessions').updateOne(
            { _id: sessionId },
            { 
                $set: { 
                    rawData: data,
                    dataUploaded: true,
                    uploadedAt: new Date(),
                    completedTransfers: {} // Reset completed transfers on new upload
                }
            }
        );
        
        console.log(`[UPLOAD] Update result: ${updateResult.modifiedCount} document(s) modified`);
        
        if (updateResult.modifiedCount === 0 && updateResult.matchedCount === 0) {
            console.error(`[UPLOAD] Failed to update session - session might have been deleted`);
            return res.status(404).json({ error: 'Session no longer exists' });
        }
        
        // Process multi-variant DFUs for summary
        const multiVariantDFUs = processMultiVariantDFUs(data);
        console.log(`[UPLOAD] Found ${Object.keys(multiVariantDFUs).length} multi-variant DFUs`);
        
        res.json({ 
            success: true, 
            rowCount: data.length,
            dfuCount: Object.keys(multiVariantDFUs).length 
        });
        
        // Notify all connected clients
        io.to(sessionId).emit('dataUploaded', { 
            dfuCount: Object.keys(multiVariantDFUs).length,
            uploadedBy: req.body.userName || 'Unknown User'
        });
        
    } catch (error) {
        console.error('[UPLOAD] Error uploading file:', error);
        res.status(500).json({ error: 'Failed to process file: ' + error.message });
    }
});

// Upload Variant Cycle Dates file
app.post('/api/upload-cycle/:sessionId', checkDbConnection, upload.single('file'), async (req, res) => {
    try {
        const { sessionId } = req.params;
        const file = req.file;
        
        if (!file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }
        
        console.log(`[CYCLE] Processing variant cycle dates file for session: ${sessionId}`);
        
        // Check if session exists
        const sessionExists = await db.collection('sessions').findOne({ _id: sessionId });
        if (!sessionExists) {
            return res.status(404).json({ error: 'Session not found' });
        }
        
        // Parse Excel file
        const workbook = XLSX.read(file.buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const cycleData = XLSX.utils.sheet_to_json(worksheet);
        
        console.log(`[CYCLE] Parsed ${cycleData.length} cycle date records`);
        
        // Process the cycle data into a more usable format
        const processedCycleData = {};
        cycleData.forEach(record => {
            // Find the DFU and Part Code columns (case-insensitive)
            const dfuCode = record['DFU'] || record['dfu'] || record['Dfu'];
            const partCode = record['Part Code'] || record['PART CODE'] || record['Part_Code'] || 
                           record['Product Number'] || record['Part Number'];
            const sos = record['SOS'] || record['sos'] || record['Sos'];
            const eos = record['EOS'] || record['eos'] || record['Eos'];
            const comments = record['Comments'] || record['COMMENTS'] || record['Comment'] || '';
            
            if (dfuCode && partCode) {
                if (!processedCycleData[dfuCode]) {
                    processedCycleData[dfuCode] = {};
                }
                processedCycleData[dfuCode][partCode] = {
                    sos: sos || 'N/A',
                    eos: eos || 'N/A',
                    comments: comments
                };
            }
        });
        
        console.log(`[CYCLE] Processed cycle data for ${Object.keys(processedCycleData).length} DFUs`);
        
        // Store cycle data in session
        const updateResult = await db.collection('sessions').updateOne(
            { _id: sessionId },
            { 
                $set: { 
                    variantCycleData: processedCycleData,
                    hasVariantCycleData: true,
                    cycleDataUploadedAt: new Date()
                }
            }
        );
        
        console.log(`[CYCLE] Update result: ${updateResult.modifiedCount} document(s) modified`);
        
        res.json({ 
            success: true, 
            recordCount: cycleData.length,
            dfuCount: Object.keys(processedCycleData).length
        });
        
        // Notify all connected clients that cycle data was uploaded
        io.to(sessionId).emit('cycleDataUploaded', { 
            dfuCount: Object.keys(processedCycleData).length,
            uploadedBy: req.body.userName || 'Unknown User'
        });
        
    } catch (error) {
        console.error('[CYCLE] Error uploading cycle data file:', error);
        res.status(500).json({ error: 'Failed to process cycle data file: ' + error.message });
    }
});

// Get session data
app.get('/api/session/:sessionId/data', checkDbConnection, async (req, res) => {
    try {
        const { sessionId } = req.params;
        
        console.log(`[DATA] Getting data for session: ${sessionId}`);
        
        const session = await db.collection('sessions').findOne({ 
            _id: sessionId
        });
        
        if (!session) {
            console.log(`[DATA] Session not found: ${sessionId}`);
            // Instead of 404, return empty data structure
            return res.json({
                session: {
                    name: 'Session Not Found',
                    dataUploaded: false,
                    userCount: 0,
                    hasVariantCycleData: false
                },
                multiVariantDFUs: {},
                transfers: [],
                rawData: [],
                completedTransfers: {},
                variantCycleData: {},
                hasVariantCycleData: false
            });
        }
        
        console.log(`[DATA] Found session: ${session.name}`);
        console.log(`[DATA] Data uploaded: ${session.dataUploaded || false}`);
        console.log(`[DATA] Raw data records: ${session.rawData ? session.rawData.length : 0}`);
        console.log(`[DATA] Has cycle data: ${session.hasVariantCycleData || false}`);
        
        // Process multi-variant DFUs if data exists
        let multiVariantDFUs = {};
        if (session.rawData && session.rawData.length > 0) {
            multiVariantDFUs = processMultiVariantDFUs(session.rawData);
            console.log(`[DATA] Processed ${Object.keys(multiVariantDFUs).length} multi-variant DFUs`);
        }
        
        // Get transfers for this session
        const transfers = await db.collection('transfers').find({ 
            sessionId: sessionId 
        }).toArray();
        
        console.log(`[DATA] Found ${transfers.length} transfers for session`);
        
        res.json({
            session: {
                name: session.name,
                dataUploaded: session.dataUploaded || false,
                userCount: session.users ? session.users.length : 0,
                hasVariantCycleData: session.hasVariantCycleData || false
            },
            multiVariantDFUs: multiVariantDFUs || {},
            transfers: transfers || [],
            rawData: session.rawData || [],
            completedTransfers: session.completedTransfers || {},
            variantCycleData: session.variantCycleData || {},
            hasVariantCycleData: session.hasVariantCycleData || false
        });
        
    } catch (error) {
        console.error('[DATA] Error getting session data:', error);
        // Return empty structure instead of error
        res.json({
            session: {
                name: 'Error Loading Session',
                dataUploaded: false,
                userCount: 0,
                hasVariantCycleData: false
            },
            multiVariantDFUs: {},
            transfers: [],
            rawData: [],
            completedTransfers: {},
            variantCycleData: {},
            hasVariantCycleData: false
        });
    }
});

// Save transfer configuration
app.post('/api/session/:sessionId/transfer', checkDbConnection, async (req, res) => {
    try {
        const { sessionId } = req.params;
        const { dfuCode, transferConfig, userName } = req.body;
        
        await db.collection('transfers').updateOne(
            { sessionId, dfuCode },
            { 
                $set: { 
                    ...transferConfig,
                    sessionId,
                    dfuCode,
                    updatedBy: userName,
                    updatedAt: new Date()
                }
            },
            { upsert: true }
        );
        
        console.log(`[TRANSFER] Transfer saved for DFU ${dfuCode} by ${userName}`);
        
        res.json({ success: true });
        
        // Notify other users
        io.to(sessionId).emit('transferUpdated', { dfuCode, transferConfig, userName });
        
    } catch (error) {
        console.error('[TRANSFER] Error saving transfer:', error);
        res.status(500).json({ error: 'Failed to save transfer: ' + error.message });
    }
});

// Update session data after client-side transfer
app.post('/api/session/:sessionId/updateData', checkDbConnection, async (req, res) => {
    try {
        const { sessionId } = req.params;
        const { rawData, completedTransfers, transfer } = req.body;
        
        console.log(`[UPDATE] Updating session data after transfer for DFU: ${transfer.dfuCode}`);
        
        // Update the session with the modified data and completed transfers
        await db.collection('sessions').updateOne(
            { _id: sessionId },
            { 
                $set: { 
                    rawData: rawData,
                    completedTransfers: completedTransfers,
                    lastModified: new Date(),
                    lastModifiedBy: transfer.completedBy
                }
            }
        );
        
        // Log the transfer (only if not an undo operation)
        if (transfer.type !== 'undo') {
            await db.collection('transfers').insertOne({
                sessionId,
                dfuCode: transfer.dfuCode,
                type: transfer.type,
                targetVariant: transfer.targetVariant,
                transfers: transfer.transfers,
                granularTransfers: transfer.granularTransfers,
                completedBy: transfer.completedBy,
                completedAt: new Date()
            });
        }
        
        console.log(`[UPDATE] Data updated successfully`);
        res.json({ success: true });
        
        // Notify other users
        io.to(sessionId).emit('dataUpdated', { 
            dfuCode: transfer.dfuCode,
            updatedBy: transfer.completedBy 
        });
        
    } catch (error) {
        console.error('[UPDATE] Error updating data:', error);
        res.status(500).json({ error: 'Failed to update data' });
    }
});

// Undo transfer - restore original data for a DFU
app.post('/api/session/:sessionId/undoTransfer', checkDbConnection, async (req, res) => {
    try {
        const { sessionId } = req.params;
        const { dfuCode, userName } = req.body;
        
        console.log(`[UNDO] Undoing transfer for DFU ${dfuCode} by ${userName}`);
        
        // Get the current session
        const session = await db.collection('sessions').findOne({ _id: sessionId });
        
        if (!session) {
            return res.status(404).json({ error: 'Session not found' });
        }
        
        // Remove the DFU from completedTransfers
        const completedTransfers = session.completedTransfers || {};
        delete completedTransfers[dfuCode];
        
        // Update the session with the modified completedTransfers
        await db.collection('sessions').updateOne(
            { _id: sessionId },
            { 
                $set: { 
                    completedTransfers: completedTransfers,
                    lastModified: new Date(),
                    lastModifiedBy: userName
                }
            }
        );
        
        // Remove transfer logs for this DFU
        await db.collection('transfers').deleteMany({
            sessionId,
            dfuCode
        });
        
        console.log(`[UNDO] Transfer undone for DFU ${dfuCode}`);
        res.json({ success: true });
        
        // Notify all users that a transfer was undone
        io.to(sessionId).emit('transferUndone', { 
            dfuCode,
            undoneBy: userName
        });
        
    } catch (error) {
        console.error('[UNDO] Error undoing transfer:', error);
        res.status(500).json({ error: 'Failed to undo transfer' });
    }
});

// Export current data
app.post('/api/session/:sessionId/export', checkDbConnection, async (req, res) => {
    try {
        const { sessionId } = req.params;
        const { rawData } = req.body;
        
        console.log(`[EXPORT] Exporting data for session: ${sessionId}`);
        
        // Use the provided rawData (already has transfers applied)
        const dataToExport = rawData || [];
        
        // Create Excel file
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(dataToExport);
        XLSX.utils.book_append_sheet(wb, ws, 'Updated Demand');
        
        // Generate buffer
        const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
        
        console.log(`[EXPORT] Generated file with ${dataToExport.length} records`);
        
        // Send file
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=DFU_Transfers_${Date.now()}.xlsx`);
        res.send(buffer);
        
    } catch (error) {
        console.error('[EXPORT] Error exporting file:', error);
        res.status(500).json({ error: 'Failed to export file' });
    }
});

// Clear session data
app.post('/api/session/:sessionId/clear', checkDbConnection, async (req, res) => {
    try {
        const { sessionId } = req.params;
        
        // Clear all transfers for this session
        await db.collection('transfers').deleteMany({ sessionId });
        
        // Clear raw data and cycle data from session
        await db.collection('sessions').updateOne(
            { _id: sessionId },
            { 
                $set: { 
                    rawData: null,
                    dataUploaded: false,
                    variantCycleData: null,
                    hasVariantCycleData: false,
                    completedTransfers: {}
                }
            }
        );
        
        console.log(`[CLEAR] Cleared all data for session: ${sessionId}`);
        res.json({ success: true });
        
        // Notify all connected users
        io.to(sessionId).emit('dataCleared');
        
    } catch (error) {
        console.error('[CLEAR] Error clearing session:', error);
        res.status(500).json({ error: 'Failed to clear session' });
    }
});

// End session and clear all data
app.post('/api/session/:sessionId/end', checkDbConnection, async (req, res) => {
    try {
        const { sessionId } = req.params;
        const { userName } = req.body;
        
        console.log(`[END] Ending session ${sessionId} by ${userName}`);
        
        // Delete all transfers for this session
        await db.collection('transfers').deleteMany({ sessionId });
        
        // Delete the session
        await db.collection('sessions').deleteOne({ _id: sessionId });
        
        console.log(`[END] Session ${sessionId} ended and all data cleared`);
        
        // Notify all connected users that session has ended
        io.to(sessionId).emit('sessionEnded', { endedBy: userName });
        
        res.json({ success: true, message: 'Session ended and data cleared' });
        
    } catch (error) {
        console.error('[END] Error ending session:', error);
        res.status(500).json({ error: 'Failed to end session: ' + error.message });
    }
});

// Catch-all for unmatched API routes (for debugging)
app.get('/api/*', (req, res) => {
    console.log(`[404] Unmatched API route: ${req.originalUrl}`);
    res.status(404).json({ error: `Route not found: ${req.originalUrl}` });
});

app.post('/api/*', (req, res) => {
    console.log(`[404] Unmatched API POST route: ${req.originalUrl}`);
    res.status(404).json({ error: `Route not found: ${req.originalUrl}` });
});

// ============= WEBSOCKET HANDLING =============

io.on('connection', (socket) => {
    console.log('[WS] New client connected');
    
    socket.on('joinSession', ({ sessionId, userName }) => {
        socket.join(sessionId);
        socket.sessionId = sessionId;
        socket.userName = userName;
        
        // Track active users
        if (!activeSessions.has(sessionId)) {
            activeSessions.set(sessionId, new Set());
        }
        activeSessions.get(sessionId).add(userName);
        
        console.log(`[WS] ${userName} joined session via WebSocket`);
        
        // Notify others
        socket.to(sessionId).emit('userJoined', { userName });
        
        // Send active users list
        io.to(sessionId).emit('activeUsers', Array.from(activeSessions.get(sessionId)));
    });
    
    socket.on('disconnect', () => {
        if (socket.sessionId && socket.userName) {
            const sessionUsers = activeSessions.get(socket.sessionId);
            if (sessionUsers) {
                sessionUsers.delete(socket.userName);
                if (sessionUsers.size === 0) {
                    activeSessions.delete(socket.sessionId);
                } else {
                    io.to(socket.sessionId).emit('activeUsers', Array.from(sessionUsers));
                }
                console.log(`[WS] ${socket.userName} disconnected from ${socket.sessionId}`);
            }
        }
    });
});

// ============= HELPER FUNCTIONS =============

function processMultiVariantDFUs(data) {
    const grouped = {};
    
    data.forEach(record => {
        const dfuCode = record['DFU'];
        if (!dfuCode) return;
        
        if (!grouped[dfuCode]) {
            grouped[dfuCode] = {
                records: [],
                variants: new Set(),
                partDescriptions: {}
            };
        }
        
        const partNumber = record['Product Number'] || record['Part Number'] || record['Part Code'];
        const partDescription = record['PartDescription'] || record['Part Description'] || '';
        
        grouped[dfuCode].records.push(record);
        if (partNumber) {
            grouped[dfuCode].variants.add(partNumber);
            grouped[dfuCode].partDescriptions[partNumber] = partDescription;
        }
    });
    
    const multiVariants = {};
    Object.keys(grouped).forEach(dfuCode => {
        const variants = Array.from(grouped[dfuCode].variants);
        if (variants.length > 1) {
            // Calculate total demand for each variant
            const variantDemand = {};
            variants.forEach(variant => {
                const variantRecords = grouped[dfuCode].records.filter(r => {
                    const partNumber = r['Product Number'] || r['Part Number'] || r['Part Code'];
                    return partNumber === variant;
                });
                
                const totalDemand = variantRecords.reduce((sum, r) => {
                    const demand = parseFloat(r['weekly fcst'] || r['Demand'] || r['Forecast'] || 0);
                    return sum + demand;
                }, 0);
                
                variantDemand[variant] = {
                    totalDemand,
                    recordCount: variantRecords.length,
                    description: grouped[dfuCode].partDescriptions[variant] || ''
                };
            });
            
            multiVariants[dfuCode] = {
                variants: variants,
                recordCount: grouped[dfuCode].records.length,
                variantDemand: variantDemand
            };
        }
    });
    
    return multiVariants;
}

// ============= START SERVER =============

const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
    console.log(`
    [SERVER] Server is running!
    [SERVER] Local: http://localhost:${PORT}
    [SERVER] Network: http://[your-computer-ip]:${PORT}
    
    Waiting for MongoDB connection...
    `);
});