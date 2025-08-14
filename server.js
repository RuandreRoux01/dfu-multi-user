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
app.use(express.json({ limit: '50mb' }));
app.use(express.static('public'));

// Debug middleware
app.use('/api', (req, res, next) => {
    console.log(`[API] ${req.method} ${req.path}`);
    next();
});

// MongoDB connection
let db;
const mongoUri = process.env.MONGODB_URI;
let isDbConnected = false;

// SINGLE SESSION ID - Always use the same one
const TEAM_SESSION_ID = 'TEAM_DFU_TRANSFER_SESSION';

MongoClient.connect(mongoUri)
    .then(client => {
        console.log('[DB] Connected to MongoDB Atlas!');
        db = client.db('dfu_transfer_db');
        isDbConnected = true;
        initializeCollections();
    })
    .catch(error => {
        console.error('[DB] MongoDB connection error:', error);
    });

// Initialize database and ensure team session exists
async function initializeCollections() {
    try {
        await db.createCollection('sessions');
        await db.createCollection('transfers');
        
        // Ensure the team session exists
        const existingSession = await db.collection('sessions').findOne({ _id: TEAM_SESSION_ID });
        if (!existingSession) {
            await db.collection('sessions').insertOne({
                _id: TEAM_SESSION_ID,
                name: 'Team DFU Transfer Session',
                createdAt: new Date(),
                users: [],
                dataUploaded: false,
                rawData: null,
                variantCycleData: null,
                hasVariantCycleData: false,
                completedTransfers: {},
                status: 'active'
            });
            console.log('[DB] Team session created');
        } else {
            console.log('[DB] Team session already exists');
        }
        
        console.log('[DB] Database ready!');
    } catch (error) {
        console.log('[DB] Error initializing:', error);
    }
}

// File upload setup
const storage = multer.memoryStorage();
const upload = multer({ 
    storage: storage,
    limits: { fileSize: 50 * 1024 * 1024 }
});

// Store active users in memory
const activeUsers = new Set();

// Database check middleware
const checkDbConnection = (req, res, next) => {
    if (!isDbConnected || !db) {
        return res.status(503).json({ 
            error: 'Database connection not ready. Please try again.' 
        });
    }
    next();
};

// ============= REST API ENDPOINTS =============

// Health check
app.get('/api/health', (req, res) => {
    res.json({ 
        status: 'Server is running!', 
        database: isDbConnected ? 'Connected' : 'Not connected',
        sessionId: TEAM_SESSION_ID
    });
});

// Join the team session (simplified - just add user)
app.post('/api/session/join', checkDbConnection, async (req, res) => {
    try {
        const { userName } = req.body;
        
        if (!userName) {
            return res.status(400).json({ error: 'User name is required' });
        }
        
        console.log(`[SESSION] ${userName} joining team session`);
        
        // Add user to active users list
        const userId = `user_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
        
        // Update session with user info (optional - for persistence)
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            { 
                $addToSet: { 
                    users: { 
                        id: userId, 
                        name: userName, 
                        joinedAt: new Date() 
                    }
                }
            }
        );
        
        console.log(`[SESSION] ${userName} joined successfully`);
        
        res.json({ 
            success: true,
            sessionId: TEAM_SESSION_ID,  // Always the same
            sessionName: 'Team Session',
            userId: userId,
            userName: userName
        });
    } catch (error) {
        console.error('[SESSION] Error joining:', error);
        res.status(500).json({ error: 'Failed to join session' });
    }
});

// Upload Excel file (simplified)
app.post('/api/upload', checkDbConnection, upload.single('file'), async (req, res) => {
    try {
        const file = req.file;
        const userName = req.body.userName || 'Unknown User';
        
        if (!file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }
        
        console.log(`[UPLOAD] Processing file from ${userName}`);
        
        // Parse Excel file
        const workbook = XLSX.read(file.buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet);
        
        console.log(`[UPLOAD] Parsed ${data.length} rows`);
        
        // Store in the team session
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            { 
                $set: { 
                    rawData: data,
                    dataUploaded: true,
                    uploadedAt: new Date(),
                    uploadedBy: userName,
                    completedTransfers: {}
                }
            }
        );
        
        // Process multi-variant DFUs
        const multiVariantDFUs = processMultiVariantDFUs(data);
        
        res.json({ 
            success: true, 
            rowCount: data.length,
            dfuCount: Object.keys(multiVariantDFUs).length 
        });
        
        // Notify all connected users
        io.emit('dataUploaded', { 
            dfuCount: Object.keys(multiVariantDFUs).length,
            uploadedBy: userName
        });
        
    } catch (error) {
        console.error('[UPLOAD] Error:', error);
        res.status(500).json({ error: 'Failed to process file' });
    }
});

// Upload variant cycle dates
app.post('/api/upload-cycle', checkDbConnection, upload.single('file'), async (req, res) => {
    try {
        const file = req.file;
        const userName = req.body.userName || 'Unknown User';
        
        if (!file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }
        
        console.log(`[CYCLE] Processing cycle file from ${userName}`);
        
        // Parse Excel file
        const workbook = XLSX.read(file.buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const cycleData = XLSX.utils.sheet_to_json(worksheet);
        
        // Process cycle data
        const processedCycleData = {};
        cycleData.forEach(record => {
            const dfuCode = record['DFU'] || record['dfu'];
            const partCode = record['Part Code'] || record['Product Number'];
            const sos = record['SOS'] || record['sos'];
            const eos = record['EOS'] || record['eos'];
            const comments = record['Comments'] || '';
            
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
        
        // Update session
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            { 
                $set: { 
                    variantCycleData: processedCycleData,
                    hasVariantCycleData: true,
                    cycleDataUploadedAt: new Date(),
                    cycleDataUploadedBy: userName
                }
            }
        );
        
        res.json({ 
            success: true, 
            recordCount: cycleData.length,
            dfuCount: Object.keys(processedCycleData).length
        });
        
        // Notify all users
        io.emit('cycleDataUploaded', { 
            dfuCount: Object.keys(processedCycleData).length,
            uploadedBy: userName
        });
        
    } catch (error) {
        console.error('[CYCLE] Error:', error);
        res.status(500).json({ error: 'Failed to process cycle data' });
    }
});

// Get session data (simplified)
app.get('/api/session/data', checkDbConnection, async (req, res) => {
    try {
        const session = await db.collection('sessions').findOne({ _id: TEAM_SESSION_ID });
        
        if (!session) {
            // Should never happen, but handle it
            return res.json({
                session: { name: 'Team Session', dataUploaded: false },
                multiVariantDFUs: {},
                rawData: [],
                completedTransfers: {},
                variantCycleData: {},
                hasVariantCycleData: false
            });
        }
        
        // Process multi-variant DFUs
        let multiVariantDFUs = {};
        if (session.rawData && session.rawData.length > 0) {
            multiVariantDFUs = processMultiVariantDFUs(session.rawData);
        }
        
        res.json({
            session: {
                name: session.name,
                dataUploaded: session.dataUploaded || false,
                userCount: activeUsers.size,
                hasVariantCycleData: session.hasVariantCycleData || false
            },
            multiVariantDFUs: multiVariantDFUs,
            rawData: session.rawData || [],
            completedTransfers: session.completedTransfers || {},
            variantCycleData: session.variantCycleData || {},
            hasVariantCycleData: session.hasVariantCycleData || false
        });
        
    } catch (error) {
        console.error('[DATA] Error:', error);
        res.status(500).json({ error: 'Failed to get data' });
    }
});

// Update data after transfer
app.post('/api/updateData', checkDbConnection, async (req, res) => {
    try {
        const { rawData, completedTransfers, transfer } = req.body;
        
        console.log(`[UPDATE] Transfer by ${transfer.completedBy} for DFU ${transfer.dfuCode}`);
        
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            { 
                $set: { 
                    rawData: rawData,
                    completedTransfers: completedTransfers,
                    lastModified: new Date(),
                    lastModifiedBy: transfer.completedBy
                }
            }
        );
        
        res.json({ success: true });
        
        // Notify all users
        io.emit('dataUpdated', { 
            dfuCode: transfer.dfuCode,
            updatedBy: transfer.completedBy 
        });
        
    } catch (error) {
        console.error('[UPDATE] Error:', error);
        res.status(500).json({ error: 'Failed to update data' });
    }
});

// Undo transfer
app.post('/api/undoTransfer', checkDbConnection, async (req, res) => {
    try {
        const { dfuCode, userName } = req.body;
        
        console.log(`[UNDO] ${userName} undoing transfer for DFU ${dfuCode}`);
        
        // Get current session
        const session = await db.collection('sessions').findOne({ _id: TEAM_SESSION_ID });
        const completedTransfers = session.completedTransfers || {};
        delete completedTransfers[dfuCode];
        
        // Update session
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            { 
                $set: { 
                    completedTransfers: completedTransfers,
                    lastModified: new Date(),
                    lastModifiedBy: userName
                }
            }
        );
        
        res.json({ success: true });
        
        // Notify all users
        io.emit('transferUndone', { 
            dfuCode,
            undoneBy: userName
        });
        
    } catch (error) {
        console.error('[UNDO] Error:', error);
        res.status(500).json({ error: 'Failed to undo transfer' });
    }
});

// Export data
app.post('/api/export', checkDbConnection, async (req, res) => {
    try {
        const { rawData } = req.body;
        
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(rawData || []);
        XLSX.utils.book_append_sheet(wb, ws, 'Updated Demand');
        
        const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=DFU_Transfers_${Date.now()}.xlsx`);
        res.send(buffer);
        
    } catch (error) {
        console.error('[EXPORT] Error:', error);
        res.status(500).json({ error: 'Failed to export' });
    }
});

// Clear all data (reset the session)
app.post('/api/clear', checkDbConnection, async (req, res) => {
    try {
        const { userName } = req.body;
        
        console.log(`[CLEAR] ${userName} clearing all data`);
        
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            { 
                $set: { 
                    rawData: null,
                    dataUploaded: false,
                    variantCycleData: null,
                    hasVariantCycleData: false,
                    completedTransfers: {},
                    clearedAt: new Date(),
                    clearedBy: userName
                }
            }
        );
        
        res.json({ success: true });
        
        // Notify all users
        io.emit('dataCleared', { clearedBy: userName });
        
    } catch (error) {
        console.error('[CLEAR] Error:', error);
        res.status(500).json({ error: 'Failed to clear data' });
    }
});

// ============= WEBSOCKET HANDLING =============

io.on('connection', (socket) => {
    console.log('[WS] New client connected');
    
    socket.on('joinSession', ({ userName }) => {
        socket.userName = userName;
        activeUsers.add(userName);
        
        console.log(`[WS] ${userName} joined`);
        
        // Notify all users
        io.emit('userJoined', { userName });
        io.emit('activeUsers', Array.from(activeUsers));
    });
    
    socket.on('disconnect', () => {
        if (socket.userName) {
            activeUsers.delete(socket.userName);
            console.log(`[WS] ${socket.userName} disconnected`);
            io.emit('activeUsers', Array.from(activeUsers));
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
        
        const partNumber = record['Product Number'] || record['Part Number'];
        const partDescription = record['PartDescription'] || '';
        
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
            const variantDemand = {};
            variants.forEach(variant => {
                const variantRecords = grouped[dfuCode].records.filter(r => {
                    const partNumber = r['Product Number'] || r['Part Number'];
                    return partNumber === variant;
                });
                
                const totalDemand = variantRecords.reduce((sum, r) => {
                    const demand = parseFloat(r['weekly fcst'] || 0);
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
    [SERVER] Port: ${PORT}
    [SERVER] Team Session ID: ${TEAM_SESSION_ID}
    `);
});