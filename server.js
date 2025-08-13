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

app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// MongoDB connection
let db;
const mongoUri = process.env.MONGODB_URI;

MongoClient.connect(mongoUri)
    .then(client => {
        console.log('✅ Connected to MongoDB Atlas!');
        db = client.db('dfu_transfer_db');
        initializeCollections();
    })
    .catch(error => {
        console.error('❌ MongoDB connection error:', error);
        console.log('Please check your connection string in .env file');
    });

// Initialize database collections
async function initializeCollections() {
    try {
        await db.createCollection('sessions');
        await db.createCollection('transfers');
        console.log('✅ Database collections ready!');
    } catch (error) {
        console.log('Collections might already exist, that\'s okay!');
    }
}

// File upload setup
const storage = multer.memoryStorage();
const upload = multer({ 
    storage: storage,
    limits: { fileSize: 10 * 1024 * 1024 } // 10MB limit
});

// Store active sessions in memory
const activeSessions = new Map();

// ============= REST API ENDPOINTS =============

// Health check endpoint
app.get('/api/health', (req, res) => {
    res.json({ status: 'Server is running!', database: db ? 'Connected' : 'Not connected' });
});

// Create or join a session
app.post('/api/session/join', async (req, res) => {
    try {
        const { sessionName, userName } = req.body;
        
        if (!sessionName || !userName) {
            return res.status(400).json({ error: 'Session name and user name are required' });
        }
        
        // Find or create session by name
        let session = await db.collection('sessions').findOne({ name: sessionName });
        
        if (!session) {
            // Create new session with a simple string ID
            const sessionId = `session_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
            session = {
                _id: sessionId,  // Use string ID instead of ObjectId
                name: sessionName,
                createdAt: new Date(),
                users: [],
                dataUploaded: false,
                rawData: null,
                status: 'active'
            };
            await db.collection('sessions').insertOne(session);
            console.log(`📁 New session created: ${sessionName} with ID: ${sessionId}`);
        }
        
        // Add user to session
        const userId = `user_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
        await db.collection('sessions').updateOne(
            { _id: session._id },
            { 
                $push: { users: { id: userId, name: userName, joinedAt: new Date() } }
            }
        );
        
        console.log(`👤 ${userName} joined session: ${sessionName}`);
        
        res.json({ 
            success: true,
            sessionId: session._id,
            sessionName: session.name,
            userId: userId
        });
    } catch (error) {
        console.error('Error joining session:', error);
        res.status(500).json({ error: 'Failed to join session: ' + error.message });
    }
});

// Upload Excel file
app.post('/api/upload/:sessionId', upload.single('file'), async (req, res) => {
    try {
        const { sessionId } = req.params;
        const file = req.file;
        
        if (!file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }
        
        console.log(`📤 Processing uploaded file for session: ${sessionId}`);
        
        // Parse Excel file
        const workbook = XLSX.read(file.buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet);
        
        console.log(`📊 Parsed ${data.length} rows from Excel file`);
        
        // Store raw data in session - use sessionId as string
        const updateResult = await db.collection('sessions').updateOne(
            { _id: sessionId },  // sessionId is already a string
            { 
                $set: { 
                    rawData: data,
                    dataUploaded: true,
                    uploadedAt: new Date()
                }
            }
        );
        
        console.log(`📝 Update result: ${updateResult.modifiedCount} document(s) modified`);
        
        // Process multi-variant DFUs
        const multiVariantDFUs = processMultiVariantDFUs(data);
        console.log(`🔍 Found ${Object.keys(multiVariantDFUs).length} multi-variant DFUs`);
        
        res.json({ 
            success: true, 
            rowCount: data.length,
            dfuCount: Object.keys(multiVariantDFUs).length 
        });
        
        // Notify all connected clients
        io.to(sessionId).emit('dataUploaded', { dfuCount: Object.keys(multiVariantDFUs).length });
        
    } catch (error) {
        console.error('Error uploading file:', error);
        res.status(500).json({ error: 'Failed to process file: ' + error.message });
    }
});

// Get session data
app.get('/api/session/:sessionId/data', async (req, res) => {
    try {
        const { sessionId } = req.params;
        
        console.log(`📊 Getting data for session: ${sessionId}`);
        
        const session = await db.collection('sessions').findOne({ 
            _id: sessionId  // Use string ID directly
        });
        
        if (!session) {
            console.log(`❌ Session not found: ${sessionId}`);
            return res.status(404).json({ error: 'Session not found' });
        }
        
        console.log(`✅ Found session: ${session.name}`);
        
        // Process multi-variant DFUs if data exists
        let multiVariantDFUs = {};
        if (session.rawData && session.rawData.length > 0) {
            multiVariantDFUs = processMultiVariantDFUs(session.rawData);
            console.log(`🔍 Processed ${Object.keys(multiVariantDFUs).length} multi-variant DFUs`);
        }
        
        // Get transfers for this session
        const transfers = await db.collection('transfers').find({ 
            sessionId: sessionId 
        }).toArray();
        
        res.json({
            session: {
                name: session.name,
                dataUploaded: session.dataUploaded || false,
                userCount: session.users ? session.users.length : 0
            },
            multiVariantDFUs,
            transfers,
            rawData: session.rawData || []
        });
        
    } catch (error) {
        console.error('Error getting session data:', error);
        res.status(500).json({ error: 'Failed to get session data: ' + error.message });
    }
});

// Save transfer configuration
app.post('/api/session/:sessionId/transfer', async (req, res) => {
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
        
        console.log(`💾 Transfer saved for DFU ${dfuCode} by ${userName}`);
        
        res.json({ success: true });
        
        // Notify other users
        io.to(sessionId).emit('transferUpdated', { dfuCode, transferConfig, userName });
        
    } catch (error) {
        console.error('Error saving transfer:', error);
        res.status(500).json({ error: 'Failed to save transfer: ' + error.message });
    }
});

// Execute transfers and generate file
app.post('/api/session/:sessionId/generate', async (req, res) => {
    try {
        const { sessionId } = req.params;
        
        console.log(`📦 Generating file for session: ${sessionId}`);
        
        // Get session data
        const session = await db.collection('sessions').findOne({ 
            _id: sessionId  // String ID
        });
        
        if (!session || !session.rawData) {
            return res.status(400).json({ error: 'No data available for this session' });
        }
        
        // Get all transfers for this session
        const transfers = await db.collection('transfers').find({ sessionId }).toArray();
        console.log(`🔄 Found ${transfers.length} transfers to apply`);
        
        // Start with the original raw data
        let processedData = [...session.rawData];
        
        // Apply each transfer sequentially
        transfers.forEach((transfer, index) => {
            console.log(`Applying transfer ${index + 1}/${transfers.length}: DFU ${transfer.dfuCode}`);
            processedData = applyTransfer(processedData, transfer);
        });
        
        // Log summary of changes
        const originalDFUs = {};
        const processedDFUs = {};
        
        session.rawData.forEach(record => {
            const dfu = record['DFU'];
            const part = record['Product Number'];
            if (!originalDFUs[dfu]) originalDFUs[dfu] = new Set();
            originalDFUs[dfu].add(part);
        });
        
        processedData.forEach(record => {
            const dfu = record['DFU'];
            const part = record['Product Number'];
            if (!processedDFUs[dfu]) processedDFUs[dfu] = new Set();
            processedDFUs[dfu].add(part);
        });
        
        console.log('Transfer Summary:');
        Object.keys(originalDFUs).forEach(dfu => {
            const originalCount = originalDFUs[dfu].size;
            const processedCount = processedDFUs[dfu] ? processedDFUs[dfu].size : 0;
            if (originalCount !== processedCount) {
                console.log(`  DFU ${dfu}: ${originalCount} variants -> ${processedCount} variants`);
            }
        });
        
        // Create Excel file
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(processedData);
        XLSX.utils.book_append_sheet(wb, ws, 'Updated Demand');
        
        // Generate buffer
        const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
        
        console.log(`📦 Generated file with ${transfers.length} transfers applied`);
        
        // Send file
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=DFU_Transfers_${Date.now()}.xlsx`);
        res.send(buffer);
        
    } catch (error) {
        console.error('Error generating file:', error);
        res.status(500).json({ error: 'Failed to generate file: ' + error.message });
    }
});
// Add this new endpoint to clear session data
app.post('/api/session/:sessionId/clear', async (req, res) => {
    try {
        const { sessionId } = req.params;
        
        // Clear all transfers for this session
        await db.collection('transfers').deleteMany({ sessionId });
        
        // Clear raw data from session
        await db.collection('sessions').updateOne(
            { _id: sessionId },
            { 
                $set: { 
                    rawData: null,
                    dataUploaded: false
                }
            }
        );
        
        console.log(`🗑️ Cleared data for session: ${sessionId}`);
        res.json({ success: true });
        
    } catch (error) {
        console.error('Error clearing session:', error);
        res.status(500).json({ error: 'Failed to clear session' });
    }
});

// ============= WEBSOCKET HANDLING =============

io.on('connection', (socket) => {
    console.log('🔌 New client connected');
    
    socket.on('joinSession', ({ sessionId, userName }) => {
        socket.join(sessionId);
        socket.sessionId = sessionId;
        socket.userName = userName;
        
        // Track active users
        if (!activeSessions.has(sessionId)) {
            activeSessions.set(sessionId, new Set());
        }
        activeSessions.get(sessionId).add(userName);
        
        console.log(`👥 ${userName} joined session via WebSocket`);
        
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
                io.to(socket.sessionId).emit('activeUsers', Array.from(sessionUsers));
                console.log(`👋 ${socket.userName} disconnected`);
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

function applyTransfer(data, transfer) {
    console.log(`Applying transfer for DFU ${transfer.dfuCode}, type: ${transfer.type}`);
    
    return data.map(record => {
        const recordDFU = record['DFU'];
        const recordPartNumber = record['Product Number'] || record['Part Number'] || record['Part Code'];
        
        if (recordDFU === transfer.dfuCode) {
            if (transfer.type === 'bulk' && recordPartNumber !== transfer.targetVariant) {
                // For bulk transfer, change all non-target variants to target
                const originalPart = recordPartNumber;
                const originalDemand = record['weekly fcst'] || 0;
                
                record['Product Number'] = transfer.targetVariant;
                record['Transfer History'] = `Bulk transferred from ${originalPart} to ${transfer.targetVariant} (${originalDemand} units) by ${transfer.updatedBy}`;
                
                console.log(`Bulk transfer: ${originalPart} -> ${transfer.targetVariant}`);
                
            } else if (transfer.type === 'individual' && transfer.transfers) {
                // For individual transfers, check if this variant has a transfer mapping
                const targetVariant = transfer.transfers[recordPartNumber];
                if (targetVariant && targetVariant !== recordPartNumber) {
                    const originalPart = recordPartNumber;
                    const originalDemand = record['weekly fcst'] || 0;
                    
                    record['Product Number'] = targetVariant;
                    record['Transfer History'] = `Transferred from ${originalPart} to ${targetVariant} (${originalDemand} units) by ${transfer.updatedBy}`;
                    
                    console.log(`Individual transfer: ${originalPart} -> ${targetVariant}`);
                }
            }
        }
        return record;
    });
}

// ============= START SERVER =============

const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
    console.log(`
    🚀 Server is running!
    📍 Local: http://localhost:3000
    📍 Network: http://[your-computer-ip]:3000
    
    Waiting for MongoDB connection...
    `);
});
