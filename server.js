const express = require('express');
const http = require('http');
const socketIo = require('socket.io');
const cors = require('cors');
const multer = require('multer');
const { MongoClient } = require('mongodb');
const XLSX = require('xlsx');
require('dotenv').config();

const app = express();
const server = http.createServer(app);
const io = socketIo(server, {
    cors: { origin: "*", methods: ["GET", "POST"] }
});

app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.static('public'));

let db, isDbConnected = false;
const TEAM_SESSION_ID = 'TEAM_DFU_TRANSFER_SESSION';
const activeUsers = new Set();

// MongoDB connection
MongoClient.connect(process.env.MONGODB_URI)
    .then(client => {
        console.log('[DB] Connected!');
        db = client.db('dfu_transfer_db');
        isDbConnected = true;
        initializeDatabase();
    })
    .catch(err => console.error('[DB] Error:', err));

async function initializeDatabase() {
    try {
        await db.createCollection('sessions').catch(() => {});
        await db.createCollection('transfers').catch(() => {});
        
        let session = await db.collection('sessions').findOne({ _id: TEAM_SESSION_ID });
        
        if (!session) {
            await db.collection('sessions').insertOne({
                _id: TEAM_SESSION_ID,
                name: 'Team Session',
                createdAt: new Date(),
                dataUploaded: false,
                rawData: null,
                originalRawData: null,
                variantCycleData: null,
                hasVariantCycleData: false,
                supplyChainData: {
                    stockData: {},
                    openSupplyData: {},
                    transitData: {}
                }
            });
        }
        
        console.log('[DB] Ready!');
    } catch (err) {
        console.error('[DB INIT]', err);
    }
}

const upload = multer({ storage: multer.memoryStorage() });

// WebSocket
io.on('connection', (socket) => {
    console.log('[WS] Client connected');
    
    socket.on('joinSession', ({ userName }) => {
        socket.userName = userName;
        activeUsers.add(userName);
        io.emit('userJoined', { userName });
        io.emit('activeUsers', Array.from(activeUsers));
    });
    
    socket.on('disconnect', () => {
        if (socket.userName) {
            activeUsers.delete(socket.userName);
            io.emit('userLeft', { userName: socket.userName });
            io.emit('activeUsers', Array.from(activeUsers));
        }
    });
});

// Health check
app.get('/api/health', (req, res) => {
    res.json({ 
        status: 'OK', 
        db: isDbConnected,
        session: TEAM_SESSION_ID 
    });
});

// Join session
app.post('/api/session/join', async (req, res) => {
    try {
        const { userName } = req.body;
        if (!userName) return res.status(400).json({ error: 'Name required' });
        
        const userId = `user_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
        
        res.json({ 
            success: true,
            sessionId: TEAM_SESSION_ID,
            userId,
            userName
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Get session data
app.get('/api/session/data', async (req, res) => {
    try {
        const session = await db.collection('sessions').findOne({ _id: TEAM_SESSION_ID });
        const completedTransfers = await getCompletedTransfers();
        
        let multiVariantDFUs = {};
        if (session?.rawData?.length > 0) {
            multiVariantDFUs = processMultiVariantDFUs(session.rawData);
        }
        
        res.json({
            success: true,
            session: {
                name: 'Team Session',
                dataUploaded: session?.dataUploaded || false,
                hasVariantCycleData: session?.hasVariantCycleData || false
            },
            // Return data at root level for loadTeamData() compatibility
            rawData: session?.rawData || [],
            originalRawData: session?.originalRawData || [],
            multiVariantDFUs,
            completedTransfers,
            variantCycleData: session?.variantCycleData || {},
            hasVariantCycleData: session?.hasVariantCycleData || false,
            supplyChainData: session?.supplyChainData || {
                stockData: {},
                openSupplyData: {},
                transitData: {}
            }
        });
    } catch (error) {
        console.error('[SESSION DATA] Error:', error);
        res.status(500).json({ error: error.message });
    }
});

// Helper function to get completed transfers
async function getCompletedTransfers() {
    try {
        const transfers = await db.collection('transfers').find({ 
            sessionId: TEAM_SESSION_ID 
        }).toArray();
        
        const result = {};
        transfers.forEach(t => {
            if (!t.isPlaceholder) {
                result[t.dfuCode] = {
                    type: t.type || 'individual',
                    timestamp: t.timestamp,
                    completedBy: t.completedBy
                };
            }
        });
        
        return result;
    } catch (error) {
        console.error('[COMPLETED TRANSFERS]', error);
        return {};
    }
}

// Helper function to process multi-variant DFUs
function processMultiVariantDFUs(rawData) {
    const grouped = {};
    
    rawData.forEach(record => {
        const dfuCode = record['DFU'];
        const partNumber = record['Product Number'] || record['Part Number'];
        const partDescription = record['PartDescription'] || '';
        
        if (!grouped[dfuCode]) {
            grouped[dfuCode] = {
                records: [],
                variants: new Set(),
                partDescriptions: {}
            };
        }
        
        grouped[dfuCode].records.push(record);
        if (partNumber) {
            grouped[dfuCode].variants.add(partNumber);
            grouped[dfuCode].partDescriptions[partNumber] = partDescription;
        }
    });
    
    const allDFUs = {};
    Object.keys(grouped).forEach(dfuCode => {
        const variants = Array.from(grouped[dfuCode].variants);
        const variantDemand = {};
        variants.forEach(variant => {
            const variantRecords = grouped[dfuCode].records.filter(r => 
                (r['Product Number'] || r['Part Number']) === variant
            );
            
            const totalDemand = variantRecords.reduce((sum, r) => 
                sum + parseFloat(r['weekly fcst'] || 0), 0
            );
            
            variantDemand[variant] = {
                totalDemand,
                recordCount: variantRecords.length,
                description: grouped[dfuCode].partDescriptions[variant] || ''
            };
        });
        
        allDFUs[dfuCode] = {
            variants,
            recordCount: grouped[dfuCode].records.length,
            variantDemand,
            isSingleVariant: variants.length === 1
        };
    });
    
    return allDFUs;
}

// Upload demand data file
app.post('/api/upload', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
        
        const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet);
        
        console.log(`[UPLOAD] Processed ${data.length} rows`);
        
        // Store original data backup for undo functionality
        const originalData = JSON.parse(JSON.stringify(data));
        
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            { 
                $set: { 
                    rawData: data,
                    originalRawData: originalData,
                    dataUploaded: true,
                    lastModified: new Date(),
                    completedTransfers: {}
                }
            },
            { upsert: true }
        );
        
        io.emit('dataUploaded', { 
            rowCount: data.length,
            uploadedBy: req.body.userName || 'Unknown'
        });
        
        res.json({ 
            success: true,
            rowCount: data.length
        });
    } catch (error) {
        console.error('[UPLOAD] Error:', error);
        res.status(500).json({ error: error.message });
    }
});

// Upload supply chain data
app.post('/api/upload-supply-chain', async (req, res) => {
    try {
        const { type, data, userName } = req.body;
        
        if (!type || !data) {
            return res.status(400).json({ error: 'Missing type or data' });
        }
        
        const validTypes = ['stock', 'openSupply', 'transit'];
        if (!validTypes.includes(type)) {
            return res.status(400).json({ error: 'Invalid type' });
        }
        
        console.log(`[SUPPLY CHAIN] Uploading ${type} data: ${Object.keys(data).length} items`);
        
        const session = await db.collection('sessions').findOne({ _id: TEAM_SESSION_ID });
        const supplyChainData = session?.supplyChainData || {
            stockData: {},
            openSupplyData: {},
            transitData: {}
        };
        
        // Update the specific type
        if (type === 'stock') {
            supplyChainData.stockData = data;
        } else if (type === 'openSupply') {
            supplyChainData.openSupplyData = data;
        } else if (type === 'transit') {
            supplyChainData.transitData = data;
        }
        
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            { 
                $set: { 
                    supplyChainData,
                    lastModified: new Date()
                }
            },
            { upsert: true }
        );
        
        // Emit socket event to all users
        io.emit('supplyChainDataUpdated', { 
            type,
            data,
            uploadedBy: userName || 'Unknown'
        });
        
        res.json({ 
            success: true,
            message: `${type} data uploaded successfully`,
            itemCount: Object.keys(data).length
        });
    } catch (error) {
        console.error('[SUPPLY CHAIN] Error:', error);
        res.status(500).json({ error: error.message });
    }
});

// Update session data (after transfers)
app.post('/api/update', async (req, res) => {
    try {
        const { rawData, completedTransfers, transfer } = req.body;
        
        // Check if we need to store original data for the first transfer
        if (transfer && transfer.dfuCode) {
            const existingTransfer = await db.collection('transfers').findOne({ 
                sessionId: TEAM_SESSION_ID, 
                dfuCode: transfer.dfuCode 
            });
            
            if (!existingTransfer) {
                const currentSession = await db.collection('sessions').findOne({ _id: TEAM_SESSION_ID });
                if (currentSession && currentSession.originalRawData) {
                    const dfuOriginalRecords = currentSession.originalRawData.filter(r => r['DFU'] === transfer.dfuCode);
                    console.log(`[UPDATE] Found ${dfuOriginalRecords.length} original records for DFU ${transfer.dfuCode} in backup`);
                    
                    await db.collection('transfers').insertOne({
                        sessionId: TEAM_SESSION_ID,
                        dfuCode: transfer.dfuCode,
                        type: 'placeholder',
                        originalData: JSON.parse(JSON.stringify(dfuOriginalRecords)),
                        createdAt: new Date(),
                        isPlaceholder: true
                    });
                }
            }
        }
        
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            { $set: { rawData, lastModified: new Date() }}
        );
        
        if (completedTransfers && Object.keys(completedTransfers).length > 0) {
            for (const [dfuCode, transferData] of Object.entries(completedTransfers)) {
                await saveTransfer(dfuCode, transferData, transfer?.completedBy || 'Unknown');
            }
        }
        
        res.json({ success: true });
        io.emit('dataUpdated', { 
            dfuCode: transfer?.dfuCode,
            updatedBy: transfer?.completedBy 
        });
        
    } catch (error) {
        console.error('[UPDATE] Error:', error);
        res.status(500).json({ error: error.message });
    }
});

// Save transfer record
async function saveTransfer(dfuCode, transferData, userName) {
    const transferDoc = {
        sessionId: TEAM_SESSION_ID,
        dfuCode,
        completedBy: userName,
        timestamp: new Date()
    };
    
    const existingTransfer = await db.collection('transfers').findOne({ 
        sessionId: TEAM_SESSION_ID, 
        dfuCode 
    });
    
    if (!existingTransfer || !existingTransfer.originalData) {
        const currentSession = await db.collection('sessions').findOne({ _id: TEAM_SESSION_ID });
        if (currentSession && currentSession.originalRawData) {
            const dfuOriginalRecords = currentSession.originalRawData.filter(r => r['DFU'] === dfuCode);
            transferDoc.originalData = JSON.parse(JSON.stringify(dfuOriginalRecords));
            console.log(`[TRANSFER] Using original data backup: ${dfuOriginalRecords.length} records for DFU ${dfuCode}`);
        }
    } else {
        transferDoc.originalData = existingTransfer.originalData;
        console.log(`[TRANSFER] Preserving existing original data for DFU ${dfuCode}: ${existingTransfer.originalData?.length || 0} records`);
    }
    
    if (transferData.bulkTransfer) {
        transferDoc.bulkTransfer = transferData.bulkTransfer;
    }
    if (transferData.granularTransfers) {
        transferDoc.granularTransfers = transferData.granularTransfers;
    }
    if (transferData.transfers) {
        transferDoc.transfers = transferData.transfers;
    }
    if (transferData.targetVariant) {
        transferDoc.targetVariant = transferData.targetVariant;
    }
    
    await db.collection('transfers').replaceOne(
        { sessionId: TEAM_SESSION_ID, dfuCode },
        transferDoc,
        { upsert: true }
    );
    
    console.log(`[TRANSFER] Saved transfer record for DFU ${dfuCode} with ${transferDoc.originalData?.length || 0} original records`);
}

// Undo transfer
app.post('/api/undoTransfer', async (req, res) => {
    try {
        const { dfuCode, userName } = req.body;
        
        console.log(`[UNDO] ${userName} undoing transfer for DFU ${dfuCode}`);
        
        const session = await db.collection('sessions').findOne({ _id: TEAM_SESSION_ID });
        if (!session) {
            return res.status(400).json({ error: 'No session data found' });
        }
        
        const transferToUndo = await db.collection('transfers').findOne({ 
            sessionId: TEAM_SESSION_ID, 
            dfuCode 
        });
        
        if (!transferToUndo || !transferToUndo.originalData) {
            return res.status(400).json({ error: 'No original data found for this transfer' });
        }
        
        console.log(`[UNDO] Found transfer to undo with ${transferToUndo.originalData.length} original records`);
        
        let currentData = [...session.rawData];
        
        const beforeCount = currentData.length;
        currentData = currentData.filter(r => r['DFU'] !== dfuCode);
        console.log(`[UNDO] Removed ${beforeCount - currentData.length} records for DFU ${dfuCode}`);
        
        const restoredRecords = transferToUndo.originalData.map(record => ({...record}));
        currentData = [...currentData, ...restoredRecords];
        console.log(`[UNDO] Restored ${restoredRecords.length} original records`);
        
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            { $set: { rawData: currentData, lastModified: new Date() }}
        );
        
        await db.collection('transfers').deleteOne({ 
            sessionId: TEAM_SESSION_ID, 
            dfuCode 
        });
        
        res.json({ 
            success: true, 
            message: `Transfer undone for DFU ${dfuCode}`,
            restoredRecords: restoredRecords.length
        });
        
        io.emit('transferUndone', { 
            dfuCode,
            undoneBy: userName 
        });
        
    } catch (error) {
        console.error('[UNDO] Error:', error);
        res.status(500).json({ error: error.message });
    }
});

// Add variant manually
app.post('/api/addVariant', async (req, res) => {
    try {
        const { dfuCode, variantCode, newRecords, userName } = req.body;
        
        console.log(`[ADD VARIANT] Adding variant ${variantCode} to DFU ${dfuCode}`);
        console.log(`[ADD VARIANT] Received ${newRecords ? newRecords.length : 0} new records`);
        
        if (!dfuCode || !variantCode) {
            return res.status(400).json({ error: 'Missing required fields' });
        }
        
        const recordsToAdd = Array.isArray(newRecords) ? newRecords : [];
        
        const session = await db.collection('sessions').findOne({ _id: TEAM_SESSION_ID });
        if (!session || !session.rawData) {
            return res.status(400).json({ error: 'No session data found' });
        }
        
        const updatedRawData = [...session.rawData, ...recordsToAdd];
        
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            { 
                $set: { 
                    rawData: updatedRawData,
                    lastModified: new Date()
                }
            }
        );
        
        console.log(`[ADD VARIANT] Successfully added ${recordsToAdd.length} records`);
        
        res.json({ 
            success: true, 
            message: `Variant ${variantCode} added to DFU ${dfuCode}`,
            recordsAdded: recordsToAdd.length
        });
        
        io.emit('variantAdded', { 
            dfuCode, 
            variantCode, 
            addedBy: userName 
        });
        
    } catch (error) {
        console.error('[ADD VARIANT] Error:', error);
        res.status(500).json({ error: error.message });
    }
});

// Clear all data
app.post('/api/clear', async (req, res) => {
    try {
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            { 
                $set: { 
                    rawData: null,
                    originalRawData: null,
                    dataUploaded: false,
                    variantCycleData: null,
                    hasVariantCycleData: false,
                    supplyChainData: {
                        stockData: {},
                        openSupplyData: {},
                        transitData: {}
                    }
                }
            }
        );
        
        await db.collection('transfers').deleteMany({ sessionId: TEAM_SESSION_ID });
        
        res.json({ success: true });
        io.emit('dataCleared', { clearedBy: req.body.userName });
        
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Export data
app.post('/api/export', async (req, res) => {
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
        res.status(500).json({ error: error.message });
    }
});

const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
    console.log(`[SERVER] Running on port ${PORT}`);
});