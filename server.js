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
        console.log('[DB] Connected to MongoDB!');
        db = client.db('dfu_transfer_db');
        isDbConnected = true;
        initializeDatabase();
    })
    .catch(err => console.error('[DB] Connection Error:', err));

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
                users: [],
                dataUploaded: false,
                rawData: null,
                originalRawData: null,
                variantCycleData: null,
                hasVariantCycleData: false,
                completedTransfers: {},
                supplyChainData: {
                    stockData: {},
                    openSupplyData: {},
                    transitData: {}
                },
                status: 'active'
            });
            console.log('[DB] Session initialized');
        }
        
        console.log('[DB] Database ready!');
    } catch (error) {
        console.error('[DB] Initialization error:', error);
    }
}

const storage = multer.memoryStorage();
const upload = multer({ storage, limits: { fileSize: 50 * 1024 * 1024 }});

// Helper functions
async function getCompletedTransfers() {
    const transfers = await db.collection('transfers').find({ 
        sessionId: TEAM_SESSION_ID 
    }).toArray();
    
    const completed = {};
    transfers.forEach(t => {
        completed[t.dfuCode] = t;
    });
    return completed;
}

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

// API Routes
app.get('/api/health', (req, res) => {
    res.json({ 
        status: 'OK', 
        db: isDbConnected,
        session: TEAM_SESSION_ID 
    });
});

app.post('/api/session/join', async (req, res) => {
    try {
        const { userName } = req.body;
        if (!userName) return res.status(400).json({ error: 'Name required' });
        
        const userId = `user_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
        
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            { $addToSet: { users: userName } }
        );
        
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

app.post('/api/upload', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
        
        const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
        const data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        
        await db.collection('transfers').deleteMany({ sessionId: TEAM_SESSION_ID });
        
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            { 
                $set: { 
                    rawData: data,
                    originalRawData: JSON.parse(JSON.stringify(data)),
                    dataUploaded: true,
                    uploadedAt: new Date(),
                    uploadedBy: req.body.userName,
                    completedTransfers: {}
                }
            }
        );
        
        console.log(`[UPLOAD] Stored ${data.length} records with backup copy`);
        
        res.json({ success: true, rowCount: data.length });
        io.emit('dataUploaded', { uploadedBy: req.body.userName, cycleDataReset: true });
        
    } catch (error) {
        console.error('[UPLOAD] Error:', error);
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/upload-cycle', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
        
        const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
        const cycleData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        
        const processed = {};
        cycleData.forEach(rec => {
            const dfu = rec['DFU'] || rec['dfu'];
            const part = rec['Part Code'] || rec['Product Number'];
            if (dfu && part) {
                if (!processed[dfu]) processed[dfu] = {};
                processed[dfu][part] = {
                    sos: rec['SOS'] || 'N/A',
                    eos: rec['EOS'] || 'N/A',
                    comments: rec['Comments'] || ''
                };
            }
        });
        
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            { 
                $set: { 
                    variantCycleData: processed,
                    hasVariantCycleData: true,
                    cycleUploadedAt: new Date()
                }
            }
        );
        
        res.json({ success: true, dfuCount: Object.keys(processed).length });
        io.emit('cycleDataUploaded', { uploadedBy: req.body.userName });
        
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/upload-supply-chain', async (req, res) => {
    try {
        const { type, data, userName } = req.body;
        
        const updateField = type === 'stock' ? 'supplyChainData.stockData' :
                           type === 'openSupply' ? 'supplyChainData.openSupplyData' :
                           type === 'transit' ? 'supplyChainData.transitData' : null;
        
        if (!updateField) {
            return res.status(400).json({ error: 'Invalid supply chain type' });
        }
        
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            { $set: { [updateField]: data } }
        );
        
        console.log(`[SUPPLY CHAIN] ${type} data saved:`, Object.keys(data).length, 'items');
        
        res.json({ success: true, itemCount: Object.keys(data).length });
        io.emit('supplyChainDataUpdated', { type, uploadedBy: userName, data });
        
    } catch (error) {
        console.error('[SUPPLY CHAIN] Error:', error);
        res.status(500).json({ error: error.message });
    }
});

app.get('/api/session/data', async (req, res) => {
    try {
        const session = await db.collection('sessions').findOne({ _id: TEAM_SESSION_ID });
        
        if (!session) {
            return res.status(404).json({ error: 'Session not found' });
        }
        
        const completedTransfers = session.completedTransfers || {};
        
        let multiVariantDFUs = {};
        if (session.rawData && session.rawData.length > 0) {
            multiVariantDFUs = processMultiVariantDFUs(session.rawData);
        }
        
        res.json({
            success: true,
            session: {
                name: 'Team Session',
                dataUploaded: session.dataUploaded || false,
                hasVariantCycleData: session.hasVariantCycleData || false,
                rawData: session.rawData || [],
                variantCycleData: session.variantCycleData || {},
                completedTransfers: completedTransfers,
                supplyChainData: session.supplyChainData || {
                    stockData: {},
                    openSupplyData: {},
                    transitData: {}
                }
            },
            multiVariantDFUs,
            rawData: session.rawData || [],
            completedTransfers,
            variantCycleData: session.variantCycleData || {},
            hasVariantCycleData: session.hasVariantCycleData || false
        });
        
    } catch (error) {
        console.error('[SESSION DATA] Error:', error);
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/updateData', async (req, res) => {
    try {
        const { rawData, completedTransfers, transfer } = req.body;
        
        console.log(`[UPDATE] Processing transfer for DFU ${transfer?.dfuCode}`);
        
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            { 
                $set: { 
                    rawData: rawData,
                    completedTransfers: completedTransfers || {},
                    lastModified: new Date()
                }
            }
        );
        
        if (transfer && transfer.dfuCode) {
            await db.collection('transfers').updateOne(
                { sessionId: TEAM_SESSION_ID, dfuCode: transfer.dfuCode },
                { 
                    $set: {
                        sessionId: TEAM_SESSION_ID,
                        dfuCode: transfer.dfuCode,
                        type: transfer.type,
                        completedBy: transfer.completedBy,
                        completedAt: new Date()
                    }
                },
                { upsert: true }
            );
        }
        
        res.json({ success: true });
        
        if (transfer && transfer.dfuCode) {
            io.emit('transferExecuted', { 
                dfuCode: transfer.dfuCode, 
                executedBy: transfer.completedBy 
            });
        }
        
    } catch (error) {
        console.error('[UPDATE] Error:', error);
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/addVariant', async (req, res) => {
    try {
        const { dfuCode, variantCode, userName } = req.body;
        
        const session = await db.collection('sessions').findOne({ _id: TEAM_SESSION_ID });
        if (!session || !session.rawData) {
            return res.status(404).json({ error: 'No data found' });
        }
        
        const dfuRecords = session.rawData.filter(r => r['DFU'] === dfuCode);
        if (dfuRecords.length === 0) {
            return res.status(404).json({ error: 'DFU not found' });
        }
        
        const templateRecord = dfuRecords[0];
        const uniqueWeeks = new Set();
        dfuRecords.forEach(r => {
            const key = `${r['Week Number']}-${r['Source Location']}-${r['Calendar.week']}`;
            uniqueWeeks.add(key);
        });
        
        const newRecords = [];
        uniqueWeeks.forEach(key => {
            const [weekNum, sourceLoc, calWeek] = key.split('-');
            const newRecord = { ...templateRecord };
            newRecord['Product Number'] = variantCode;
            newRecord['weekly fcst'] = 0;
            newRecord['Week Number'] = weekNum;
            newRecord['Source Location'] = sourceLoc;
            newRecord['Calendar.week'] = calWeek;
            newRecord['Transfer History'] = 'PIPO [Manually added variant]';
            newRecords.push(newRecord);
        });
        
        session.rawData.push(...newRecords);
        
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            { $set: { rawData: session.rawData } }
        );
        
        res.json({ success: true, recordsAdded: newRecords.length });
        io.emit('variantAdded', { dfuCode, variantCode, addedBy: userName });
        
    } catch (error) {
        console.error('[ADD VARIANT] Error:', error);
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/undoTransfer', async (req, res) => {
    try {
        const { dfuCode, userName } = req.body;
        
        console.log(`[UNDO] ${userName} undoing transfer for DFU ${dfuCode}`);
        
        const session = await db.collection('sessions').findOne({ _id: TEAM_SESSION_ID });
        
        if (!session) {
            return res.status(404).json({ error: 'Session not found' });
        }
        
        if (!session.originalRawData || session.originalRawData.length === 0) {
            return res.status(404).json({ error: 'No original data found for this transfer' });
        }
        
        const originalDfuRecords = session.originalRawData.filter(record => 
            String(record['DFU']).trim() === String(dfuCode).trim()
        );
        
        if (originalDfuRecords.length === 0) {
            return res.status(404).json({ error: 'No original data found for this DFU' });
        }
        
        let currentRawData = session.rawData || [];
        
        currentRawData = currentRawData.filter(record => 
            String(record['DFU']).trim() !== String(dfuCode).trim()
        );
        
        const restoredRecords = originalDfuRecords.map(record => ({ ...record }));
        currentRawData.push(...restoredRecords);
        
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            {
                $set: {
                    rawData: currentRawData,
                    lastModified: new Date()
                }
            }
        );
        
        const completedTransfers = session.completedTransfers || {};
        delete completedTransfers[dfuCode];
        
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            {
                $set: {
                    completedTransfers: completedTransfers
                }
            }
        );
        
        await db.collection('transfers').deleteMany({ 
            sessionId: TEAM_SESSION_ID,
            dfuCode: dfuCode
        });
        
        console.log(`[UNDO] Successfully restored ${restoredRecords.length} original records for DFU ${dfuCode}`);
        
        io.emit('transferUndone', { 
            dfuCode, 
            undoneBy: userName,
            restoredRecords: restoredRecords.length
        });
        
        res.json({ 
            success: true, 
            message: 'Transfer undone successfully',
            restoredRecords: restoredRecords.length,
            remainingTransfers: completedTransfers
        });
        
    } catch (error) {
        console.error('[UNDO] Error:', error);
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/clear', async (req, res) => {
    try {
        console.log('[CLEAR] Clearing all data...');
        
        await db.collection('sessions').deleteOne({ _id: TEAM_SESSION_ID });
        await db.collection('transfers').deleteMany({ sessionId: TEAM_SESSION_ID });
        
        await db.collection('sessions').insertOne({
            _id: TEAM_SESSION_ID,
            name: 'Team Session',
            createdAt: new Date(),
            users: [],
            dataUploaded: false,
            rawData: null,
            originalRawData: null,
            variantCycleData: null,
            hasVariantCycleData: false,
            completedTransfers: {},
            supplyChainData: {
                stockData: {},
                openSupplyData: {},
                transitData: {}
            },
            status: 'active',
            lastModified: new Date()
        });
        
        console.log('[CLEAR] All data cleared, session reset');
        
        res.json({ success: true });
        io.emit('dataCleared', { clearedBy: req.body.userName });
        
    } catch (error) {
        console.error('[CLEAR] Error:', error);
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/export', async (req, res) => {
    try {
        const { rawData } = req.body;
        
        // Add OrderNumber as first column
        const dataWithOrderNumber = (rawData || []).map((row, index) => ({
            OrderNumber: index + 1,
            ...row
        }));
        
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(dataWithOrderNumber);
        XLSX.utils.book_append_sheet(wb, ws, 'Updated Demand');
        
        const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=DFU_Transfers_${Date.now()}.xlsx`);
        res.send(buffer);
        
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

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
            io.emit('activeUsers', Array.from(activeUsers));
        }
    });
});

const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
    console.log(`[SERVER] Running on port ${PORT}`);
});