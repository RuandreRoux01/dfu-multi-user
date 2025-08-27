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
                variantCycleData: null,
                hasVariantCycleData: false
            });
        }
        
        console.log('[DB] Ready!');
    } catch (error) {
        console.error('[DB] Init error:', error);
    }
}

const storage = multer.memoryStorage();
const upload = multer({ storage, limits: { fileSize: 50 * 1024 * 1024 }});

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

async function saveTransfer(dfuCode, transferData, userName) {
    const transferDoc = {
        sessionId: TEAM_SESSION_ID,
        dfuCode,
        type: transferData.type || 'individual',
        completedBy: userName,
        completedAt: new Date()
    };
    
    const existingTransfer = await db.collection('transfers').findOne({ 
        sessionId: TEAM_SESSION_ID, 
        dfuCode 
    });
    
    if (!existingTransfer) {
        const session = await db.collection('sessions').findOne({ _id: TEAM_SESSION_ID });
        if (session && session.rawData) {
            const dfuOriginalRecords = session.rawData.filter(r => r['DFU'] === dfuCode);
            transferDoc.originalData = dfuOriginalRecords;
            console.log(`[TRANSFER] Storing original data for DFU ${dfuCode}:`, dfuOriginalRecords.length, 'records');
        }
    } else {
        transferDoc.originalData = existingTransfer.originalData;
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
}

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
        if (!req.file) return res.status(400).json({ error: 'No file' });
        
        const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
        const data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        
        await db.collection('transfers').deleteMany({ sessionId: TEAM_SESSION_ID });
        
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            { 
                $set: { 
                    rawData: data,
                    dataUploaded: true,
                    uploadedAt: new Date(),
                    uploadedBy: req.body.userName
                }
            }
        );
        
        res.json({ success: true, rowCount: data.length });
        io.emit('dataUploaded', { uploadedBy: req.body.userName });
        
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/upload-cycle', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'No file' });
        
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

app.get('/api/session/data', async (req, res) => {
    try {
        const session = await db.collection('sessions').findOne({ _id: TEAM_SESSION_ID });
        const completedTransfers = await getCompletedTransfers();
        
        let multiVariantDFUs = {};
        if (session?.rawData?.length > 0) {
            multiVariantDFUs = processMultiVariantDFUs(session.rawData);
        }
        
        res.json({
            session: {
                name: 'Team Session',
                dataUploaded: session?.dataUploaded || false,
                hasVariantCycleData: session?.hasVariantCycleData || false
            },
            multiVariantDFUs,
            rawData: session?.rawData || [],
            completedTransfers,
            variantCycleData: session?.variantCycleData || {},
            hasVariantCycleData: session?.hasVariantCycleData || false
        });
        
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/updateData', async (req, res) => {
    try {
        const { rawData, completedTransfers, transfer } = req.body;
        
        console.log(`[UPDATE] Processing transfer for DFU ${transfer?.dfuCode}`);
        
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
        
        if (recordsToAdd.length === 0) {
            const dfuRecords = session.rawData.filter(r => r['DFU'] === dfuCode);
            
            if (dfuRecords.length > 0) {
                const uniqueCombos = new Map();
                dfuRecords.forEach(r => {
                    const key = `${r['Week Number']}_${r['Source Location']}`;
                    if (!uniqueCombos.has(key)) {
                        uniqueCombos.set(key, r);
                    }
                });
                
                uniqueCombos.forEach((templateRecord) => {
                    const newRecord = { ...templateRecord };
                    newRecord['Product Number'] = variantCode.trim();
                    newRecord['weekly fcst'] = 0;
                    newRecord['PartDescription'] = 'Manually added variant';
                    newRecord['Transfer History'] = `Manually added by ${userName} on ${new Date().toLocaleString()}`;
                    recordsToAdd.push(newRecord);
                });
            } else {
                const sampleRecord = session.rawData[0];
                if (sampleRecord) {
                    const newRecord = { ...sampleRecord };
                    newRecord['DFU'] = dfuCode;
                    newRecord['Product Number'] = variantCode.trim();
                    newRecord['weekly fcst'] = 0;
                    newRecord['PartDescription'] = 'Manually added variant';
                    newRecord['Transfer History'] = `Manually added by ${userName} on ${new Date().toLocaleString()}`;
                    recordsToAdd.push(newRecord);
                }
            }
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
        
        let currentData = [...session.rawData];
        
        currentData = currentData.filter(r => r['DFU'] !== dfuCode);
        currentData = [...currentData, ...transferToUndo.originalData];
        
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            { 
                $set: { 
                    rawData: currentData,
                    lastModified: new Date()
                }
            }
        );
        
        await db.collection('transfers').deleteOne({ 
            sessionId: TEAM_SESSION_ID, 
            dfuCode 
        });
        
        const remainingTransfers = {};
        const remainingTransfersArray = await db.collection('transfers').find({ 
            sessionId: TEAM_SESSION_ID 
        }).toArray();
        
        remainingTransfersArray.forEach(t => {
            remainingTransfers[t.dfuCode] = t;
        });
        
        console.log(`[UNDO] Successfully restored DFU ${dfuCode}, ${remainingTransfersArray.length} transfers remain`);
        
        res.json({ 
            success: true, 
            remainingTransfers,
            message: `Transfer for DFU ${dfuCode} has been undone - all variants restored to original state`
        });
        
        io.emit('dataUpdated', { 
            rawData: currentData, 
            completedTransfers: remainingTransfers,
            message: `${userName} undid transfer for DFU ${dfuCode}`
        });
        
    } catch (error) {
        console.error('[UNDO ERROR]', error);
        res.status(500).json({ error: 'Failed to undo transfer' });
    }
});

app.post('/api/clear', async (req, res) => {
    try {
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            { 
                $set: { 
                    rawData: null,
                    dataUploaded: false,
                    variantCycleData: null,
                    hasVariantCycleData: false
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

const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
    console.log(`[SERVER] Running on port ${PORT}`);
});