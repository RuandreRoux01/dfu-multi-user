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
    
    // Check if this transfer already exists to preserve original data
    const existingTransfer = await db.collection('transfers').findOne({ 
        sessionId: TEAM_SESSION_ID, 
        dfuCode 
    });
    
    // Store original data ONLY if this is the first transfer for this DFU
    if (!existingTransfer) {
        console.log(`[TRANSFER] First transfer for DFU ${dfuCode} - need to get original data from BEFORE transfer`);
        
        // We need to get the original data from the session BEFORE this update
        // This should be handled in updateData route, but as fallback try current session
        const session = await db.collection('sessions').findOne({ _id: TEAM_SESSION_ID });
        if (session && session.originalRawData) {
            // Try to get from backup first
            const dfuOriginalRecords = session.originalRawData.filter(r => r['DFU'] === dfuCode);
            transferDoc.originalData = JSON.parse(JSON.stringify(dfuOriginalRecords));
            console.log(`[TRANSFER] Using backup original data: ${dfuOriginalRecords.length} records for DFU ${dfuCode}`);
        } else if (session && session.rawData) {
            // Fallback to current data (not ideal but better than nothing)
            const dfuOriginalRecords = session.rawData.filter(r => r['DFU'] === dfuCode);
            transferDoc.originalData = JSON.parse(JSON.stringify(dfuOriginalRecords));
            console.log(`[TRANSFER] WARNING: Using current data as original: ${dfuOriginalRecords.length} records for DFU ${dfuCode}`);
        }
    } else {
        // Preserve existing original data
        transferDoc.originalData = existingTransfer.originalData;
        console.log(`[TRANSFER] Preserving existing original data for DFU ${dfuCode}: ${existingTransfer.originalData?.length || 0} records`);
    }
    
    // Add transfer-specific data
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
                    originalRawData: JSON.parse(JSON.stringify(data)), // Store backup copy
                    dataUploaded: true,
                    uploadedAt: new Date(),
                    uploadedBy: req.body.userName
                }
            }
        );
        
        console.log(`[UPLOAD] Stored ${data.length} records with backup copy`);
        
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
        
        // CRITICAL: Store original data BEFORE updating rawData if this is first transfer
        if (transfer && transfer.dfuCode) {
            const existingTransfer = await db.collection('transfers').findOne({ 
                sessionId: TEAM_SESSION_ID, 
                dfuCode: transfer.dfuCode 
            });
            
            if (!existingTransfer) {
                console.log(`[UPDATE] First transfer for DFU ${transfer.dfuCode} - storing original data now`);
                
                // Get current session data (before update) to store as original
                const currentSession = await db.collection('sessions').findOne({ _id: TEAM_SESSION_ID });
                if (currentSession && currentSession.originalRawData) {
                    // Use the backup original data
                    const dfuOriginalRecords = currentSession.originalRawData.filter(r => r['DFU'] === transfer.dfuCode);
                    console.log(`[UPDATE] Found ${dfuOriginalRecords.length} original records for DFU ${transfer.dfuCode} in backup`);
                    
                    // Store these as original data for this transfer
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
        
        // Now update the session with the new data
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            { $set: { rawData, lastModified: new Date() }}
        );
        
        // Save completed transfers
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
        
        // Find the transfer to undo
        const transferToUndo = await db.collection('transfers').findOne({ 
            sessionId: TEAM_SESSION_ID, 
            dfuCode 
        });
        
        if (!transferToUndo || !transferToUndo.originalData) {
            return res.status(400).json({ error: 'No original data found for this transfer' });
        }
        
        console.log(`[UNDO] Found transfer to undo with ${transferToUndo.originalData.length} original records`);
        console.log(`[UNDO] Sample original record:`, transferToUndo.originalData[0]);
        
        let currentData = [...session.rawData];
        
        // Remove all current records for this DFU
        const beforeCount = currentData.length;
        currentData = currentData.filter(r => r['DFU'] !== dfuCode);
        const afterFilterCount = currentData.length;
        console.log(`[UNDO] Removed ${beforeCount - afterFilterCount} records for DFU ${dfuCode}`);
        
        // Add back the original records
        const originalRecords = JSON.parse(JSON.stringify(transferToUndo.originalData));
        currentData = [...currentData, ...originalRecords];
        console.log(`[UNDO] Added back ${originalRecords.length} original records`);
        console.log(`[UNDO] Final data count: ${currentData.length}`);
        
        // Update session with restored data
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            { 
                $set: { 
                    rawData: currentData,
                    lastModified: new Date()
                }
            }
        );
        
        // Delete the transfer record
        await db.collection('transfers').deleteOne({ 
            sessionId: TEAM_SESSION_ID, 
            dfuCode 
        });
        
        // Get remaining transfers
        const remainingTransfersArray = await db.collection('transfers').find({ 
            sessionId: TEAM_SESSION_ID 
        }).toArray();
        
        const remainingTransfers = {};
        remainingTransfersArray.forEach(t => {
            remainingTransfers[t.dfuCode] = t;
        });
        
        console.log(`[UNDO] Successfully restored DFU ${dfuCode}, ${remainingTransfersArray.length} transfers remain`);
        
        res.json({ 
            success: true, 
            remainingTransfers,
            restoredRecords: originalRecords.length,
            message: `Transfer for DFU ${dfuCode} has been undone - all variants restored to original state`
        });
        
        // Notify all other users
        io.emit('dataUpdated', { 
            dfuCode: dfuCode,
            message: `${userName} undid transfer for DFU ${dfuCode}`,
            updatedBy: userName,
            requiresReload: true
        });
        
    } catch (error) {
        console.error('[UNDO ERROR]', error);
        res.status(500).json({ error: 'Failed to undo transfer: ' + error.message });
    }
});

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

// Helper function - Updated to show ALL DFUs
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