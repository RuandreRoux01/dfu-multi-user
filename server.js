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
        // Ensure collections exist
        await db.createCollection('sessions').catch(() => {});
        await db.createCollection('transfers').catch(() => {});
        
        // Initialize or restore session
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

// Helper to get completed transfers from DB
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

// Save individual transfer to DB
async function saveTransfer(dfuCode, transferData, userName) {
    console.log(`[SAVE TRANSFER] Saving transfer for DFU ${dfuCode} by ${userName}`);
    
    const transferDoc = {
        sessionId: TEAM_SESSION_ID,
        dfuCode,
        type: transferData.type || 'individual',
        completedBy: userName,
        completedAt: new Date()
    };
    
    // Check if this transfer already exists (to preserve original data)
    const existingTransfer = await db.collection('transfers').findOne({ 
        sessionId: TEAM_SESSION_ID, 
        dfuCode 
    });
    
    // Store original data ONLY if this is the first transfer for this DFU
    if (!existingTransfer) {
        const session = await db.collection('sessions').findOne({ _id: TEAM_SESSION_ID });
        if (session && session.rawData) {
            const dfuOriginalRecords = session.rawData.filter(r => r['DFU'] === dfuCode);
            transferDoc.originalData = JSON.parse(JSON.stringify(dfuOriginalRecords)); // Deep copy
            console.log(`[SAVE TRANSFER] Storing ${dfuOriginalRecords.length} original records for DFU ${dfuCode}`);
        }
    } else {
        // Preserve existing original data
        transferDoc.originalData = existingTransfer.originalData;
        console.log(`[SAVE TRANSFER] Preserving existing original data for DFU ${dfuCode}`);
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
    
    console.log(`[SAVE TRANSFER] Transfer saved successfully for DFU ${dfuCode}`);
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
        
        // Clear old transfers when new data uploaded
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
        
        // Update raw data
        await db.collection('sessions').updateOne(
            { _id: TEAM_SESSION_ID },
            { $set: { rawData, lastModified: new Date() }}
        );
        
        // Save each completed transfer to DB
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
        
        // Validate input
        if (!dfuCode || !variantCode) {
            return res.status(400).json({ error: 'Missing required fields' });
        }
        
        // Ensure newRecords is an array
        const recordsToAdd = Array.isArray(newRecords) ? newRecords : [];
        
        // Get current session data
        const session = await db.collection('sessions').findOne({ _id: TEAM_SESSION_ID });
        
        if (!session || !session.rawData) {
            return res.status(400).json({ error: 'No session data found' });
        }
        
        // If no new records provided, create them based on existing DFU records
        if (recordsToAdd.length === 0) {
            const dfuRecords = session.rawData.filter(r => r['DFU'] === dfuCode);
            
            if (dfuRecords.length > 0) {
                // Get unique week/location combinations
                const uniqueCombos = new Map();
                dfuRecords.forEach(r => {
                    const key = `${r['Week Number']}_${r['Source Location']}`;
                    if (!uniqueCombos.has(key)) {
                        uniqueCombos.set(key, r);
                    }
                });
                
                // Create new records for each unique combination
                uniqueCombos.forEach((templateRecord) => {
                    const newRecord = { ...templateRecord };
                    newRecord['Product Number'] = variantCode.trim();
                    newRecord['weekly fcst'] = 0;
                    newRecord['PartDescription'] = 'Manually added variant';
                    newRecord['Transfer History'] = `Manually added by ${userName} on ${new Date().toLocaleString()}`;
                    recordsToAdd.push(newRecord);
                });
            } else {
                // Create at least one record if no DFU records found
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
        
        // Add new records to rawData
        const updatedRawData = [...session.rawData, ...recordsToAdd];
        
        // Update session with new raw data
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
        
        // Notify all users
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
        
        // Get the current session data
        const session = await db.collection('sessions').findOne({ _id: TEAM_SESSION_ID });
        if (!session || !session.rawData) {
            return res.status(400).json({ error: 'No session data found' });
        }
        
        // Get the specific transfer being undone to retrieve original data
        const transferToUndo = await db.collection('transfers').findOne({ 
            sessionId: TEAM_SESSION_ID, 
            dfuCode 
        });
        
        if (!transferToUndo) {
            return res.status(400).json({ error: 'No transfer found for this DFU' });
        }
        
        // Start with current session data
        let currentData = [...session.rawData];
        
        // Remove all current records for this DFU
        currentData = currentData.filter(r => r['DFU'] !== dfuCode);
        
        // If we have original data stored, restore it
        if (transferToUndo.originalData && transferToUndo.originalData.length > 0) {
            console.log(`[UNDO] Restoring ${transferToUndo.originalData.length} original records for DFU ${dfuCode}`);
            
            // Add back the original records
            currentData = [...currentData, ...transferToUndo.originalData];
        } else {
            console.log(`[UNDO] No original data found for DFU ${dfuCode}, attempting reconstruction`);
            
            // If no original data stored, we need to reconstruct from transfer history
            // This is a fallback - ideally originalData should always be stored
            return res.status(400).json({ 
                error: 'Cannot undo - original data not found. Please reload the original file.' 
            });
        }
        
        // Update the session with restored data
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
        
        // Get all remaining transfers for the response
        const remainingTransfers = await db.collection('transfers').find({ 
            sessionId: TEAM_SESSION_ID 
        }).toArray();
        
        // Convert array to object format expected by frontend
        const completedTransfers = {};
        remainingTransfers.forEach(t => {
            completedTransfers[t.dfuCode] = t;
        });
        
        console.log(`[UNDO] Successfully restored DFU ${dfuCode}, ${remainingTransfers.length} transfers remain`);
        
        res.json({ 
            success: true, 
            completedTransfers: completedTransfers,
            message: `Transfer for DFU ${dfuCode} has been undone - all variants restored to original state`
        });
        
        // Notify all users of the update
        io.emit('dataUpdated', { 
            dfuCode: dfuCode,
            message: `${userName} undid transfer for DFU ${dfuCode}`,
            updatedBy: userName 
        });
        
    } catch (error) {
        console.error('[UNDO ERROR]', error);
        res.status(500).json({ error: 'Failed to undo transfer: ' + error.message });
    }
});

app.post('/api/clear', async (req, res) => {
    try {
        // Clear session data
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
        
        // Clear all transfers
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
        // Include ALL DFUs, not just multi-variant ones
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