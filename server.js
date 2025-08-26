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

// Helper function to calculate date from week number
function getDateFromWeekNumber(year, weekNumber) {
    const jan1 = new Date(year, 0, 1);
    const jan1DayOfWeek = jan1.getDay();
    const daysToFirstMonday = jan1DayOfWeek === 0 ? 1 : (8 - jan1DayOfWeek) % 7;
    const firstMonday = new Date(year, 0, 1 + daysToFirstMonday);
    const targetDate = new Date(firstMonday);
    targetDate.setDate(firstMonday.getDate() + (weekNumber - 1) * 7);
    return targetDate;
}

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
    const transferDoc = {
        sessionId: TEAM_SESSION_ID,
        dfuCode,
        type: transferData.type || 'individual',
        completedBy: userName,
        completedAt: new Date()
    };
    
    // Add any transfer-specific data
    if (transferData.bulkTransfer) {
        transferDoc.bulkTransfer = transferData.bulkTransfer;
    }
    if (transferData.granularTransfers) {
        transferDoc.granularTransfers = transferData.granularTransfers;
    }
    
    await db.collection('transfers').replaceOne(
        { sessionId: TEAM_SESSION_ID, dfuCode },
        transferDoc,
        { upsert: true }
    );
}

// Delete transfer from DB
async function deleteTransfer(dfuCode) {
    await db.collection('transfers').deleteOne({
        sessionId: TEAM_SESSION_ID,
        dfuCode
    });
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

app.post('/api/undoTransfer', async (req, res) => {
    try {
        const { dfuCode, userName } = req.body;
        
        console.log(`[UNDO] ${userName} undoing transfer for DFU ${dfuCode}`);
        
        // Get the original raw data before any transfers
        const session = await db.collection('sessions').findOne({ _id: TEAM_SESSION_ID });
        
        // Delete from transfers collection
        await deleteTransfer(dfuCode);
        
        // Get all remaining transfers
        const remainingTransfers = await getCompletedTransfers();
        
        res.json({ success: true, remainingTransfers });
        
        // Notify all users with updated transfer list
        io.emit('transferUndone', { 
            dfuCode, 
            undoneBy: userName,
            remainingTransfers 
        });
        
    } catch (error) {
        console.error('[UNDO] Error:', error);
        res.status(500).json({ error: error.message });
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
        
        // Process the data to ensure proper date formatting
        const processedData = (rawData || []).map(record => {
            const processed = { ...record };
            
            // Fix Calendar.week based on Week Number
            if (processed['Week Number']) {
                const weekNum = parseInt(processed['Week Number']);
                
                if (!isNaN(weekNum) && weekNum >= 1 && weekNum <= 52) {
                    // Determine the year - check if Calendar.week has a valid year
                    let year = 2025; // Default year
                    
                    if (processed['Calendar.week']) {
                        const existingDate = new Date(processed['Calendar.week']);
                        if (!isNaN(existingDate.getTime()) && existingDate.getFullYear() > 2000) {
                            year = existingDate.getFullYear();
                        }
                    }
                    
                    const date = getDateFromWeekNumber(year, weekNum);
                    processed['Calendar.week'] = date.toISOString().split('T')[0];
                }
            }
            
            return processed;
        });
        
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(processedData);
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

// Helper function
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
            
            multiVariants[dfuCode] = {
                variants,
                recordCount: grouped[dfuCode].records.length,
                variantDemand
            };
        }
    });
    
    return multiVariants;
}

const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
    console.log(`[SERVER] Running on port ${PORT}`);
});