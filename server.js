// Update session data after client-side transfer (UPDATED VERSION)
app.post('/api/session/:sessionId/updateData', async (req, res) => {
    try {
        const { sessionId } = req.params;
        const { rawData, completedTransfers, transfer } = req.body;
        
        console.log(`📝 Updating session data after transfer for DFU: ${transfer.dfuCode}`);
        
        // Update the session with the modified data and completed transfers
        await db.collection('sessions').updateOne(
            { _id: sessionId },
            { 
                $set: { 
                    rawData: rawData,
                    completedTransfers: completedTransfers, // Store completed transfers status
                    lastModified: new Date(),
                    lastModifiedBy: transfer.completedBy
                }
            }
        );
        
        // Log the transfer
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
        
        console.log(`✅ Data updated successfully`);
        res.json({ success: true });
        
        // Notify other users
        io.to(sessionId).emit('dataUpdated', { 
            dfuCode: transfer.dfuCode,
            updatedBy: transfer.completedBy 
        });
        
    } catch (error) {
        console.error('Error updating data:', error);
        res.status(500).json({ error: 'Failed to update data' });
    }
});

// Get session data (UPDATED VERSION)
app.get('/api/session/:sessionId/data', async (req, res) => {
    try {
        const { sessionId } = req.params;
        
        console.log(`📊 Getting data for session: ${sessionId}`);
        
        const session = await db.collection('sessions').findOne({ 
            _id: sessionId
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
            rawData: session.rawData || [],
            completedTransfers: session.completedTransfers || {} // Include completed transfers status
        });
        
    } catch (error) {
        console.error('Error getting session data:', error);
        res.status(500).json({ error: 'Failed to get session data: ' + error.message });
    }
});

// End session and clear all data (NEW ENDPOINT)
app.post('/api/session/:sessionId/end', async (req, res) => {
    try {
        const { sessionId } = req.params;
        const { userName } = req.body;
        
        console.log(`🔚 Ending session ${sessionId} by ${userName}`);
        
        // Delete all transfers for this session
        await db.collection('transfers').deleteMany({ sessionId });
        
        // Delete the session
        await db.collection('sessions').deleteOne({ _id: sessionId });
        
        console.log(`✅ Session ${sessionId} ended and all data cleared`);
        
        // Notify all connected users that session has ended
        io.to(sessionId).emit('sessionEnded', { endedBy: userName });
        
        res.json({ success: true, message: 'Session ended and data cleared' });
        
    } catch (error) {
        console.error('Error ending session:', error);
        res.status(500).json({ error: 'Failed to end session: ' + error.message });
    }
});