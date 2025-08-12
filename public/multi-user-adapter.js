// Multi-User Adapter for DFU Transfer App
// This bridges the original app with the multi-user backend

// Wait for DOM to be ready
document.addEventListener('DOMContentLoaded', function() {
    let socket = null;
    let sessionId = null;
    let userName = null;
    let originalApp = null;
    
    // Prevent the original app from auto-initializing
    const originalInit = DemandTransferApp.prototype.init;
    
    // Override the init method
    DemandTransferApp.prototype.init = function() {
        // Show session modal instead of regular init
        this.showSessionModal();
    };
    
    // Add session modal functionality
    DemandTransferApp.prototype.showSessionModal = function() {
        const modal = document.getElementById('sessionModal');
        const joinBtn = document.getElementById('joinSessionBtn');
        const app = this;
        
        joinBtn.addEventListener('click', async () => {
            const sessionNameInput = document.getElementById('sessionName').value;
            const userNameInput = document.getElementById('userName').value;
            
            if (!sessionNameInput || !userNameInput) {
                alert('Please enter both session name and your name');
                return;
            }
            
            await app.joinSession(sessionNameInput, userNameInput);
        });
    };
    
    // Add join session functionality
    DemandTransferApp.prototype.joinSession = async function(sessionName, userNameValue) {
        const app = this;
        
        try {
            const response = await fetch('/api/session/join', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ sessionName, userName: userNameValue })
            });
            
            const data = await response.json();
            
            if (data.success) {
                sessionId = data.sessionId;
                userName = userNameValue;
                app.sessionId = sessionId;
                app.userName = userName;
                
                // Connect WebSocket
                app.connectWebSocket();
                
                // Hide modal
                document.getElementById('sessionModal').style.display = 'none';
                
                // Show active users panel
                document.getElementById('activeUsersPanel').classList.remove('hidden');
                
                // Now initialize the original app functionality
                originalInit.call(app);
                
                // Load session data
                app.loadSessionData();
            }
        } catch (error) {
            console.error('Error joining session:', error);
            alert('Failed to join session. Please try again.');
        }
    };
    
    // Add WebSocket connection
    DemandTransferApp.prototype.connectWebSocket = function() {
        const app = this;
        socket = io();
        app.socket = socket;
        
        socket.emit('joinSession', { sessionId, userName });
        
        socket.on('userJoined', ({ userName }) => {
            app.showNotification(`${userName} joined the session`, 'info');
        });
        
        socket.on('activeUsers', (users) => {
            const usersList = document.getElementById('activeUsersList');
            if (usersList) {
                usersList.innerHTML = users.map(u => `<div class="text-sm py-1">${u}</div>`).join('');
            }
        });
        
        socket.on('dataUploaded', () => {
            app.showNotification('Another user uploaded data', 'info');
            app.loadSessionData();
        });
    };
    
    // Add load session data functionality
    DemandTransferApp.prototype.loadSessionData = async function() {
        const app = this;
        
        try {
            const response = await fetch(`/api/session/${sessionId}/data`);
            const data = await response.json();
            
            if (data.rawData && data.rawData.length > 0) {
                app.rawData = data.rawData;
                app.originalRawData = JSON.parse(JSON.stringify(data.rawData));
                app.processMultiVariantDFUs(data.rawData);
                app.isProcessed = true;
                app.render();
            }
        } catch (error) {
            console.error('Error loading session data:', error);
        }
    };
    
    // Override the original loadData method
    const originalLoadData = DemandTransferApp.prototype.loadData;
    DemandTransferApp.prototype.loadData = async function(file) {
        const app = this;
        
        console.log('Uploading file to server...');
        app.isLoading = true;
        app.render();
        
        const formData = new FormData();
        formData.append('file', file);
        
        try {
            const response = await fetch(`/api/upload/${sessionId}`, {
                method: 'POST',
                body: formData
            });
            
            const result = await response.json();
            
            if (result.success) {
                app.showNotification(`Successfully uploaded ${result.rowCount} records`);
                // Load the processed data
                await app.loadSessionData();
            } else {
                app.showNotification('Failed to upload file', 'error');
            }
        } catch (error) {
            console.error('Error uploading file:', error);
            app.showNotification('Failed to upload file', 'error');
        } finally {
            app.isLoading = false;
        }
    };
    
    // Override executeTransfer
    const originalExecuteTransfer = DemandTransferApp.prototype.executeTransfer;
    DemandTransferApp.prototype.executeTransfer = function(dfuCode) {
        const app = this;
        
        // First execute locally
        originalExecuteTransfer.call(app, dfuCode);
        
        // Then save to server
        if (socket && sessionId) {
            fetch(`/api/session/${sessionId}/transfer`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    dfuCode,
                    transferConfig: {
                        type: app.bulkTransfers[dfuCode] ? 'bulk' : 'individual',
                        targetVariant: app.bulkTransfers[dfuCode],
                        transfers: app.transfers[dfuCode]
                    },
                    userName
                })
            });
        }
    };
    
    // Override exportData
    DemandTransferApp.prototype.exportData = async function() {
        const app = this;
        
        try {
            const response = await fetch(`/api/session/${sessionId}/generate`, {
                method: 'POST'
            });
            
            if (response.ok) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `DFU_Transfers_${sessionId}_${Date.now()}.xlsx`;
                a.click();
                
                app.showNotification('File exported successfully');
            } else {
                app.showNotification('Failed to export file', 'error');
            }
        } catch (error) {
            console.error('Error exporting data:', error);
            app.showNotification('Failed to export file', 'error');
        }
    };
});