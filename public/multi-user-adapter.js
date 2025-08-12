// Multi-User Adapter for DFU Transfer App
// This bridges the original app with the multi-user backend

(function() {
    // Wait for the original app to load
    let originalApp = null;
    let socket = null;
    let sessionId = null;
    let userName = null;
    
    // Override the DemandTransferApp class to add multi-user features
    const OriginalDemandTransferApp = window.DemandTransferApp;
    
    window.DemandTransferApp = class MultiUserDemandTransferApp extends OriginalDemandTransferApp {
        constructor() {
            super();
            originalApp = this;
            
            // Don't initialize until session is joined
            this.isProcessed = false;
            this.render = this.render.bind(this);
        }
        
        init() {
            // Show session modal instead of regular init
            this.showSessionModal();
        }
        
        showSessionModal() {
            const modal = document.getElementById('sessionModal');
            const joinBtn = document.getElementById('joinSessionBtn');
            
            joinBtn.addEventListener('click', async () => {
                const sessionNameInput = document.getElementById('sessionName').value;
                const userNameInput = document.getElementById('userName').value;
                
                if (!sessionNameInput || !userNameInput) {
                    alert('Please enter both session name and your name');
                    return;
                }
                
                await this.joinSession(sessionNameInput, userNameInput);
            });
        }
        
        async joinSession(sessionName, userNameValue) {
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
                    
                    // Connect WebSocket
                    this.connectWebSocket();
                    
                    // Hide modal
                    document.getElementById('sessionModal').style.display = 'none';
                    
                    // Show active users panel
                    document.getElementById('activeUsersPanel').classList.remove('hidden');
                    
                    // Initialize the original app
                    super.init();
                    
                    // Load session data
                    this.loadSessionData();
                }
            } catch (error) {
                console.error('Error joining session:', error);
                alert('Failed to join session. Please try again.');
            }
        }
        
        connectWebSocket() {
            socket = io();
            
            socket.emit('joinSession', { sessionId, userName });
            
            socket.on('userJoined', ({ userName }) => {
                this.showNotification(`${userName} joined the session`, 'info');
            });
            
            socket.on('activeUsers', (users) => {
                const usersList = document.getElementById('activeUsersList');
                if (usersList) {
                    usersList.innerHTML = users.map(u => `<div class="text-sm py-1">${u}</div>`).join('');
                }
            });
            
            socket.on('dataUploaded', () => {
                this.showNotification('Another user uploaded data', 'info');
                this.loadSessionData();
            });
        }
        
        async loadSessionData() {
            try {
                const response = await fetch(`/api/session/${sessionId}/data`);
                const data = await response.json();
                
                if (data.rawData && data.rawData.length > 0) {
                    this.rawData = data.rawData;
                    this.originalRawData = JSON.parse(JSON.stringify(data.rawData));
                    this.processMultiVariantDFUs(data.rawData);
                    this.isProcessed = true;
                    this.render();
                }
            } catch (error) {
                console.error('Error loading session data:', error);
            }
        }
        
        // Override loadData to upload to server
        async loadData(file) {
            console.log('Uploading file to server...');
            this.isLoading = true;
            this.render();
            
            const formData = new FormData();
            formData.append('file', file);
            
            try {
                const response = await fetch(`/api/upload/${sessionId}`, {
                    method: 'POST',
                    body: formData
                });
                
                const result = await response.json();
                
                if (result.success) {
                    this.showNotification(`Successfully uploaded ${result.rowCount} records`);
                    // Load the processed data
                    await this.loadSessionData();
                } else {
                    this.showNotification('Failed to upload file', 'error');
                }
            } catch (error) {
                console.error('Error uploading file:', error);
                this.showNotification('Failed to upload file', 'error');
            } finally {
                this.isLoading = false;
            }
        }
        
        // Override executeTransfer to sync with server
        executeTransfer(dfuCode) {
            // First execute locally
            super.executeTransfer(dfuCode);
            
            // Then save to server
            if (socket && sessionId) {
                fetch(`/api/session/${sessionId}/transfer`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        dfuCode,
                        transferConfig: {
                            type: this.bulkTransfers[dfuCode] ? 'bulk' : 'individual',
                            targetVariant: this.bulkTransfers[dfuCode],
                            transfers: this.transfers[dfuCode]
                        },
                        userName
                    })
                });
            }
        }
        
        // Override exportData to use server
        async exportData() {
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
                    
                    this.showNotification('File exported successfully');
                } else {
                    this.showNotification('Failed to export file', 'error');
                }
            } catch (error) {
                console.error('Error exporting data:', error);
                this.showNotification('Failed to export file', 'error');
            }
        }
    };
})();