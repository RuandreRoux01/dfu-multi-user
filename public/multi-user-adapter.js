// Multi-User Adapter for DFU Transfer App
// This bridges the original app with the multi-user backend

// Wait for DOM to be ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeMultiUser);
} else {
    // DOM is already ready
    initializeMultiUser();
}

function initializeMultiUser() {
    console.log('Initializing multi-user adapter...');
    
    let socket = null;
    let sessionId = null;
    let userName = null;
    
    // Check if DemandTransferApp exists
    if (typeof DemandTransferApp === 'undefined') {
        console.error('DemandTransferApp not found! Waiting...');
        setTimeout(initializeMultiUser, 500);
        return;
    }
    
    // Store original methods
    const originalInit = DemandTransferApp.prototype.init;
    const originalLoadData = DemandTransferApp.prototype.loadData;
    const originalExecuteTransfer = DemandTransferApp.prototype.executeTransfer;
    
    // Override the init method
    DemandTransferApp.prototype.init = function() {
        console.log('Overridden init called');
        this.showSessionModal();
    };
    
    // Add session modal functionality
    DemandTransferApp.prototype.showSessionModal = function() {
        console.log('Showing session modal...');
        const modal = document.getElementById('sessionModal');
        const joinBtn = document.getElementById('joinSessionBtn');
        const app = this;
        
        if (!modal || !joinBtn) {
            console.error('Modal elements not found!');
            return;
        }
        
        // Remove any existing listeners
        const newJoinBtn = joinBtn.cloneNode(true);
        joinBtn.parentNode.replaceChild(newJoinBtn, joinBtn);
        
        // Add click listener
        newJoinBtn.addEventListener('click', async function(e) {
            console.log('Join button clicked!');
            e.preventDefault();
            
            const sessionNameInput = document.getElementById('sessionName').value.trim();
            const userNameInput = document.getElementById('userName').value.trim();
            
            console.log('Session name:', sessionNameInput);
            console.log('User name:', userNameInput);
            
            if (!sessionNameInput || !userNameInput) {
                alert('Please enter both session name and your name');
                return;
            }
            
            // Disable button to prevent double-clicks
            newJoinBtn.disabled = true;
            newJoinBtn.textContent = 'Joining...';
            
            try {
                await app.joinSession(sessionNameInput, userNameInput);
            } catch (error) {
                console.error('Error joining session:', error);
                newJoinBtn.disabled = false;
                newJoinBtn.textContent = 'Join Session';
            }
        });
        
        console.log('Event listener attached to join button');
    };
    
    // Add join session functionality
    DemandTransferApp.prototype.joinSession = async function(sessionName, userNameValue) {
        console.log('Joining session:', sessionName, 'as', userNameValue);
        const app = this;
        
        try {
            const response = await fetch('/api/session/join', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ sessionName, userName: userNameValue })
            });
            
            console.log('Response status:', response.status);
            const data = await response.json();
            console.log('Response data:', data);
            
            if (data.success) {
                sessionId = data.sessionId;
                userName = userNameValue;
                app.sessionId = sessionId;
                app.userName = userName;
                
                console.log('Session joined successfully:', sessionId);
                
                // Connect WebSocket
                app.connectWebSocket();
                
                // Hide modal
                document.getElementById('sessionModal').style.display = 'none';
                
                // Show active users panel
                const activeUsersPanel = document.getElementById('activeUsersPanel');
                if (activeUsersPanel) {
                    activeUsersPanel.classList.remove('hidden');
                }
                
                // Now initialize the original app functionality
                console.log('Calling original init...');
                originalInit.call(app);
                
                // Load session data
                setTimeout(() => {
                    app.loadSessionData();
                }, 500);
                
            } else {
                throw new Error(data.error || 'Failed to join session');
            }
        } catch (error) {
            console.error('Error joining session:', error);
            alert('Failed to join session: ' + error.message);
            throw error;
        }
    };
    
    // Add WebSocket connection
    DemandTransferApp.prototype.connectWebSocket = function() {
        console.log('Connecting WebSocket...');
        const app = this;
        
        try {
            socket = io();
            app.socket = socket;
            
            socket.on('connect', () => {
                console.log('WebSocket connected');
                socket.emit('joinSession', { sessionId, userName });
            });
            
            socket.on('userJoined', ({ userName }) => {
                console.log('User joined:', userName);
                if (app.showNotification) {
                    app.showNotification(`${userName} joined the session`, 'info');
                }
            });
            
            socket.on('activeUsers', (users) => {
                console.log('Active users:', users);
                const usersList = document.getElementById('activeUsersList');
                if (usersList) {
                    usersList.innerHTML = users.map(u => `<div class="text-sm py-1">${u}</div>`).join('');
                }
            });
            
            socket.on('dataUploaded', () => {
                console.log('Data uploaded by another user');
                if (app.showNotification) {
                    app.showNotification('Another user uploaded data', 'info');
                }
                app.loadSessionData();
            });
        } catch (error) {
            console.error('Error connecting WebSocket:', error);
        }
    };
    
    // Add load session data functionality
    DemandTransferApp.prototype.loadSessionData = async function() {
        console.log('Loading session data...');
        const app = this;
        
        try {
            const response = await fetch(`/api/session/${sessionId}/data`);
            const data = await response.json();
            
            console.log('Session data loaded:', data);
            
            if (data.rawData && data.rawData.length > 0) {
                app.rawData = data.rawData;
                app.originalRawData = JSON.parse(JSON.stringify(data.rawData));
                app.processMultiVariantDFUs(data.rawData);
                app.isProcessed = true;
                app.render();
                console.log('Data processed and rendered');
            }
        } catch (error) {
            console.error('Error loading session data:', error);
        }
    };
    
    // Override the original loadData method
    DemandTransferApp.prototype.loadData = async function(file) {
        console.log('Loading data - multi-user mode');
        const app = this;
        
        if (!sessionId) {
            alert('Please join a session first!');
            return;
        }
        
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
            console.log('Upload result:', result);
            
            if (result.success) {
                app.showNotification(`Successfully uploaded ${result.rowCount} records`);
                // Load the processed data
                await app.loadSessionData();
            } else {
                app.showNotification('Failed to upload file: ' + (result.error || 'Unknown error'), 'error');
            }
        } catch (error) {
            console.error('Error uploading file:', error);
            app.showNotification('Failed to upload file: ' + error.message, 'error');
        } finally {
            app.isLoading = false;
            app.render();
        }
    };
    
    // Override executeTransfer
    DemandTransferApp.prototype.executeTransfer = function(dfuCode) {
        console.log('Executing transfer - multi-user mode');
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
            }).then(response => response.json())
              .then(data => console.log('Transfer saved:', data))
              .catch(error => console.error('Error saving transfer:', error));
        }
    };
    
    // Override exportData
    DemandTransferApp.prototype.exportData = async function() {
        console.log('Exporting data - multi-user mode');
        const app = this;
        
        if (!sessionId) {
            alert('Please join a session first!');
            return;
        }
        
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
            app.showNotification('Failed to export file: ' + error.message, 'error');
        }
    };
    
    console.log('Multi-user adapter initialized successfully');
}