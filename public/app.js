// Global variables
let socket = null;
let sessionId = null;
let userName = null;
let sessionData = {};

// Join session
async function joinSession() {
    const sessionNameInput = document.getElementById('sessionName').value.trim();
    const userNameInput = document.getElementById('userName').value.trim();
    
    if (!sessionNameInput || !userNameInput) {
        showError('Please enter both session name and your name');
        return;
    }
    
    try {
        const response = await fetch('/api/session/join', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ 
                sessionName: sessionNameInput, 
                userName: userNameInput 
            })
        });
        
        const data = await response.json();
        
        if (data.success) {
            sessionId = data.sessionId;
            userName = userNameInput;
            
            // Connect WebSocket
            connectWebSocket();
            
            // Show main app
            document.getElementById('loginScreen').classList.add('hidden');
            document.getElementById('mainApp').classList.remove('hidden');
            document.getElementById('sessionNameDisplay').textContent = sessionNameInput;
            document.getElementById('userNameDisplay').textContent = userNameInput;
            
            // Load session data
            loadSessionData();
            
            showNotification('Successfully joined session!', 'success');
        } else {
            showError(data.error || 'Failed to join session');
        }
    } catch (error) {
        showError('Connection error. Please try again.');
        console.error(error);
    }
}

// Connect WebSocket for real-time updates
function connectWebSocket() {
    socket = io();
    
    socket.emit('joinSession', { sessionId, userName });
    
    socket.on('userJoined', ({ userName }) => {
        showNotification(`${userName} joined the session`, 'info');
    });
    
    socket.on('activeUsers', (users) => {
        document.getElementById('activeUsersDisplay').textContent = 
            `Active Users: ${users.join(', ')}`;
    });
    
    socket.on('dataUploaded', ({ dfuCount }) => {
        showNotification(`Data uploaded: ${dfuCount} DFUs found`, 'success');
        loadSessionData();
    });
    
    socket.on('transferUpdated', ({ dfuCode, userName }) => {
        showNotification(`${userName} updated transfer for DFU ${dfuCode}`, 'info');
        loadSessionData();
    });
}

// Load session data
async function loadSessionData() {
    try {
        const response = await fetch(`/api/session/${sessionId}/data`);
        const data = await response.json();
        
        sessionData = data;
        
        if (data.session.dataUploaded) {
            document.getElementById('uploadSection').classList.add('opacity-50');
            document.getElementById('dfuSection').classList.remove('hidden');
            document.getElementById('generateSection').classList.remove('hidden');
            displayDFUs(data.multiVariantDFUs);
        }
    } catch (error) {
        console.error('Error loading session data:', error);
    }
}

// File upload
document.getElementById('fileInput')?.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    
    const formData = new FormData();
    formData.append('file', file);
    
    document.getElementById('uploadStatus').innerHTML = 
        '<p class="text-blue-600">Uploading and processing file...</p>';
    
    try {
        const response = await fetch(`/api/upload/${sessionId}`, {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (data.success) {
            document.getElementById('uploadStatus').innerHTML = 
                `<p class="text-green-600">✓ Successfully processed ${data.rowCount} rows, found ${data.dfuCount} multi-variant DFUs</p>`;
            loadSessionData();
        } else {
            document.getElementById('uploadStatus').innerHTML = 
                `<p class="text-red-600">Error: ${data.error}</p>`;
        }
    } catch (error) {
        document.getElementById('uploadStatus').innerHTML = 
            '<p class="text-red-600">Upload failed. Please try again.</p>';
        console.error(error);
    }
});

// Display DFUs
function displayDFUs(multiVariantDFUs) {
    const dfuList = document.getElementById('dfuList');
    
    if (Object.keys(multiVariantDFUs).length === 0) {
        dfuList.innerHTML = '<p class="text-gray-600">No multi-variant DFUs found</p>';
        return;
    }
    
    dfuList.innerHTML = Object.keys(multiVariantDFUs).map(dfuCode => {
        const dfu = multiVariantDFUs[dfuCode];
        return `
            <div class="border rounded-lg p-4">
                <div class="flex justify-between items-center">
                    <div>
                        <h3 class="font-semibold">DFU: ${dfuCode}</h3>
                        <p class="text-sm text-gray-600">
                            ${dfu.variants.length} variants | ${dfu.recordCount} records
                        </p>
                    </div>
                    <button 
                        onclick="configureTransfer('${dfuCode}')" 
                        class="px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700"
                    >
                        Configure Transfer
                    </button>
                </div>
                <div class="mt-2">
                    <p class="text-sm">Variants: ${dfu.variants.join(', ')}</p>
                </div>
            </div>
        `;
    }).join('');
}

// Configure transfer (simplified)
async function configureTransfer(dfuCode) {
    const dfu = sessionData.multiVariantDFUs[dfuCode];
    const targetVariant = prompt(`Transfer all variants to which one?\n\nAvailable variants:\n${dfu.variants.join('\n')}`);
    
    if (targetVariant && dfu.variants.includes(targetVariant)) {
        try {
            const response = await fetch(`/api/session/${sessionId}/transfer`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    dfuCode,
                    transferConfig: {
                        type: 'bulk',
                        targetVariant
                    },
                    userName
                })
            });
            
            if (response.ok) {
                showNotification(`Transfer configured for DFU ${dfuCode}`, 'success');
                loadSessionData();
            }
        } catch (error) {
            console.error('Error saving transfer:', error);
            showNotification('Failed to save transfer', 'error');
        }
    }
}

// Generate final file
async function generateFile() {
    try {
        const response = await fetch(`/api/session/${sessionId}/generate`, {
            method: 'POST'
        });
        
        if (response.ok) {
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `DFU_Transfers_${Date.now()}.xlsx`;
            a.click();
            
            showNotification('File generated successfully!', 'success');
        } else {
            showNotification('Failed to generate file', 'error');
        }
    } catch (error) {
        console.error('Error generating file:', error);
        showNotification('Failed to generate file', 'error');
    }
}

// UI Helper functions
function showError(message) {
    const errorDiv = document.getElementById('loginError');
    errorDiv.textContent = message;
    errorDiv.classList.remove('hidden');
}

function showNotification(message, type = 'info') {
    const notifications = document.getElementById('notifications');
    const notification = document.createElement('div');
    
    const colors = {
        success: 'bg-green-500',
        error: 'bg-red-500',
        info: 'bg-blue-500'
    };
    
    notification.className = `${colors[type]} text-white px-4 py-2 rounded-lg shadow-lg`;
    notification.textContent = message;
    
    notifications.appendChild(notification);
    
    setTimeout(() => {
        notification.remove();
    }, 3000);
}