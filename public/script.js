// DFU Demand Transfer Management Application
// Version: 2.18.0 - Build: 2025-08-05-scrolling-fix
// Fixed DFU list scrolling and Calendar.week date format

class DemandTransferApp {
    constructor() {
        this.rawData = [];
        this.originalRawData = []; // Backup of original data for undo functionality
        this.multiVariantDFUs = {};
        this.filteredDFUs = {};
        this.selectedDFU = null;
        this.searchTerm = '';
        this.selectedPlantLocation = '';
        this.availablePlantLocations = [];
        this.transfers = {}; // Format: { dfuCode: { sourceVariant: targetVariant } }
        this.bulkTransfers = {}; // Format: { dfuCode: targetVariant }
        this.granularTransfers = {}; // Format: { dfuCode: { sourceVariant: { targetVariant: { weekKey: { selected: boolean, customQuantity: number } } } } }
        this.completedTransfers = {}; // Format: { dfuCode: { type: 'bulk'|'individual', targetVariant, timestamp } }
        this.isProcessed = false;
        this.isLoading = false;
        this.lastExecutionSummary = {}; // Store last execution summary for display
        this.variantCycleDates = {}; // Store SOS/EOS data: { dfuCode: { partCode: { sos: date, eos: date } } }
        this.hasVariantCycleData = false; // Flag to check if cycle data is loaded
        this.keepZeroVariants = true; // Flag to keep variants with 0 demand visible
        this.searchDebounceTimer = null; // Debounce timer for search input
        
        this.init();
    }
    
    init() {
        console.log('🚀 DFU Demand Transfer App v2.18.0 - Build: 2025-08-05-scrolling-fix');
        console.log('📋 Fixed DFU list scrolling and Calendar.week date format');
        this.render();
        this.attachEventListeners();
    }
    
    // Helper method to ensure consistent string comparison
    toComparableString(value) {
        if (value === null || value === undefined) return '';
        return String(value).trim();
    }
    
    showNotification(message, type = 'success') {
        const notification = document.createElement('div');
        notification.className = `notification ${type}`;
        notification.textContent = message;
        document.body.appendChild(notification);
        
        setTimeout(() => {
            notification.remove();
        }, 3000);
    }
    
    formatNumber(num) {
        return new Intl.NumberFormat().format(num);
    }
    
    async loadVariantCycleData(file) {
        console.log('Starting to load variant cycle data...');
        
        try {
            const arrayBuffer = await file.arrayBuffer();
            console.log('Variant cycle array buffer size:', arrayBuffer.byteLength);
            
            const workbook = XLSX.read(arrayBuffer, { 
                cellStyles: true, 
                cellFormulas: true, 
                cellDates: true,
                cellNF: true,
                sheetStubs: true
            });
            
            console.log('Available sheets in cycle file:', workbook.SheetNames);
            
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const data = XLSX.utils.sheet_to_json(worksheet);
            
            console.log('Loaded variant cycle data:', data.length, 'records');
            
            if (data.length > 0) {
                console.log('Sample cycle record:', data[0]);
                console.log('Available columns:', Object.keys(data[0]));
                
                // Process the cycle data
                this.processCycleData(data);
                this.hasVariantCycleData = true;
                this.showNotification(`Successfully loaded ${data.length} variant cycle records`);
                
                // Re-render to show the new data
                this.render();
            } else {
                this.showNotification('No data found in the variant cycle file', 'error');
            }
            
        } catch (error) {
            console.error('Error loading variant cycle data:', error);
            this.showNotification('Error loading variant cycle data: ' + error.message, 'error');
        }
    }
    
    processCycleData(data) {
        console.log('Processing cycle data...');
        
        // Clear existing cycle data
        this.variantCycleDates = {};
        
        // Expected column names (adjust if your file has different column names)
        const dfuColumn = 'DFU';
        const partCodeColumn = 'Part Code';
        const sosColumn = 'SOS';
        const eosColumn = 'EOS';
        const commentsColumn = 'Comments';
        
        // Check if columns exist
        if (data.length > 0) {
            const sampleRecord = data[0];
            const columns = Object.keys(sampleRecord);
            console.log('Cycle data columns:', columns);
            
            // Find actual column names (case-insensitive and flexible matching)
            const actualDfuColumn = columns.find(col => col.toUpperCase() === 'DFU') || dfuColumn;
            const actualPartCodeColumn = columns.find(col => 
                col.toUpperCase().includes('PART') || 
                col.toUpperCase().includes('CODE') ||
                col.toUpperCase() === 'PART CODE'
            ) || partCodeColumn;
            const actualSosColumn = columns.find(col => col.toUpperCase() === 'SOS') || sosColumn;
            const actualEosColumn = columns.find(col => col.toUpperCase() === 'EOS') || eosColumn;
            const actualCommentsColumn = columns.find(col => col.toUpperCase() === 'COMMENTS' || col.toUpperCase() === 'COMMENT') || commentsColumn;
            
            console.log('Using columns:', { 
                dfu: actualDfuColumn, 
                partCode: actualPartCodeColumn, 
                sos: actualSosColumn, 
                eos: actualEosColumn,
                comments: actualCommentsColumn
            });
            
            // Process each record
            let processedCount = 0;
            data.forEach(record => {
                const dfuCode = record[actualDfuColumn];
                const partCode = record[actualPartCodeColumn];
                const sos = record[actualSosColumn];
                const eos = record[actualEosColumn];
                const comments = record[actualCommentsColumn];
                
                if (dfuCode && partCode) {
                    // Ensure both are strings for consistent comparison
                    const dfuStr = this.toComparableString(dfuCode);
                    const partStr = this.toComparableString(partCode);
                    
                    if (!this.variantCycleDates[dfuStr]) {
                        this.variantCycleDates[dfuStr] = {};
                    }
                    
                    this.variantCycleDates[dfuStr][partStr] = {
                        sos: sos || 'N/A',
                        eos: eos || 'N/A',
                        comments: comments || ''
                    };
                    processedCount++;
                }
            });
            
            console.log(`Processed ${processedCount} cycle data records`);
            console.log('Cycle data for DFUs:', Object.keys(this.variantCycleDates).length);
            console.log('Sample cycle data:', this.variantCycleDates);
        }
    }
    
    getCycleDataForVariant(dfuCode, partCode) {
        // Ensure both are strings for consistent comparison
        const dfuStr = this.toComparableString(dfuCode);
        const partStr = this.toComparableString(partCode);
        
        if (this.variantCycleDates[dfuStr] && this.variantCycleDates[dfuStr][partStr]) {
            return this.variantCycleDates[dfuStr][partStr];
        }
        return null;
    }
    
    handleFileUpload(event) {
        const file = event.target.files[0];
        console.log('File selected:', file);
        
        if (!file) {
            console.log('No file selected');
            return;
        }
        
        if (!file.name.toLowerCase().endsWith('.xlsx') && !file.name.toLowerCase().endsWith('.xls')) {
            this.showNotification('Please select an Excel file (.xlsx or .xls)', 'error');
            return;
        }
        
        this.loadData(file);
    }
    
    async loadData(file) {
        console.log('Starting to load data...');
        this.isLoading = true;
        this.render();
        
        try {
            const arrayBuffer = await file.arrayBuffer();
            console.log('Array buffer size:', arrayBuffer.byteLength);
            
            const workbook = XLSX.read(arrayBuffer, { 
                cellStyles: true, 
                cellFormulas: true, 
                cellDates: true,
                cellNF: true,
                sheetStubs: true
            });
            
            console.log('Available sheets:', workbook.SheetNames);
            
            // Updated sheet name detection for new format
            let sheetName = 'Total Demand';
            if (!workbook.Sheets[sheetName]) {
                // Try other common sheet names
                const possibleNames = ['Open Fcst', 'Demand', 'Sheet1'];
                sheetName = possibleNames.find(name => workbook.Sheets[name]) || workbook.SheetNames[0];
                console.log('Using sheet:', sheetName);
            }
            
            const worksheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(worksheet);
            
            console.log('Data conversion complete');
            console.log('Loaded data:', data.length, 'records');
            
            if (data.length > 0) {
                console.log('Sample record:', data[0]);
                console.log('Available columns:', Object.keys(data[0]));
                
                // Log data types for debugging
                const sampleRow = data[0];
                console.log('Data types in first row:');
                Object.keys(sampleRow).forEach(key => {
                    console.log(`  ${key}: ${typeof sampleRow[key]} (${sampleRow[key]})`);
                });
                
                this.rawData = data;
                // Create a deep copy of the original data for undo functionality
                this.originalRawData = JSON.parse(JSON.stringify(data));
                this.processMultiVariantDFUs(data);
                this.isProcessed = true;
                this.showNotification(`Successfully loaded ${data.length} records`);
            } else {
                this.showNotification('No data found in the Excel file', 'error');
            }
            
        } catch (error) {
            console.error('Error loading data:', error);
            this.showNotification('Error loading data: ' + error.message, 'error');
        } finally {
            this.isLoading = false;
            this.render();
        }
    }
    
    processMultiVariantDFUs(data) {
        console.log('Processing data:', data.length, 'records');
        
        if (data.length === 0) {
            this.showNotification('No data found in the file', 'error');
            return;
        }
        
        const sampleRecord = data[0];
        console.log('Sample record:', sampleRecord);
        console.log('Available columns:', Object.keys(sampleRecord));
        
        const columns = Object.keys(sampleRecord);
        
        // Updated column mapping for new file format
        const dfuColumn = 'DFU';
        const partNumberColumn = 'Product Number';
        const demandColumn = 'weekly fcst';
        const partDescriptionColumn = 'PartDescription';
        const plantLocationColumn = 'Plant Location';
        const calendarWeekColumn = 'Calendar.week';
        const sourceLocationColumn = 'Source Location';
        const weekNumberColumn = 'Week Number';
        
        console.log('Using column mapping:', { 
            dfuColumn, 
            partNumberColumn, 
            demandColumn, 
            partDescriptionColumn, 
            plantLocationColumn,
            calendarWeekColumn,
            sourceLocationColumn,
            weekNumberColumn
        });
        
        // Validate required columns exist
        const requiredColumns = [dfuColumn, partNumberColumn, demandColumn, plantLocationColumn, weekNumberColumn];
        const missingColumns = requiredColumns.filter(col => !columns.includes(col));
        
        if (missingColumns.length > 0) {
            this.showNotification(`Missing required columns: ${missingColumns.join(', ')}`, 'error');
            console.error('Missing columns:', missingColumns);
            console.log('Available columns:', columns);
            return;
        }
        
        // Extract unique plant locations for filtering
        this.availablePlantLocations = [...new Set(data.map(record => this.toComparableString(record[plantLocationColumn])))].filter(Boolean).sort();
        console.log('Available Plant Locations:', this.availablePlantLocations);
        
        const groupedByDFU = {};
        
        // Filter data by plant location if selected
        const filteredData = this.selectedPlantLocation ? 
            data.filter(record => this.toComparableString(record[plantLocationColumn]) === this.selectedPlantLocation) : 
            data;
            
        console.log('Total data records:', data.length);
        console.log('Filtered data records:', filteredData.length, 'for plant location:', this.selectedPlantLocation || 'All');
        
        if (this.selectedPlantLocation && filteredData.length === 0) {
            console.warn('No records found for plant location:', this.selectedPlantLocation);
        }
        
        filteredData.forEach(record => {
            const dfuCode = this.toComparableString(record[dfuColumn]);
            if (dfuCode) {
                if (!groupedByDFU[dfuCode]) {
                    groupedByDFU[dfuCode] = [];
                }
                groupedByDFU[dfuCode].push(record);
            }
        });

        console.log('Grouped by DFU:', Object.keys(groupedByDFU).length, 'unique DFUs');

        const multiVariants = {};
        let multiVariantCount = 0;
        
        Object.keys(groupedByDFU).forEach(dfuCode => {
            const records = groupedByDFU[dfuCode];
            
            // Get unique part codes, ensuring we treat them as strings for consistency
            const uniquePartCodes = [...new Set(records.map(r => this.toComparableString(r[partNumberColumn])))].filter(Boolean);
            
            // Check if this DFU has completed transfers
            const isCompleted = this.completedTransfers[dfuCode];
            
            if (uniquePartCodes.length > 1 || isCompleted) {
                multiVariantCount++;
                const variantDemand = {};
                
                uniquePartCodes.forEach(partCode => {
                    // Filter records for this part code, ensuring string comparison
                    const partCodeRecords = records.filter(r => this.toComparableString(r[partNumberColumn]) === partCode);
                    
                    // Sum up all demand for this variant across all records
                    const totalDemand = partCodeRecords.reduce((sum, r) => {
                        const demand = parseFloat(r[demandColumn]) || 0;
                        return sum + demand;
                    }, 0);
                    
                    // Get part description from the first record
                    const partDescription = partCodeRecords[0] ? partCodeRecords[0][partDescriptionColumn] : '';
                    
                    // Include all variants that have records (including those with 0 demand)
                    if (partCodeRecords.length > 0) {
                        // Group records by week for granular control
                        const weeklyRecords = {};
                        partCodeRecords.forEach(record => {
                            const weekNum = this.toComparableString(record[weekNumberColumn]);
                            const demand = parseFloat(record[demandColumn]) || 0;
                            const sourceLocation = this.toComparableString(record[sourceLocationColumn]);
                            
                            const weekKey = `${weekNum}-${sourceLocation}`;
                            if (!weeklyRecords[weekKey]) {
                                weeklyRecords[weekKey] = {
                                    weekNumber: weekNum,
                                    sourceLocation: sourceLocation,
                                    demand: 0,
                                    records: []
                                };
                            }
                            weeklyRecords[weekKey].demand += demand;
                            weeklyRecords[weekKey].records.push(record);
                        });
                        
                        variantDemand[partCode] = {
                            totalDemand,
                            recordCount: partCodeRecords.length,
                            records: partCodeRecords,
                            partDescription: partDescription || 'Description not available',
                            weeklyRecords: weeklyRecords
                        };
                    }
                });
                
                // Always include DFUs that have completed transfers, even if they now have only one variant
                const activeVariants = Object.keys(variantDemand);
                
                // For completed transfers or multi-variant DFUs, include all variants (even with 0 demand)
                if (activeVariants.length > 1 || isCompleted || (isCompleted && this.keepZeroVariants)) {
                    multiVariants[dfuCode] = {
                        variants: activeVariants,
                        variantDemand,
                        totalRecords: records.length,
                        dfuColumn,
                        partNumberColumn,
                        demandColumn,
                        partDescriptionColumn,
                        plantLocationColumn,
                        calendarWeekColumn,
                        sourceLocationColumn,
                        weekNumberColumn,
                        isCompleted: !!isCompleted,
                        completionInfo: isCompleted || null,
                        plantLocation: records[0] ? this.toComparableString(records[0][plantLocationColumn]) : null
                    };
                    
                    console.log(`DFU ${dfuCode} variants after processing:`, activeVariants.map(v => ({
                        variant: v,
                        demand: variantDemand[v].totalDemand,
                        records: variantDemand[v].recordCount,
                        description: variantDemand[v].partDescription
                    })));
                }
            }
        });

        console.log('Multi-variant DFUs found:', multiVariantCount);

        this.multiVariantDFUs = multiVariants;
        this.filteredDFUs = multiVariants;
        
        if (multiVariantCount === 0) {
            this.showNotification('No DFU codes with multiple variants found in the data', 'error');
        } else {
            this.showNotification(`Found ${multiVariantCount} DFUs with multiple variants`);
        }
    }
    
    filterDFUs() {
        if (this.searchTerm) {
            const filtered = {};
            const searchLower = this.searchTerm.toLowerCase();
            Object.keys(this.multiVariantDFUs).forEach(dfuCode => {
                if (dfuCode.toLowerCase().includes(searchLower) ||
                    this.multiVariantDFUs[dfuCode].variants.some(v => 
                        v.toLowerCase().includes(searchLower))) {
                    filtered[dfuCode] = this.multiVariantDFUs[dfuCode];
                }
            });
            this.filteredDFUs = filtered;
        } else {
            this.filteredDFUs = this.multiVariantDFUs;
        }
        this.render();
        
        // Restore focus to search input after render
        setTimeout(() => {
            const searchInput = document.getElementById('searchInput');
            if (searchInput) {
                searchInput.focus();
                // Move cursor to end of input
                searchInput.setSelectionRange(searchInput.value.length, searchInput.value.length);
            }
        }, 10);
    }
    
    filterByPlantLocation(plantLocation) {
        console.log('Filtering by plant location:', plantLocation);
        this.selectedPlantLocation = plantLocation;
        
        // Clear existing data and re-process with filter
        this.multiVariantDFUs = {};
        this.filteredDFUs = {};
        
        // Re-process data with the new plant location filter
        this.processMultiVariantDFUs(this.rawData);
        this.render();
    }
    
    selectDFU(dfuCode) {
        // Ensure consistent type
        this.selectedDFU = this.toComparableString(dfuCode);
        this.render();
    }
    
    selectBulkTarget(dfuCode, targetVariant) {
        const dfuStr = this.toComparableString(dfuCode);
        const targetStr = this.toComparableString(targetVariant);
        
        this.bulkTransfers[dfuStr] = targetStr;
        // Clear individual transfers when bulk transfer is selected
        this.transfers[dfuStr] = {};
        this.render();
    }
    
    setIndividualTransfer(dfuCode, sourceVariant, targetVariant) {
        const dfuStr = this.toComparableString(dfuCode);
        const sourceStr = this.toComparableString(sourceVariant);
        const targetStr = this.toComparableString(targetVariant);
        
        if (!this.transfers[dfuStr]) {
            this.transfers[dfuStr] = {};
        }
        this.transfers[dfuStr][sourceStr] = targetStr;
        
        // Clear granular transfers when setting individual transfer
        if (this.granularTransfers[dfuStr] && this.granularTransfers[dfuStr][sourceStr]) {
            delete this.granularTransfers[dfuStr][sourceStr];
        }
        
        // Don't render immediately - let the caller handle it to preserve scroll
        if (targetVariant === '') {
            this.render();
        }
    }
    
    toggleGranularWeek(dfuCode, sourceVariant, targetVariant, weekKey) {
        const dfuStr = this.toComparableString(dfuCode);
        const sourceStr = this.toComparableString(sourceVariant);
        const targetStr = this.toComparableString(targetVariant);
        
        if (!this.granularTransfers[dfuStr]) {
            this.granularTransfers[dfuStr] = {};
        }
        if (!this.granularTransfers[dfuStr][sourceStr]) {
            this.granularTransfers[dfuStr][sourceStr] = {};
        }
        if (!this.granularTransfers[dfuStr][sourceStr][targetStr]) {
            this.granularTransfers[dfuStr][sourceStr][targetStr] = {};
        }
        
        // Toggle selection
        const current = this.granularTransfers[dfuStr][sourceStr][targetStr][weekKey];
        if (current && current.selected) {
            delete this.granularTransfers[dfuStr][sourceStr][targetStr][weekKey];
        } else {
            this.granularTransfers[dfuStr][sourceStr][targetStr][weekKey] = {
                selected: true,
                customQuantity: null // null means use full quantity
            };
        }
        
        // Clear individual transfer when granular is used
        if (this.transfers[dfuStr] && this.transfers[dfuStr][sourceStr]) {
            delete this.transfers[dfuStr][sourceStr];
        }
        
        // Update UI without full re-render to preserve scroll position
        this.updateActionButtonsOnly();
    }
    
    updateGranularQuantity(dfuCode, sourceVariant, targetVariant, weekKey, quantity) {
        const dfuStr = this.toComparableString(dfuCode);
        const sourceStr = this.toComparableString(sourceVariant);
        const targetStr = this.toComparableString(targetVariant);
        
        if (this.granularTransfers[dfuStr] && 
            this.granularTransfers[dfuStr][sourceStr] && 
            this.granularTransfers[dfuStr][sourceStr][targetStr] && 
            this.granularTransfers[dfuStr][sourceStr][targetStr][weekKey]) {
            
            this.granularTransfers[dfuStr][sourceStr][targetStr][weekKey].customQuantity = 
                quantity === '' ? null : parseFloat(quantity);
            
            // Update action buttons only
            this.updateActionButtonsOnly();
        }
    }
    
    updateActionButtonsOnly() {
        // Only update the action buttons section without full re-render
        const actionButtonsContainer = document.querySelector('.action-buttons-container');
        if (actionButtonsContainer && this.selectedDFU) {
            const hasTransfers = ((this.transfers[this.selectedDFU] && Object.keys(this.transfers[this.selectedDFU]).length > 0) || 
                                 this.bulkTransfers[this.selectedDFU] || 
                                 (this.granularTransfers[this.selectedDFU] && Object.keys(this.granularTransfers[this.selectedDFU]).length > 0));
            
            // Check if we have a recent execution summary to show
            const hasExecutionSummary = this.lastExecutionSummary[this.selectedDFU];
            
            if (hasTransfers) {
                actionButtonsContainer.innerHTML = `
                    <div class="p-3 bg-blue-50 rounded-lg">
                        <div class="text-sm text-blue-800 mb-3">
                            ${this.bulkTransfers[this.selectedDFU] ? `
                                <p><strong>Bulk Transfer:</strong> All variants → ${this.bulkTransfers[this.selectedDFU]}</p>
                            ` : ''}
                            ${this.transfers[this.selectedDFU] && Object.keys(this.transfers[this.selectedDFU]).length > 0 ? `
                                <p><strong>Individual Transfers:</strong></p>
                                <ul class="list-disc list-inside ml-4">
                                    ${Object.keys(this.transfers[this.selectedDFU]).map(sourceVariant => {
                                        const targetVariant = this.transfers[this.selectedDFU][sourceVariant];
                                        return sourceVariant !== targetVariant ? 
                                            `<li>${sourceVariant} → ${targetVariant}</li>` : '';
                                    }).filter(Boolean).join('')}
                                </ul>
                            ` : ''}
                            ${this.granularTransfers[this.selectedDFU] && Object.keys(this.granularTransfers[this.selectedDFU]).length > 0 ? `
                                <p><strong>Granular Transfers:</strong></p>
                                <ul class="list-disc list-inside ml-4 text-xs">
                                    ${Object.keys(this.granularTransfers[this.selectedDFU]).map(sourceVariant => {
                                        const sourceTransfers = this.granularTransfers[this.selectedDFU][sourceVariant];
                                        return Object.keys(sourceTransfers).map(targetVariant => {
                                            const weekTransfers = sourceTransfers[targetVariant];
                                            const weekCount = Object.keys(weekTransfers).length;
                                            return weekCount > 0 ? `<li>${sourceVariant} → ${targetVariant} (${weekCount} weeks)</li>` : '';
                                        }).filter(Boolean).join('');
                                    }).filter(Boolean).join('')}
                                </ul>
                            ` : ''}
                        </div>
                        <div class="flex gap-2">
                            <button class="btn btn-success" id="executeBtn">
                                <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 7l5 5m0 0l-5 5m5-5H6" />
                                </svg>
                                Execute Transfer
                            </button>
                            <button class="btn btn-secondary" id="cancelBtn">
                                <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12" />
                                </svg>
                                Cancel
                            </button>
                        </div>
                    </div>
                `;
                
                // Re-attach button event listeners
                const executeBtn = document.getElementById('executeBtn');
                if (executeBtn) {
                    executeBtn.addEventListener('click', () => this.executeTransfer(this.selectedDFU));
                }
                
                const cancelBtn = document.getElementById('cancelBtn');
                if (cancelBtn) {
                    cancelBtn.addEventListener('click', () => this.cancelTransfer(this.selectedDFU));
                }
            } else if (hasExecutionSummary) {
                // Show last execution summary
                const summary = this.lastExecutionSummary[this.selectedDFU];
                actionButtonsContainer.innerHTML = `
                    <div class="p-3 bg-gray-50 rounded-lg">
                        <div class="text-sm text-gray-700">
                            <h5 class="font-semibold mb-2 text-gray-800">Last Execution Summary:</h5>
                            <p><strong>Type:</strong> ${summary.type}</p>
                            <p><strong>Time:</strong> ${summary.timestamp}</p>
                            <p><strong>Result:</strong> ${summary.message}</p>
                            ${summary.details ? `
                                <div class="mt-2 text-xs">
                                    ${summary.details}
                                </div>
                            ` : ''}
                        </div>
                    </div>
                `;
            } else {
                actionButtonsContainer.innerHTML = '';
            }
        }
    }
    
    executeTransfer(dfuCode) {
        const dfuStr = this.toComparableString(dfuCode);
        const dfuData = this.multiVariantDFUs[dfuStr];
        const { dfuColumn, partNumberColumn, demandColumn, calendarWeekColumn, sourceLocationColumn, weekNumberColumn } = dfuData;
        
        // IMPORTANT: Store all original variants BEFORE any modifications
        const originalVariants = new Set();
        const dfuRecords = this.rawData.filter(record => this.toComparableString(record[dfuColumn]) === dfuStr);
        dfuRecords.forEach(record => {
            originalVariants.add(this.toComparableString(record[partNumberColumn]));
        });
        console.log('Original variants before transfer:', Array.from(originalVariants));
        
        let transferCount = 0;
        const transferHistory = []; // Track all transfers for audit trail
        const timestamp = new Date().toLocaleString('en-GB', { 
            day: '2-digit', 
            month: '2-digit', 
            year: 'numeric', 
            hour: '2-digit', 
            minute: '2-digit', 
            second: '2-digit' 
        });
        
        let executionType = '';
        let executionMessage = '';
        let executionDetails = '';
        
        // Handle bulk transfer
        if (this.bulkTransfers[dfuStr]) {
            const targetVariant = this.bulkTransfers[dfuStr];
            
            console.log(`Executing bulk transfer for DFU ${dfuStr} to ${targetVariant}`);
            console.log(`Found ${dfuRecords.length} records for this DFU`);
            
            dfuRecords.forEach(record => {
                const recordPartNumber = this.toComparableString(record[partNumberColumn]);
                if (recordPartNumber !== targetVariant) {
                    const sourceVariant = recordPartNumber;
                    const transferDemand = parseFloat(record[demandColumn]) || 0;
                    
                    const targetRecord = dfuRecords.find(r => 
                        this.toComparableString(r[partNumberColumn]) === targetVariant && 
                        this.toComparableString(r[calendarWeekColumn]) === this.toComparableString(record[calendarWeekColumn]) &&
                        this.toComparableString(r[sourceLocationColumn]) === this.toComparableString(record[sourceLocationColumn])
                    );
                    
                    if (targetRecord) {
                        const oldDemand = parseFloat(targetRecord[demandColumn]) || 0;
                        targetRecord[demandColumn] = oldDemand + transferDemand;
                        
                        // Add transfer history to target record
                        const existingHistory = targetRecord['Transfer History'] || '';
                        const newHistoryEntry = `[${sourceVariant} → ${transferDemand} @ ${timestamp}]`;
                        const pipoPrefix = existingHistory.startsWith('PIPO') ? '' : 'PIPO ';
                        targetRecord['Transfer History'] = existingHistory ? 
                            `${existingHistory} ${newHistoryEntry}` : `${pipoPrefix}${newHistoryEntry}`;
                        
                        record[demandColumn] = 0;
                        transferCount++;
                        
                        transferHistory.push({
                            from: sourceVariant,
                            to: targetVariant,
                            amount: transferDemand,
                            timestamp
                        });
                    } else {
                        // Change the source record to target variant
                        const originalVariant = recordPartNumber;
                        record[partNumberColumn] = isNaN(targetVariant) ? targetVariant : Number(targetVariant);
                        
                        // Add transfer history
                        record['Transfer History'] = `PIPO [${originalVariant} → ${transferDemand} @ ${timestamp}]`;
                        
                        transferCount++;
                        
                        transferHistory.push({
                            from: originalVariant,
                            to: targetVariant,
                            amount: transferDemand,
                            timestamp
                        });
                    }
                }
            });
            
            delete this.bulkTransfers[dfuStr];
            
            // Mark as completed transfer
            this.completedTransfers[dfuStr] = {
                type: 'bulk',
                targetVariant: targetVariant,
                timestamp: timestamp,
                originalVariantCount: dfuData.variants.length,
                transferHistory
            };
            
            executionType = 'Bulk Transfer';
            executionMessage = `${dfuData.variants.length - 1} variants transferred to ${targetVariant}`;
            executionDetails = `<p>All variants consolidated into: <strong>${targetVariant}</strong></p>`;
            
            this.showNotification(`Bulk transfer completed for DFU ${dfuStr}: ${executionMessage}`);
        }
        
        // Handle individual transfers
        else if (this.transfers[dfuStr] && Object.keys(this.transfers[dfuStr]).length > 0) {
            const individualTransfers = this.transfers[dfuStr];
            
            console.log(`Executing individual transfers for DFU ${dfuStr}`);
            console.log(`Individual transfers:`, individualTransfers);
            console.log(`Found ${dfuRecords.length} records for this DFU`);
            
            // Process each individual transfer
            Object.keys(individualTransfers).forEach(sourceVariant => {
                const targetVariant = individualTransfers[sourceVariant];
                
                console.log(`Processing transfer: ${sourceVariant} → ${targetVariant}`);
                
                // Only transfer if source and target are different
                if (sourceVariant !== targetVariant) {
                    // Find all records for this source variant
                    const sourceRecords = dfuRecords.filter(r => 
                        this.toComparableString(r[partNumberColumn]) === sourceVariant
                    );
                    
                    console.log(`Found ${sourceRecords.length} records for source variant ${sourceVariant}`);
                    
                    sourceRecords.forEach(record => {
                        const transferDemand = parseFloat(record[demandColumn]) || 0;
                        
                        // Try to find a matching target record with same week and location
                        const targetRecord = dfuRecords.find(r => 
                            this.toComparableString(r[partNumberColumn]) === targetVariant && 
                            this.toComparableString(r[calendarWeekColumn]) === this.toComparableString(record[calendarWeekColumn]) &&
                            this.toComparableString(r[sourceLocationColumn]) === this.toComparableString(record[sourceLocationColumn])
                        );
                        
                        if (targetRecord) {
                            // Add to existing target record
                            const oldDemand = parseFloat(targetRecord[demandColumn]) || 0;
                            targetRecord[demandColumn] = oldDemand + transferDemand;
                            
                            // Add transfer history to target record
                            const existingHistory = targetRecord['Transfer History'] || '';
                            const newHistoryEntry = `[${sourceVariant} → ${transferDemand} @ ${timestamp}]`;
                            const pipoPrefix = existingHistory.startsWith('PIPO') ? '' : 'PIPO ';
                            targetRecord['Transfer History'] = existingHistory ? 
                                `${existingHistory} ${newHistoryEntry}` : `${pipoPrefix}${newHistoryEntry}`;
                            
                            record[demandColumn] = 0; // Zero out source
                            console.log(`Added ${transferDemand} demand to existing target record`);
                        } else {
                            // Change the source record to target variant
                            const originalVariant = this.toComparableString(record[partNumberColumn]);
                            record[partNumberColumn] = isNaN(targetVariant) ? targetVariant : Number(targetVariant);
                            
                            // Add transfer history
                            record['Transfer History'] = `PIPO [${originalVariant} → ${transferDemand} @ ${timestamp}]`;
                            
                            console.log(`Changed record part number from ${sourceVariant} to ${targetVariant}`);
                        }
                        
                        transferHistory.push({
                            from: sourceVariant,
                            to: targetVariant,
                            amount: transferDemand,
                            timestamp
                        });
                    });
                    
                    transferCount++;
                }
            });
            
            this.transfers[dfuStr] = {};
            
            // Mark as completed transfer
            this.completedTransfers[dfuStr] = {
                type: 'individual',
                transfers: individualTransfers,
                timestamp: timestamp,
                transferCount: transferCount,
                transferHistory
            };
            
            executionType = 'Individual Transfers';
            executionMessage = `${transferCount} variant transfers executed`;
            executionDetails = `<ul class="list-disc list-inside ml-4">
                ${Object.keys(individualTransfers).map(src => 
                    src !== individualTransfers[src] ? `<li>${src} → ${individualTransfers[src]}</li>` : ''
                ).filter(Boolean).join('')}
            </ul>`;
            
            this.showNotification(`Individual transfers completed for DFU ${dfuStr}: ${executionMessage}`);
        }
        
        // Handle granular transfers
        else if (this.granularTransfers[dfuStr] && Object.keys(this.granularTransfers[dfuStr]).length > 0) {
            const granularTransfers = this.granularTransfers[dfuStr];
            
            console.log(`Executing granular transfers for DFU ${dfuStr}`);
            console.log(`Granular transfers:`, granularTransfers);
            
            let granularTransferCount = 0;
            
            // Process each source variant's granular transfers
            Object.keys(granularTransfers).forEach(sourceVariant => {
                const sourceTargets = granularTransfers[sourceVariant];
                
                Object.keys(sourceTargets).forEach(targetVariant => {
                    const weekTransfers = sourceTargets[targetVariant];
                    
                    Object.keys(weekTransfers).forEach(weekKey => {
                        const weekTransfer = weekTransfers[weekKey];
                        if (!weekTransfer.selected) return;
                        
                        const [weekNumber, sourceLocation] = weekKey.split('-');
                        
                        // Find the specific source record for this week and location
                        const sourceRecord = dfuRecords.find(r => 
                            this.toComparableString(r[partNumberColumn]) === sourceVariant &&
                            this.toComparableString(r[weekNumberColumn]) === weekNumber &&
                            this.toComparableString(r[sourceLocationColumn]) === sourceLocation
                        );
                        
                        if (sourceRecord) {
                            const originalDemand = parseFloat(sourceRecord[demandColumn]) || 0;
                            const transferAmount = weekTransfer.customQuantity !== null ? 
                                weekTransfer.customQuantity : originalDemand;
                            
                            console.log(`Transferring ${transferAmount} from ${sourceVariant} to ${targetVariant} for week ${weekNumber}`);
                            
                            // Find matching target record
                            const targetRecord = dfuRecords.find(r => 
                                this.toComparableString(r[partNumberColumn]) === targetVariant && 
                                this.toComparableString(r[weekNumberColumn]) === weekNumber &&
                                this.toComparableString(r[sourceLocationColumn]) === sourceLocation
                            );
                            
                            if (targetRecord) {
                                // Add to existing target record
                                const oldDemand = parseFloat(targetRecord[demandColumn]) || 0;
                                targetRecord[demandColumn] = oldDemand + transferAmount;
                                
                                // Add transfer history with proper week number format
                                const existingHistory = targetRecord['Transfer History'] || '';
                                const newHistoryEntry = `[W${weekNumber} ${sourceVariant} → ${transferAmount} @ ${timestamp}]`;
                                const pipoPrefix = existingHistory.startsWith('PIPO') ? '' : 'PIPO ';
                                targetRecord['Transfer History'] = existingHistory ? 
                                    `${existingHistory} ${newHistoryEntry}` : `${pipoPrefix}${newHistoryEntry}`;
                                
                                // Update source record
                                sourceRecord[demandColumn] = originalDemand - transferAmount;
                                
                                // Add transfer history to source record if partial transfer
                                if (transferAmount < originalDemand) {
                                    const sourceExistingHistory = sourceRecord['Transfer History'] || '';
                                    const sourceHistoryEntry = `[W${weekNumber} ${transferAmount} transferred to ${targetVariant} @ ${timestamp}]`;
                                    const sourcePipoPrefix = sourceExistingHistory.startsWith('PIPO') ? '' : 'PIPO ';
                                    sourceRecord['Transfer History'] = sourceExistingHistory ? 
                                        `${sourceExistingHistory} ${sourceHistoryEntry}` : `${sourcePipoPrefix}${sourceHistoryEntry}`;
                                }
                                
                            } else {
                                // Create new record by modifying source
                                if (transferAmount === originalDemand) {
                                    // Transfer full amount - change part number
                                    const originalVariant = this.toComparableString(sourceRecord[partNumberColumn]);
                                    sourceRecord[partNumberColumn] = isNaN(targetVariant) ? targetVariant : Number(targetVariant);
                                    sourceRecord['Transfer History'] = `PIPO [W${weekNumber} ${originalVariant} → ${transferAmount} @ ${timestamp}]`;
                                } else {
                                    // Partial transfer - need to create new record and update source
                                    const newRecord = { ...sourceRecord };
                                    newRecord[partNumberColumn] = isNaN(targetVariant) ? targetVariant : Number(targetVariant);
                                    newRecord[demandColumn] = transferAmount;
                                    newRecord['Transfer History'] = `PIPO [W${weekNumber} ${sourceVariant} → ${transferAmount} @ ${timestamp}]`;
                                    
                                    // Update source record
                                    sourceRecord[demandColumn] = originalDemand - transferAmount;
                                    
                                    // Add transfer history to source record
                                    const sourceExistingHistory = sourceRecord['Transfer History'] || '';
                                    const sourceHistoryEntry = `[W${weekNumber} ${transferAmount} transferred to ${targetVariant} @ ${timestamp}]`;
                                    const sourcePipoPrefix = sourceExistingHistory.startsWith('PIPO') ? '' : 'PIPO ';
                                    sourceRecord['Transfer History'] = sourceExistingHistory ? 
                                        `${sourceExistingHistory} ${sourceHistoryEntry}` : `${sourcePipoPrefix}${sourceHistoryEntry}`;
                                    
                                    // Add new record
                                    this.rawData.push(newRecord);
                                }
                            }
                            
                            transferHistory.push({
                                from: sourceVariant,
                                to: targetVariant,
                                amount: transferAmount,
                                week: weekNumber,
                                timestamp
                            });
                            
                            granularTransferCount++;
                        }
                    });
                });
            });
            
            this.granularTransfers[dfuStr] = {};
            
            // Mark as completed transfer
            this.completedTransfers[dfuStr] = {
                type: 'granular',
                timestamp: timestamp,
                transferCount: granularTransferCount,
                transferHistory
            };
            
            executionType = 'Granular Transfers';
            executionMessage = `${granularTransferCount} week-level transfers executed`;
            executionDetails = `<p>Specific weeks transferred between variants</p>`;
            
            this.showNotification(`Granular transfers completed for DFU ${dfuStr}: ${executionMessage}`);
        }

        // Store execution summary
        this.lastExecutionSummary[dfuStr] = {
            type: executionType,
            timestamp: timestamp,
            message: executionMessage,
            details: executionDetails
        };

        // CRITICAL: Consolidate records FIRST before recalculating UI data
        console.log('Step 1: Consolidating records...');
        this.consolidateRecords(dfuStr, originalVariants);
        
        // THEN clear cached data and recalculate
        console.log('Step 2: Clearing cached data...');
        this.multiVariantDFUs = {};
        this.filteredDFUs = {};
        
        console.log('Step 3: Recalculating variant demands...');
        this.processMultiVariantDFUs(this.rawData);
        
        console.log('Step 4: Updating UI...');
        // Force complete UI refresh by clearing selection and re-rendering
        const currentSelection = this.selectedDFU;
        this.selectedDFU = null;
        
        // First render to clear old data
        this.render();
        
        // Restore selection and render again to show fresh data
        setTimeout(() => {
            console.log('Step 5: Restoring selection with fresh data...');
            this.selectedDFU = currentSelection;
            
            // Log the current DFU data to verify it's correct
            if (this.multiVariantDFUs[currentSelection]) {
                console.log('Fresh DFU data for UI:', this.multiVariantDFUs[currentSelection]);
                console.log('Fresh variant demand data:', this.multiVariantDFUs[currentSelection].variantDemand);
            }
            
            // Force a complete DOM rebuild for the selected DFU section
            this.forceUIRefresh();
            console.log('Transfer and UI update complete!');
        }, 300);
    }
    
    consolidateRecords(dfuCode, originalVariants = null) {
        const dfuStr = this.toComparableString(dfuCode);
        console.log(`Consolidating records for DFU ${dfuStr}`);
        
        // Get the column information from the current DFU data
        const currentDFUData = this.multiVariantDFUs[dfuStr] || Object.values(this.multiVariantDFUs)[0];
        if (!currentDFUData) {
            console.error('No DFU data available for consolidation');
            return;
        }
        
        const { dfuColumn, partNumberColumn, demandColumn, calendarWeekColumn, sourceLocationColumn, weekNumberColumn } = currentDFUData;
        
        // Get all records for this DFU
        const allRecords = this.rawData;
        const dfuRecords = allRecords.filter(record => this.toComparableString(record[dfuColumn]) === dfuStr);
        
        console.log(`Found ${dfuRecords.length} records for DFU ${dfuStr} before consolidation`);
        
        // Use provided original variants or extract from current records
        const allPartNumbers = originalVariants || new Set();
        if (!originalVariants) {
            dfuRecords.forEach(record => {
                allPartNumbers.add(this.toComparableString(record[partNumberColumn]));
            });
        }
        console.log('All part numbers to preserve:', Array.from(allPartNumbers));
        
        // Create a map of consolidated records
        const consolidatedMap = new Map();
        
        dfuRecords.forEach((record) => {
            const partNumber = this.toComparableString(record[partNumberColumn]);
            const calendarWeek = this.toComparableString(record[calendarWeekColumn]);
            const weekNumber = this.toComparableString(record[weekNumberColumn]);
            const sourceLocation = this.toComparableString(record[sourceLocationColumn]);
            const demand = parseFloat(record[demandColumn]) || 0;
            const transferHistory = record['Transfer History'] || '';
            
            // Create a unique key for this combination using weekNumber instead of calendarWeek
            const key = `${partNumber}|${weekNumber}|${sourceLocation}`;
            
            if (consolidatedMap.has(key)) {
                // Add to existing consolidated record
                const existing = consolidatedMap.get(key);
                existing[demandColumn] = (parseFloat(existing[demandColumn]) || 0) + demand;
                
                // Consolidate transfer histories
                if (transferHistory && existing['Transfer History']) {
                    existing['Transfer History'] = `${existing['Transfer History']} ${transferHistory}`;
                } else if (transferHistory) {
                    existing['Transfer History'] = transferHistory;
                }
                
                console.log(`Consolidated ${demand} into existing record for ${partNumber}, total now: ${existing[demandColumn]}`);
            } else {
                // Create new consolidated record
                const consolidatedRecord = { ...record };
                consolidatedRecord[demandColumn] = demand;
                if (transferHistory) {
                    consolidatedRecord['Transfer History'] = transferHistory;
                }
                consolidatedMap.set(key, consolidatedRecord);
            }
        });
        
        // If keepZeroVariants is true, ensure all original part numbers have at least one record
        if (this.keepZeroVariants) {
            allPartNumbers.forEach(partNumber => {
                // Check if this part number has any records in the consolidated map
                let hasRecord = false;
                consolidatedMap.forEach((record, key) => {
                    if (key.startsWith(`${partNumber}|`)) {
                        hasRecord = true;
                    }
                });
                
                // If no record exists for this part number, create one with 0 demand
                if (!hasRecord) {
                    // Use the first record as a template
                    const templateRecord = dfuRecords.find(r => this.toComparableString(r[partNumberColumn]) === partNumber) || dfuRecords[0];
                    if (templateRecord) {
                        const zeroRecord = { ...templateRecord };
                        // Preserve the original data type for Product Number
                        zeroRecord[partNumberColumn] = isNaN(partNumber) ? partNumber : Number(partNumber);
                        zeroRecord[demandColumn] = 0;
                        zeroRecord['Transfer History'] = 'All demand transferred';
                        
                        // Create a key for the first week/location
                        const key = `${partNumber}|${this.toComparableString(zeroRecord[weekNumberColumn])}|${this.toComparableString(zeroRecord[sourceLocationColumn])}`;
                        consolidatedMap.set(key, zeroRecord);
                        
                        console.log(`Created zero-demand record for ${partNumber} to keep variant visible`);
                    }
                }
            });
        }
        
        console.log(`Consolidated into ${consolidatedMap.size} unique records`);
        
        // Remove old DFU records from rawData
        this.rawData = this.rawData.filter(record => this.toComparableString(record[dfuColumn]) !== dfuStr);
        
        // Add consolidated records back to rawData
        consolidatedMap.forEach((record) => {
            this.rawData.push(record);
        });
        
        const newDfuRecords = this.rawData.filter(record => this.toComparableString(record[dfuColumn]) === dfuStr);
        console.log(`After consolidation: ${newDfuRecords.length} records for DFU ${dfuStr}`);
        
        // Log the consolidated variants
        const variantSummary = {};
        newDfuRecords.forEach(record => {
            const partNumber = this.toComparableString(record[partNumberColumn]);
            const demand = parseFloat(record[demandColumn]) || 0;
            
            if (!variantSummary[partNumber]) {
                variantSummary[partNumber] = { totalDemand: 0, recordCount: 0 };
            }
            variantSummary[partNumber].totalDemand += demand;
            variantSummary[partNumber].recordCount += 1;
        });
        
        console.log(`DFU ${dfuStr} variant summary after consolidation:`, variantSummary);
    }
    
    forceUIRefresh() {
        // Get the app container and force a complete re-render
        const app = document.getElementById('app');
        
        // Store current state
        const currentSearch = this.searchTerm;
        
        // Temporarily clear the container
        app.innerHTML = '<div class="max-w-6xl mx-auto p-6 bg-white min-h-screen"><div class="text-center py-12"><div class="loading-spinner mb-2"></div><p>Refreshing interface...</p></div></div>';
        
        // Force a short delay then rebuild
        setTimeout(() => {
            // Restore search term
            this.searchTerm = currentSearch;
            
            // Rebuild the entire interface
            this.render();
            
            console.log('Forced UI refresh complete - interface rebuilt from scratch');
        }, 100);
    }
    
    cancelTransfer(dfuCode) {
        const dfuStr = this.toComparableString(dfuCode);
        delete this.transfers[dfuStr];
        delete this.bulkTransfers[dfuStr];
        delete this.granularTransfers[dfuStr];
        this.render();
    }
    
    undoTransfer(dfuCode) {
        const dfuStr = this.toComparableString(dfuCode);
        console.log(`Undoing transfer for DFU ${dfuStr}`);
        
        // Remove from completed transfers
        delete this.completedTransfers[dfuStr];
        
        // Clear any current transfer settings
        delete this.transfers[dfuStr];
        delete this.bulkTransfers[dfuStr];
        delete this.granularTransfers[dfuStr];
        
        // Clear execution summary
        delete this.lastExecutionSummary[dfuStr];
        
        // Restore the original data from backup
        console.log('Restoring original data from backup...');
        this.rawData = JSON.parse(JSON.stringify(this.originalRawData));
        
        // Force recalculation of multi-variant DFUs
        this.multiVariantDFUs = {};
        this.filteredDFUs = {};
        
        // Re-process the data to show the variants again with original quantities
        this.processMultiVariantDFUs(this.rawData);
        
        this.showNotification(`Transfer undone for DFU ${dfuStr}. Original data restored with all variants and quantities.`, 'success');
        this.render();
    }
    
    exportData() {
        try {
            // Create a copy of the data with formatted dates
            const formattedData = this.rawData.map(record => {
                const formattedRecord = { ...record };
                
                // Format Calendar.week if it exists
                if (formattedRecord['Calendar.week']) {
                    // Convert the date string to just YYYY-MM-DD format
                    const dateValue = formattedRecord['Calendar.week'];
                    if (dateValue) {
                        // Handle both Date objects and ISO strings
                        const date = new Date(dateValue);
                        if (!isNaN(date.getTime())) {
                            // Format as YYYY-MM-DD
                            const year = date.getFullYear();
                            const month = String(date.getMonth() + 1).padStart(2, '0');
                            const day = String(date.getDate()).padStart(2, '0');
                            formattedRecord['Calendar.week'] = `${year}-${month}-${day}`;
                        }
                    }
                }
                
                return formattedRecord;
            });
            
            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.json_to_sheet(formattedData);
            XLSX.utils.book_append_sheet(wb, ws, 'Updated Demand');
            XLSX.writeFile(wb, 'Updated_Demand_Data.xlsx');
            this.showNotification('Data exported successfully');
        } catch (error) {
            console.error('Error exporting data:', error);
            this.showNotification('Error exporting data: ' + error.message, 'error');
        }
    }
    
    render() {
        const app = document.getElementById('app');
        
        if (!this.isProcessed) {
            app.innerHTML = `
                <div class="max-w-6xl mx-auto p-6 bg-white min-h-screen">
                    <div class="text-center py-12">
                        <div class="bg-blue-50 rounded-lg p-8 inline-block">
                            <div class="w-12 h-12 mb-4 mx-auto bg-blue-600 rounded-full flex items-center justify-center">
                                <svg class="w-6 h-6 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                                </svg>
                            </div>
                            <h2 class="text-xl font-semibold mb-2">Upload Demand Data</h2>
                            <p class="text-gray-600 mb-4">
                                Upload your Excel file with the new "Total Demand" format
                            </p>
                            
                            ${this.isLoading ? `
                                <div class="text-blue-600">
                                    <div class="loading-spinner mb-2"></div>
                                    <p>Processing file...</p>
                                </div>
                            ` : `
                                <div class="space-y-4">
                                    <div>
                                        <input type="file" accept=".xlsx,.xls" class="file-input" id="fileInput">
                                        <p class="text-sm text-gray-500 mt-2">
                                            Supported formats: .xlsx, .xls
                                        </p>
                                    </div>
                                    
                                    <div class="text-left text-sm text-gray-600 bg-gray-50 p-4 rounded-lg">
                                        <p class="font-medium mb-2">Expected columns in your Excel file:</p>
                                        <ul class="list-disc list-inside space-y-1">
                                            <li><strong>DFU</strong> - DFU codes</li>
                                            <li><strong>Product Number</strong> - Part/product codes</li>
                                            <li><strong>weekly fcst</strong> - Demand/forecast values</li>
                                            <li><strong>PartDescription</strong> - Product descriptions</li>
                                            <li><strong>Plant Location</strong> - Plant location codes</li>
                                            <li><strong>Week Number</strong> - Week number values</li>
                                            <li><strong>Source Location</strong> - Source location codes</li>
                                        </ul>
                                    </div>
                                    
                                    <div class="border-t pt-4 mt-4">
                                        <h3 class="text-sm font-medium text-gray-700 mb-2">Optional: Upload Variant Cycle Dates</h3>
                                        <input type="file" accept=".xlsx,.xls" class="file-input" id="cycleFileInput">
                                        <p class="text-xs text-gray-500 mt-1">
                                            Upload file with DFU, Part Code, SOS, and EOS columns
                                        </p>
                                    </div>
                                </div>
                            `}
                        </div>
                    </div>
                </div>
            `;
            
            if (!this.isLoading) {
                const fileInput = document.getElementById('fileInput');
                fileInput.addEventListener('change', (e) => this.handleFileUpload(e));
                
                const cycleFileInput = document.getElementById('cycleFileInput');
                if (cycleFileInput) {
                    cycleFileInput.addEventListener('change', (e) => {
                        const file = e.target.files[0];
                        if (file) {
                            this.loadVariantCycleData(file);
                        }
                    });
                }
            }
            
            return;
        }
        
        app.innerHTML = `
            <div class="max-w-6xl mx-auto p-6 bg-white min-h-screen">
                <div class="mb-6">
                    <div class="flex justify-between items-center">
                        <div>
                            <h1 class="text-3xl font-bold text-gray-800 mb-2">DFU Demand Transfer Management</h1>
                            <p class="text-gray-600">
                                Manage demand transfers for DFU codes with multiple variants. Found ${Object.keys(this.multiVariantDFUs).length} DFUs with multiple variants.
                            </p>
                        </div>
                        <div class="text-right text-xs text-gray-400">
                            <p>Version 2.18.0</p>
                            <p>Build: 2025-08-05-scrolling-fix</p>
                        </div>
                    </div>
                </div>

                <div class="flex gap-4 mb-6 flex-responsive">
                    <div class="relative flex-1">
                        <svg class="absolute left-3 top-3 h-4 w-4 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
                        </svg>
                        <input 
                            type="text" 
                            placeholder="Search DFU codes or part codes..." 
                            value="${this.searchTerm}"
                            class="search-input"
                            id="searchInput"
                        >
                    </div>
                    <div class="relative">
                        <select class="px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent" id="plantLocationFilter">
                            <option value="">All Plant Locations</option>
                            ${this.availablePlantLocations.map(location => `
                                <option value="${location}" ${this.selectedPlantLocation === location ? 'selected' : ''}>
                                    Plant ${location}
                                </option>
                            `).join('')}
                        </select>
                    </div>
                    ${this.hasVariantCycleData ? `
                        <span class="inline-flex items-center px-3 py-2 text-sm text-green-700 bg-green-100 rounded-lg">
                            <svg class="w-4 h-4 mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                            </svg>
                            Cycle Data Loaded
                        </span>
                    ` : `
                        <label class="btn btn-secondary cursor-pointer">
                            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 4v16m8-8H4" />
                            </svg>
                            Load Cycle Dates
                            <input type="file" accept=".xlsx,.xls" class="hidden" id="cycleFileInput">
                        </label>
                    `}
                    <button class="btn btn-success" id="exportBtn">
                        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                        </svg>
                        Export Updated Data
                    </button>
                </div>

                <div class="grid grid-cols-1 lg:grid-cols-2 gap-6 grid-responsive">
                    <div class="bg-gray-50 rounded-lg p-6">
                        <h3 class="font-semibold text-gray-800 mb-4 flex items-center gap-2">
                            <svg class="w-5 h-5 text-amber-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-2.5L13.732 4c-.77-.833-1.964-.833-2.732 0L4.082 16.5c-.77.833.192 2.5 1.732 2.5z" />
                            </svg>
                            DFUs Requiring Review (${Object.keys(this.filteredDFUs).length})
                        </h3>
                        <div class="relative" style="height: 580px;">
                            <div class="absolute inset-x-0 top-0 h-4 bg-gradient-to-b from-gray-50 to-transparent pointer-events-none z-10"></div>
                            <div class="absolute inset-x-0 bottom-0 h-4 bg-gradient-to-t from-gray-50 to-transparent pointer-events-none z-10"></div>
                            <div class="space-y-2 h-full overflow-y-auto pr-1 scrollbar-custom" style="padding-top: 16px; padding-bottom: 16px;">>
                                ${Object.keys(this.filteredDFUs).map(dfuCode => {
                                const dfuData = this.filteredDFUs[dfuCode];
                                if (!dfuData || !dfuData.variants) return '';
                                
                                return `
                                    <div class="dfu-card ${this.selectedDFU === dfuCode ? 'selected' : ''}" data-dfu="${dfuCode}">
                                        <div class="flex justify-between items-start">
                                            <div>
                                                <h4 class="font-medium text-gray-800">DFU: ${dfuCode}</h4>
                                                <p class="text-sm text-gray-600">
                                                    ${dfuData.plantLocation ? `Plant ${dfuData.plantLocation} • ` : ''}${dfuData.variants.length} variant${dfuData.variants.length > 1 ? 's' : ''}${dfuData.isCompleted ? ' (transfer completed)' : ''}
                                                </p>
                                            </div>
                                            <div class="text-right">
                                                ${dfuData.isCompleted ? `
                                                    <span class="inline-flex items-center gap-1 text-green-600 text-sm">
                                                        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7" />
                                                        </svg>
                                                        Done
                                                    </span>
                                                ` : (this.transfers[dfuCode] && Object.keys(this.transfers[dfuCode]).length > 0) || this.bulkTransfers[dfuCode] || (this.granularTransfers[dfuCode] && Object.keys(this.granularTransfers[dfuCode]).length > 0) ? `
                                                    <span class="inline-flex items-center gap-1 text-green-600 text-sm">
                                                        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                                                        </svg>
                                                        Ready
                                                    </span>
                                                ` : `
                                                    <span class="text-amber-600 text-sm">Pending</span>
                                                `}
                                            </div>
                                        </div>
                                    </div>
                                `;
                                }).join('')}
                            </div>
                        </div>
                    </div>

                    <div class="bg-white border border-gray-200 rounded-lg p-6">
                        ${this.selectedDFU && this.multiVariantDFUs[this.selectedDFU] ? `
                            <div>
                                <h3 class="font-semibold text-gray-800 mb-4">
                                    DFU: ${this.selectedDFU}${this.multiVariantDFUs[this.selectedDFU].plantLocation ? ` (Plant: ${this.multiVariantDFUs[this.selectedDFU].plantLocation})` : ''} - Variant Details
                                    ${this.multiVariantDFUs[this.selectedDFU].isCompleted ? `
                                        <span class="ml-2 px-2 py-1 text-xs bg-green-100 text-green-800 rounded-full">
                                            ✓ Transfer Complete
                                        </span>
                                    ` : ''}
                                </h3>
                                
                                ${this.multiVariantDFUs[this.selectedDFU].isCompleted ? `
                                    <!-- Completed Transfer Summary -->
                                    <div class="mb-6 p-4 bg-green-50 rounded-lg border border-green-200">
                                        <div class="flex justify-between items-start">
                                            <div class="flex-1">
                                                <h4 class="font-semibold text-green-800 mb-3">✓ Transfer Completed</h4>
                                                <div class="text-sm text-green-700">
                                                    <p><strong>Type:</strong> ${this.multiVariantDFUs[this.selectedDFU].completionInfo.type === 'bulk' ? 'Bulk Transfer' : this.multiVariantDFUs[this.selectedDFU].completionInfo.type === 'granular' ? 'Granular Transfer' : 'Individual Transfers'}</p>
                                                    <p><strong>Date:</strong> ${this.multiVariantDFUs[this.selectedDFU].completionInfo.timestamp}</p>
                                                    ${this.multiVariantDFUs[this.selectedDFU].completionInfo.type === 'bulk' ? `
                                                        <p><strong>Target Variant:</strong> ${this.multiVariantDFUs[this.selectedDFU].completionInfo.targetVariant}</p>
                                                        <p><strong>Variants Consolidated:</strong> ${this.multiVariantDFUs[this.selectedDFU].completionInfo.originalVariantCount - 1} → 1</p>
                                                    ` : this.multiVariantDFUs[this.selectedDFU].completionInfo.type === 'granular' ? `
                                                        <p><strong>Granular Transfers:</strong> ${this.multiVariantDFUs[this.selectedDFU].completionInfo.transferCount} week-level transfers completed</p>
                                                    ` : `
                                                        <p><strong>Individual Transfers:</strong> ${this.multiVariantDFUs[this.selectedDFU].completionInfo.transferCount} completed</p>
                                                    `}
                                                </div>
                                            </div>
                                            <button class="px-3 py-1 text-xs bg-orange-100 text-orange-800 rounded hover:bg-orange-200 transition-colors" 
                                                    id="undoTransferBtn"
                                                    title="Undo this transfer and allow modifications">
                                                <svg class="w-3 h-3 inline mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 10h10a8 8 0 018 8v2M3 10l6 6m-6-6l6-6"></path>
                                                </svg>
                                                Undo
                                            </button>
                                        </div>
                                    </div>
                                    
                                    <!-- Current Variant Status -->
                                    <div class="mb-6">
                                        <h4 class="font-semibold text-gray-800 mb-3">Current Variant Status</h4>
                                        <div class="space-y-3">
                                            ${this.multiVariantDFUs[this.selectedDFU].variants.map(variant => {
                                                const demandData = this.multiVariantDFUs[this.selectedDFU].variantDemand[variant];
                                                
                                                return `
                                                    <div class="border rounded-lg p-3 bg-white">
                                                        <div class="flex justify-between items-center">
                                                            <div class="flex-1">
                                                                <h5 class="font-medium text-gray-800">Part: ${variant}</h5>
                                                                <p class="text-xs text-gray-500 mb-1 max-w-md break-words">${demandData?.partDescription || 'Description not available'}</p>
                                                                <p class="text-sm text-gray-600">${demandData?.recordCount || 0} records</p>
                                                            </div>
                                                            <div class="text-right">
                                                                <p class="font-medium text-gray-800">${this.formatNumber(demandData?.totalDemand || 0)}</p>
                                                                <p class="text-sm text-gray-600">consolidated demand</p>
                                                            </div>
                                                        </div>
                                                    </div>
                                                `;
                                            }).join('')}
                                        </div>
                                    </div>
                                ` : `
                                    <!-- Bulk Transfer Section -->
                                    <div class="mb-6 p-4 bg-purple-50 rounded-lg border">
                                        <h4 class="font-semibold text-purple-800 mb-3">Bulk Transfer (All Variants → One Target)</h4>
                                        <p class="text-sm text-purple-600 mb-3">Transfer all variants to a single target variant:</p>
                                        <div class="flex flex-wrap gap-2">
                                            ${this.multiVariantDFUs[this.selectedDFU].variants.map(variant => {
                                                const isSelected = this.bulkTransfers[this.selectedDFU] === variant;
                                                return `
                                                    <button 
                                                        class="px-3 py-1 rounded-full text-sm font-medium transition-all ${isSelected ? 'bg-purple-600 text-white' : 'bg-purple-100 text-purple-800 hover:bg-purple-200'}"
                                                        data-bulk-target="${variant}"
                                                    >
                                                        ${variant}
                                                    </button>
                                                `;
                                            }).join('')}
                                        </div>
                                        ${this.bulkTransfers[this.selectedDFU] ? `
                                            <p class="text-sm text-purple-700 mt-2">
                                                → All variants will transfer to: <strong>${this.bulkTransfers[this.selectedDFU]}</strong>
                                            </p>
                                        ` : ''}
                                    </div>
                                    
                                    <!-- Individual Transfer Section -->
                                    <div class="mb-6">
                                        <h4 class="font-semibold text-gray-800 mb-3">Individual Transfers (Variant → Specific Target)</h4>
                                        <div class="space-y-4">
                                            ${this.multiVariantDFUs[this.selectedDFU].variants.map(variant => {
                                                const demandData = this.multiVariantDFUs[this.selectedDFU].variantDemand[variant];
                                                const currentTransfer = this.transfers[this.selectedDFU]?.[variant];
                                                const hasGranularTransfers = this.granularTransfers[this.selectedDFU] && 
                                                    this.granularTransfers[this.selectedDFU][variant] && 
                                                    Object.keys(this.granularTransfers[this.selectedDFU][variant]).length > 0;
                                                
                                                return `
                                                    <div class="border rounded-lg p-4 bg-gray-50">
                                                        <div class="flex justify-between items-center mb-3">
                                                            <div class="flex-1">
                                                                <h5 class="font-medium text-gray-800">Part: ${variant}</h5>
                                                                <p class="text-xs text-gray-500 mb-1 max-w-md break-words">${demandData?.partDescription || 'Description not available'}</p>
                                                                <p class="text-sm text-gray-600">${demandData?.recordCount || 0} records • ${this.formatNumber(demandData?.totalDemand || 0)} total demand</p>
                                                                ${(() => {
                                                                    const cycleData = this.getCycleDataForVariant(this.selectedDFU, variant);
                                                                    if (cycleData) {
                                                                        return `
                                                                            <div class="mt-1 text-xs space-y-0.5">
                                                                                <p class="text-blue-600"><strong>SOS:</strong> ${cycleData.sos}</p>
                                                                                <p class="text-red-600"><strong>EOS:</strong> ${cycleData.eos}</p>
                                                                                ${cycleData.comments ? `<p class="text-gray-600 italic"><strong>Comments:</strong> ${cycleData.comments}</p>` : ''}
                                                                            </div>
                                                                        `;
                                                                    }
                                                                    return '';
                                                                })()}
                                                            </div>
                                                        </div>
                                                        
                                                        <div class="flex items-center gap-2 text-sm mb-3">
                                                            <span class="text-gray-600">Transfer all to:</span>
                                                            <select class="px-2 py-1 border rounded text-sm" data-source-variant="${variant}" id="select-${variant}">
                                                                <option value="">Select target...</option>
                                                                ${this.multiVariantDFUs[this.selectedDFU].variants.map(targetVariant => `
                                                                    <option value="${targetVariant}" ${currentTransfer === targetVariant ? 'selected' : ''}>
                                                                        ${targetVariant}${targetVariant === variant ? ' (self)' : ''}
                                                                    </option>
                                                                `).join('')}
                                                            </select>
                                                            ${currentTransfer && currentTransfer !== variant ? `
                                                                <span class="text-green-600 text-sm">→ ${currentTransfer}</span>
                                                            ` : ''}
                                                        </div>
                                                        
                                                        <!-- Granular Week-Level Transfers - Only show if target selected -->
                                                        ${currentTransfer && currentTransfer !== variant ? `
                                                            <div class="border-t pt-3 mt-3" id="granular-${variant}">
                                                                <div class="space-y-2 max-h-40 overflow-y-auto">
                                                                    ${Object.keys(demandData?.weeklyRecords || {}).map(weekKey => {
                                                                        const weekData = demandData.weeklyRecords[weekKey];
                                                                        const isSelected = this.granularTransfers[this.selectedDFU] && 
                                                                            this.granularTransfers[this.selectedDFU][variant] && 
                                                                            this.granularTransfers[this.selectedDFU][variant][currentTransfer] && 
                                                                            this.granularTransfers[this.selectedDFU][variant][currentTransfer][weekKey] && 
                                                                            this.granularTransfers[this.selectedDFU][variant][currentTransfer][weekKey].selected;
                                                                        
                                                                        const customQty = isSelected ? 
                                                                            this.granularTransfers[this.selectedDFU][variant][currentTransfer][weekKey].customQuantity : null;
                                                                        
                                                                        return `
                                                                            <div class="bg-white rounded border p-2 text-xs">
                                                                                <div class="flex items-center gap-3">
                                                                                    <input type="checkbox" 
                                                                                           class="w-4 h-4" 
                                                                                           ${isSelected ? 'checked' : ''}
                                                                                           data-granular-toggle
                                                                                           data-dfu="${this.selectedDFU}"
                                                                                           data-source="${variant}"
                                                                                           data-target="${currentTransfer}"
                                                                                           data-week="${weekKey}"
                                                                                    >
                                                                                    <div class="flex-1">
                                                                                        <span class="font-medium">Week ${weekData.weekNumber} (Loc: ${weekData.sourceLocation})</span>
                                                                                        <span class="text-gray-600 ml-2">${this.formatNumber(weekData.demand)} demand</span>
                                                                                    </div>
                                                                                    <input type="number" 
                                                                                           class="w-20 px-2 py-1 text-xs border rounded" 
                                                                                           placeholder="${weekData.demand}"
                                                                                           value="${customQty !== null ? customQty : ''}"
                                                                                           ${!isSelected ? 'disabled' : ''}
                                                                                           data-granular-qty
                                                                                           data-dfu="${this.selectedDFU}"
                                                                                           data-source="${variant}"
                                                                                           data-target="${currentTransfer}"
                                                                                           data-week="${weekKey}"
                                                                                    >
                                                                                </div>
                                                                            </div>
                                                                        `;
                                                                    }).join('')}
                                                                </div>
                                                            </div>
                                                        ` : ''}
                                                    </div>
                                                `;
                                            }).join('')}
                                        </div>
                                    </div>
                                    
                                    <!-- Action Buttons Container -->
                                    <div class="action-buttons-container">
                                        ${((this.transfers[this.selectedDFU] && Object.keys(this.transfers[this.selectedDFU]).length > 0) || 
                                           this.bulkTransfers[this.selectedDFU] || 
                                           (this.granularTransfers[this.selectedDFU] && Object.keys(this.granularTransfers[this.selectedDFU]).length > 0)) ? `
                                            <div class="p-3 bg-blue-50 rounded-lg">
                                                <div class="text-sm text-blue-800 mb-3">
                                                    ${this.bulkTransfers[this.selectedDFU] ? `
                                                        <p><strong>Bulk Transfer:</strong> All variants → ${this.bulkTransfers[this.selectedDFU]}</p>
                                                    ` : ''}
                                                    ${this.transfers[this.selectedDFU] && Object.keys(this.transfers[this.selectedDFU]).length > 0 ? `
                                                        <p><strong>Individual Transfers:</strong></p>
                                                        <ul class="list-disc list-inside ml-4">
                                                            ${Object.keys(this.transfers[this.selectedDFU]).map(sourceVariant => {
                                                                const targetVariant = this.transfers[this.selectedDFU][sourceVariant];
                                                                return sourceVariant !== targetVariant ? 
                                                                    `<li>${sourceVariant} → ${targetVariant}</li>` : '';
                                                            }).filter(Boolean).join('')}
                                                        </ul>
                                                    ` : ''}
                                                    ${this.granularTransfers[this.selectedDFU] && Object.keys(this.granularTransfers[this.selectedDFU]).length > 0 ? `
                                                        <p><strong>Granular Transfers:</strong></p>
                                                        <ul class="list-disc list-inside ml-4 text-xs">
                                                            ${Object.keys(this.granularTransfers[this.selectedDFU]).map(sourceVariant => {
                                                                const sourceTransfers = this.granularTransfers[this.selectedDFU][sourceVariant];
                                                                return Object.keys(sourceTransfers).map(targetVariant => {
                                                                    const weekTransfers = sourceTransfers[targetVariant];
                                                                    const weekCount = Object.keys(weekTransfers).length;
                                                                    return weekCount > 0 ? `<li>${sourceVariant} → ${targetVariant} (${weekCount} weeks)</li>` : '';
                                                                }).filter(Boolean).join('');
                                                            }).filter(Boolean).join('')}
                                                        </ul>
                                                    ` : ''}
                                                </div>
                                                <div class="flex gap-2">
                                                    <button class="btn btn-success" id="executeBtn">
                                                        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 7l5 5m0 0l-5 5m5-5H6" />
                                                        </svg>
                                                        Execute Transfer
                                                    </button>
                                                    <button class="btn btn-secondary" id="cancelBtn">
                                                        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12" />
                                                        </svg>
                                                        Cancel
                                                    </button>
                                                </div>
                                            </div>
                                        ` : this.lastExecutionSummary[this.selectedDFU] ? `
                                            <div class="p-3 bg-gray-50 rounded-lg">
                                                <div class="text-sm text-gray-700">
                                                    <h5 class="font-semibold mb-2 text-gray-800">Last Execution Summary:</h5>
                                                    <p><strong>Type:</strong> ${this.lastExecutionSummary[this.selectedDFU].type}</p>
                                                    <p><strong>Time:</strong> ${this.lastExecutionSummary[this.selectedDFU].timestamp}</p>
                                                    <p><strong>Result:</strong> ${this.lastExecutionSummary[this.selectedDFU].message}</p>
                                                    ${this.lastExecutionSummary[this.selectedDFU].details ? `
                                                        <div class="mt-2 text-xs">
                                                            ${this.lastExecutionSummary[this.selectedDFU].details}
                                                        </div>
                                                    ` : ''}
                                                </div>
                                            </div>
                                        ` : ''}
                                    </div>
                                `}
                            </div>
                        ` : `
                            <div class="text-center py-12 text-gray-500">
                                Select a DFU from the list to view variant details
                            </div>
                        `}
                    </div>
                </div>

                <div class="mt-6 bg-blue-50 rounded-lg p-4">
                    <h3 class="font-semibold text-blue-800 mb-2">How to Use</h3>
                    <ul class="text-sm text-blue-700 space-y-1">
                        <li><strong>Bulk Transfer:</strong> Click a purple button to transfer all variants to that target</li>
                        <li><strong>Individual Transfer:</strong> Use dropdowns to specify where each variant should go</li>
                        <li><strong>Granular Transfer:</strong> Select specific weeks to transfer partial demand</li>
                        <li><strong>Execute:</strong> Click "Execute Transfer" to apply your chosen transfers</li>
                        <li><strong>Export:</strong> Export the updated data when you're done with all transfers</li>
                    </ul>
                </div>
            </div>
        `;
        
        this.attachEventListeners();
        
        // FORCE CREATE granular containers if they don't exist
        this.ensureGranularContainers();
        
        // Debug: Check if granular containers were created after DOM update
        setTimeout(() => {
            if (this.selectedDFU && this.multiVariantDFUs[this.selectedDFU]) {
                console.log('=== POST-RENDER GRANULAR CONTAINER CHECK ===');
                this.multiVariantDFUs[this.selectedDFU].variants.forEach(variant => {
                    const container = document.getElementById(`granular-${variant}`);
                    console.log(`Container granular-${variant}:`, container ? 'EXISTS' : 'MISSING');
                    if (container) {
                        console.log(`  - Class: ${container.className}`);
                        console.log(`  - Parent: ${container.parentElement?.tagName}`);
                    }
                });
                
                // Also check all elements with granular in their ID
                const allGranular = document.querySelectorAll('[id*="granular"]');
                console.log('All elements with "granular" in ID:', Array.from(allGranular).map(el => el.id));
            }
        }, 100);
    }
    
    ensureGranularContainers() {
        if (this.selectedDFU && this.multiVariantDFUs[this.selectedDFU]) {
            console.log('=== FORCE CREATING GRANULAR CONTAINERS ===');
            
            this.multiVariantDFUs[this.selectedDFU].variants.forEach(variant => {
                // Find the parent container for this variant
                const selectElement = document.querySelector(`[data-source-variant="${variant}"]`);
                
                if (selectElement) {
                    const parentDiv = selectElement.closest('.border.rounded-lg');
                    
                    if (parentDiv) {
                        // Check if granular container already exists
                        let granularContainer = document.getElementById(`granular-${variant}`);
                        
                        if (!granularContainer) {
                            console.log(`Creating missing granular container for ${variant}`);
                            
                            // Create the granular container
                            granularContainer = document.createElement('div');
                            granularContainer.id = `granular-${variant}`;
                            granularContainer.className = 'border-t pt-3 mt-3 granular-section';
                            granularContainer.style.minHeight = '10px';
                            
                            // Add it to the parent div
                            parentDiv.appendChild(granularContainer);
                            
                            console.log(`✓ Created granular-${variant}`);
                        } else {
                            console.log(`✓ granular-${variant} already exists`);
                        }
                    } else {
                        console.log(`✗ Could not find parent div for ${variant}`);
                    }
                } else {
                    console.log(`✗ Could not find select element for ${variant}`);
                }
            });
        }
    }
    
    attachEventListeners() {
        const searchInput = document.getElementById('searchInput');
        if (searchInput) {
            searchInput.addEventListener('input', (e) => {
                this.searchTerm = e.target.value;
                this.filterDFUs();
            });
        }

        const plantLocationFilter = document.getElementById('plantLocationFilter');
        if (plantLocationFilter) {
            plantLocationFilter.addEventListener('change', (e) => {
                this.filterByPlantLocation(e.target.value);
            });
        }
        
        const exportBtn = document.getElementById('exportBtn');
        if (exportBtn) {
            exportBtn.addEventListener('click', () => this.exportData());
        }
        
        const executeBtn = document.getElementById('executeBtn');
        if (executeBtn) {
            executeBtn.addEventListener('click', () => this.executeTransfer(this.selectedDFU));
        }
        
        const cancelBtn = document.getElementById('cancelBtn');
        if (cancelBtn) {
            cancelBtn.addEventListener('click', () => this.cancelTransfer(this.selectedDFU));
        }
        
        const undoTransferBtn = document.getElementById('undoTransferBtn');
        if (undoTransferBtn) {
            undoTransferBtn.addEventListener('click', () => this.undoTransfer(this.selectedDFU));
        }
        
        // Add cycle file input listener
        const cycleFileInput = document.getElementById('cycleFileInput');
        if (cycleFileInput) {
            cycleFileInput.addEventListener('change', (e) => {
                const file = e.target.files[0];
                if (file) {
                    console.log('Cycle file selected:', file.name);
                    this.loadVariantCycleData(file);
                }
            });
        }
        
        // DFU card click handlers
        document.querySelectorAll('.dfu-card').forEach(card => {
            card.addEventListener('click', (e) => {
                const dfuCode = e.currentTarget.dataset.dfu;
                this.selectDFU(dfuCode);
            });
        });
        
        // Bulk target selection handlers
        document.querySelectorAll('[data-bulk-target]').forEach(button => {
            button.addEventListener('click', (e) => {
                const targetVariant = e.target.dataset.bulkTarget;
                this.selectBulkTarget(this.selectedDFU, targetVariant);
            });
        });
        
        // Individual transfer dropdown handlers  
        document.querySelectorAll('[data-source-variant]').forEach(select => {
            select.addEventListener('change', (e) => {
                const sourceVariant = e.target.dataset.sourceVariant;
                const targetVariant = e.target.value;
                
                console.log(`Dropdown changed: ${sourceVariant} → ${targetVariant}`);
                
                if (targetVariant && targetVariant !== sourceVariant) {
                    // Set the individual transfer
                    this.setIndividualTransfer(this.selectedDFU, sourceVariant, targetVariant);
                    
                    // Show granular section for this variant
                    console.log(`All granular sections:`, document.querySelectorAll('[id^="granular-"]'));
                    const granularSection = document.getElementById(`granular-${sourceVariant}`);
                    console.log(`Looking for granular section: granular-${sourceVariant}`, granularSection);
                    
                    if (granularSection) {
                        const demandData = this.multiVariantDFUs[this.selectedDFU].variantDemand[sourceVariant];
                        console.log(`Demand data for ${sourceVariant}:`, demandData);
                        
                        if (demandData && demandData.weeklyRecords) {
                            console.log(`Weekly records:`, Object.keys(demandData.weeklyRecords));
                            granularSection.innerHTML = `
                                <div class="space-y-2 max-h-40 overflow-y-auto">
                                    ${Object.keys(demandData.weeklyRecords).map(weekKey => {
                                        const weekData = demandData.weeklyRecords[weekKey];
                                        const isSelected = this.granularTransfers[this.selectedDFU] && 
                                            this.granularTransfers[this.selectedDFU][sourceVariant] && 
                                            this.granularTransfers[this.selectedDFU][sourceVariant][targetVariant] && 
                                            this.granularTransfers[this.selectedDFU][sourceVariant][targetVariant][weekKey] && 
                                            this.granularTransfers[this.selectedDFU][sourceVariant][targetVariant][weekKey].selected;
                                        
                                        const customQty = isSelected ? 
                                            this.granularTransfers[this.selectedDFU][sourceVariant][targetVariant][weekKey].customQuantity : null;
                                        
                                        return `
                                            <div class="bg-white rounded border p-2 text-xs">
                                                <div class="flex items-center gap-3">
                                                    <input type="checkbox" 
                                                           class="w-4 h-4" 
                                                           ${isSelected ? 'checked' : ''}
                                                           data-granular-toggle
                                                           data-dfu="${this.selectedDFU}"
                                                           data-source="${sourceVariant}"
                                                           data-target="${targetVariant}"
                                                           data-week="${weekKey}"
                                                    >
                                                    <div class="flex-1">
                                                        <span class="font-medium">Week ${weekData.weekNumber} (Loc: ${weekData.sourceLocation})</span>
                                                        <span class="text-gray-600 ml-2">${this.formatNumber(weekData.demand)} demand</span>
                                                    </div>
                                                    <input type="number" 
                                                           class="w-20 px-2 py-1 text-xs border rounded" 
                                                           placeholder="${weekData.demand}"
                                                           value="${customQty !== null ? customQty : ''}"
                                                           ${!isSelected ? 'disabled' : ''}
                                                           data-granular-qty
                                                           data-dfu="${this.selectedDFU}"
                                                           data-source="${sourceVariant}"
                                                           data-target="${targetVariant}"
                                                           data-week="${weekKey}"
                                                    >
                                                </div>
                                            </div>
                                        `;
                                    }).join('')}
                                </div>
                            `;
                            
                            console.log(`Updated granular section for ${sourceVariant}`);
                            
                            // Re-attach event listeners for the new granular controls
                            setTimeout(() => {
                                this.attachGranularEventListeners();
                            }, 100);
                        } else {
                            console.log(`No weekly records found for ${sourceVariant}`);
                            granularSection.innerHTML = `<p class="text-gray-500 text-sm">No weekly data available for granular transfers.</p>`;
                        }
                    } else {
                        console.error(`Granular section not found: granular-${sourceVariant}`);
                        console.log('Available elements with granular IDs:', 
                            Array.from(document.querySelectorAll('[id*="granular"]')).map(el => el.id));
                    }
                    
                    // Update action buttons
                    this.updateActionButtonsOnly();
                    
                } else {
                    // Remove transfer if empty selection or self-selection
                    if (this.transfers[this.selectedDFU]) {
                        delete this.transfers[this.selectedDFU][sourceVariant];
                    }
                    
                    // Clear granular section
                    const granularSection = document.getElementById(`granular-${sourceVariant}`);
                    if (granularSection) {
                        granularSection.innerHTML = '';
                    }
                    
                    this.updateActionButtonsOnly();
                }
            });
        });
        
        // Attach initial granular event listeners
        this.attachGranularEventListeners();
    }
    
    attachGranularEventListeners() {
        console.log('Attaching granular event listeners...');
        
        // Granular transfer checkbox handlers
        document.querySelectorAll('[data-granular-toggle]').forEach(checkbox => {
            // Remove existing listeners to avoid duplicates
            const newCheckbox = checkbox.cloneNode(true);
            checkbox.parentNode.replaceChild(newCheckbox, checkbox);
            
            newCheckbox.addEventListener('change', (e) => {
                const dfuCode = e.target.dataset.dfu;
                const sourceVariant = e.target.dataset.source;
                const targetVariant = e.target.dataset.target;
                const weekKey = e.target.dataset.week;
                
                console.log(`Checkbox toggled: ${sourceVariant} → ${targetVariant} for week ${weekKey}`);
                
                this.toggleGranularWeek(dfuCode, sourceVariant, targetVariant, weekKey);
                
                // Update the quantity input state
                const qtyInput = document.querySelector(`[data-granular-qty][data-week="${weekKey}"][data-source="${sourceVariant}"][data-target="${targetVariant}"]`);
                if (qtyInput) {
                    qtyInput.disabled = !e.target.checked;
                }
            });
        });
        
        // Granular transfer quantity handlers
        document.querySelectorAll('[data-granular-qty]').forEach(input => {
            // Remove existing listeners
            const newInput = input.cloneNode(true);
            input.parentNode.replaceChild(newInput, input);
            
            newInput.addEventListener('input', (e) => {
                const dfuCode = e.target.dataset.dfu;
                const sourceVariant = e.target.dataset.source;
                const targetVariant = e.target.dataset.target;
                const weekKey = e.target.dataset.week;
                const quantity = e.target.value;
                
                console.log(`Quantity updated: ${quantity} for ${sourceVariant} → ${targetVariant} week ${weekKey}`);
                
                this.updateGranularQuantity(dfuCode, sourceVariant, targetVariant, weekKey, quantity);
            });
        });
        
        console.log(`Attached listeners to ${document.querySelectorAll('[data-granular-toggle]').length} checkboxes and ${document.querySelectorAll('[data-granular-qty]').length} quantity inputs`);
    }
}

// Initialize the application when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    new DemandTransferApp();
});
