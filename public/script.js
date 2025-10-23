// DFU Demand Transfer Management Application
// Version: 2.21.0 - Build: 2025-10-23-network-stock-calc
// Added Network Stock + Total Demand and Selected Demand calculation boxes

class DemandTransferApp {
    constructor() {
        this.rawData = [];
        this.originalRawData = [];
        this.multiVariantDFUs = {};
        this.filteredDFUs = {};
        this.selectedDFU = null;
        this.searchTerm = '';
        this.selectedPlantLocation = '';
        this.selectedProductionLine = '';
        this.availablePlantLocations = [];
        this.availableProductionLines = [];
        this.transfers = {};
        this.bulkTransfers = {};
        this.granularTransfers = {};
        this.completedTransfers = {};
        this.isProcessed = false;
        this.isLoading = false;
        this.lastExecutionSummary = {};
        this.variantCycleDates = {};
        this.hasVariantCycleData = false;
        this.keepZeroVariants = true;
        this.searchDebounceTimer = null;
        
        // Supply chain data storage
        this.stockData = {}; // { partCode: sohValue }
        this.openSupplyData = {}; // { partCode: openSupplyValue }
        this.transitData = {}; // { partCode: transitValue }
        this.hasStockData = false;
        this.hasOpenSupplyData = false;
        this.hasTransitData = false;
        
        this.init();
    }
    
    init() {
        console.log('🚀 DFU Demand Transfer App v2.21.0 - Build: 2025-10-23-network-stock-calc');
        console.log('📊 Added Network Stock + Total Demand and Selected Demand calculations');
        this.render();
        this.attachEventListeners();
    }
    
    toComparableString(value) {
        if (value === null || value === undefined) return '';
        return String(value).trim();
    }
    
    formatNumber(num) {
        if (num === null || num === undefined || isNaN(num)) return '0';
        return Math.round(num).toLocaleString('en-US');
    }
    
    getDateFromWeekNumber(year, weekNumber) {
        const jan1 = new Date(year, 0, 1);
        const jan1DayOfWeek = jan1.getDay();
        const daysToFirstMonday = jan1DayOfWeek === 0 ? 1 : (8 - jan1DayOfWeek);
        const firstMonday = new Date(year, 0, 1 + daysToFirstMonday);
        const targetDate = new Date(firstMonday.getTime() + (weekNumber - 1) * 7 * 24 * 60 * 60 * 1000);
        return targetDate;
    }
    
    // Load Stock (SOH) file
    async loadStockFile(file) {
        console.log('Loading Stock (SOH) file...');
        try {
            const arrayBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(worksheet);
            
            console.log(`Stock file loaded: ${data.length} records`);
            
            this.stockData = {};
            data.forEach(row => {
                const partCode = this.toComparableString(row['Product Number'] || row['ProductNumber'] || row['PartCode'] || row['Part Code']);
                const stock = parseFloat(row['Stock'] || row['SOH'] || 0);
                
                if (partCode) {
                    if (this.stockData[partCode]) {
                        this.stockData[partCode] += stock;
                    } else {
                        this.stockData[partCode] = stock;
                    }
                }
            });
            
            this.hasStockData = true;
            console.log('Stock data processed:', Object.keys(this.stockData).length, 'unique parts');
            this.showNotification('Stock (SOH) data loaded successfully', 'success');
            this.render();
        } catch (error) {
            console.error('Error loading stock file:', error);
            this.showNotification('Error loading stock file: ' + error.message, 'error');
        }
    }
    
    // Load Open Supply file
    async loadOpenSupplyFile(file) {
        console.log('Loading Open Supply file...');
        try {
            const arrayBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(worksheet);
            
            console.log(`Open Supply file loaded: ${data.length} records`);
            
            // Debug: Show available columns
            if (data.length > 0) {
                console.log('Available columns in Open Supply file:', Object.keys(data[0]));
                console.log('Sample row:', data[0]);
            }
            
            this.openSupplyData = {};
            let processedCount = 0;
            
            data.forEach(row => {
                // Try multiple column name variations for part code
                const partCode = this.toComparableString(
                    row['Product Number'] || 
                    row['ProductNumber'] || 
                    row['PartCode'] || 
                    row['Part Code'] ||
                    row['Material'] ||
                    row['Material Number']
                );
                
                // Try multiple column name variations for quantity
                const openSupply = parseFloat(
                    row['Receipt Quantity / Requirements Quantity'] ||  // EXACT column name from your file
                    row['Receipt Quantity'] || 
                    row['Requirements Quantity'] || 
                    row['ReceiptQuantity'] ||
                    row['RequirementsQuantity'] ||
                    row['OpenSupply'] || 
                    row['Open Supply'] ||
                    row['Quantity'] ||
                    row['Qty'] ||
                    row['Order Quantity'] ||
                    row['OrderQuantity'] ||
                    0
                );
                
                if (partCode && openSupply !== 0) {
                    if (this.openSupplyData[partCode]) {
                        this.openSupplyData[partCode] += openSupply;
                    } else {
                        this.openSupplyData[partCode] = openSupply;
                    }
                    processedCount++;
                }
            });
            
            this.hasOpenSupplyData = true;
            console.log('Open Supply data processed:', Object.keys(this.openSupplyData).length, 'unique parts');
            console.log('Total records with quantity:', processedCount);
            console.log('Sample processed data:', Object.entries(this.openSupplyData).slice(0, 3));
            
            this.showNotification(`Open Supply data loaded: ${processedCount} records processed`, 'success');
            this.render();
        } catch (error) {
            console.error('Error loading Open Supply file:', error);
            this.showNotification('Error loading Open Supply file: ' + error.message, 'error');
        }
    }
    
    // Load Stock in Transit file
    async loadTransitFile(file) {
        console.log('Loading Stock in Transit file...');
        try {
            const arrayBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(worksheet);
            
            console.log(`Stock in Transit file loaded: ${data.length} records`);
            
            this.transitData = {};
            data.forEach(row => {
                const partCode = this.toComparableString(row['PartCode'] || row['Part Code'] || row['Product Number'] || row['ProductNumber']);
                const transit = parseFloat(row['SupplierOrderQuantityOrdered'] || row['TransitQuantity'] || row['InTransit'] || 0);
                
                if (partCode) {
                    if (this.transitData[partCode]) {
                        this.transitData[partCode] += transit;
                    } else {
                        this.transitData[partCode] = transit;
                    }
                }
            });
            
            this.hasTransitData = true;
            console.log('Stock in Transit data processed:', Object.keys(this.transitData).length, 'unique parts');
            this.showNotification('Stock in Transit data loaded successfully', 'success');
            this.render();
        } catch (error) {
            console.error('Error loading Transit file:', error);
            this.showNotification('Error loading Transit file: ' + error.message, 'error');
        }
    }
    
    // Get supply chain data for a variant
    getSupplyChainData(partCode) {
        const partStr = this.toComparableString(partCode);
        const soh = this.stockData[partStr] || 0;
        const openSupply = this.openSupplyData[partStr] || 0;
        const transit = this.transitData[partStr] || 0;
        const total = soh + openSupply + transit;
        
        return {
            soh,
            openSupply,
            transit,
            total,
            hasData: this.hasStockData || this.hasOpenSupplyData || this.hasTransitData
        };
    }
    
    // Calculate Network Stock + Total Demand for a variant
    calculateNetworkStockPlusDemand(partCode) {
        if (!this.selectedDFU) return null;
        
        const demandData = this.multiVariantDFUs[this.selectedDFU]?.variantDemand[partCode];
        if (!demandData) return null;
        
        const supplyChain = this.getSupplyChainData(partCode);
        const totalDemand = demandData.totalDemand || 0;
        const networkStock = supplyChain.total;
        const result = networkStock + totalDemand;
        
        return {
            networkStock,
            totalDemand,
            result,
            hasSupplyData: supplyChain.hasData
        };
    }
    
    // Calculate selected demand for granular transfers
    calculateSelectedDemand(sourceVariant) {
        if (!this.selectedDFU || !sourceVariant) return 0;
        
        const granularData = this.granularTransfers[this.selectedDFU]?.[sourceVariant];
        if (!granularData) return 0;
        
        let totalSelected = 0;
        
        Object.values(granularData).forEach(targetData => {
            Object.values(targetData).forEach(weekData => {
                if (weekData.selected) {
                    const quantity = weekData.customQuantity !== null && weekData.customQuantity !== '' 
                        ? parseFloat(weekData.customQuantity) 
                        : weekData.originalQuantity;
                    totalSelected += quantity || 0;
                }
            });
        });
        
        return totalSelected;
    }
    
    showNotification(message, type = 'info') {
        const container = document.getElementById('notifications');
        if (!container) return;
        
        const notification = document.createElement('div');
        notification.className = `notification notification-${type} notification-enter`;
        
        const icon = type === 'success' ? '✓' : type === 'error' ? '✕' : 'ℹ';
        const bgColor = type === 'success' ? 'bg-green-500' : type === 'error' ? 'bg-red-500' : 'bg-blue-500';
        
        notification.innerHTML = `
            <div class="${bgColor} text-white px-4 py-3 rounded-lg shadow-lg flex items-center gap-3">
                <span class="text-xl font-bold">${icon}</span>
                <span>${message}</span>
            </div>
        `;
        
        container.appendChild(notification);
        
        setTimeout(() => {
            notification.remove();
        }, 4000);
    }
    
    async loadVariantCycleData(file) {
        console.log('Loading variant cycle data file...');
        try {
            const arrayBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(worksheet);
            
            console.log(`Cycle data file loaded: ${data.length} records`);
            
            this.variantCycleDates = {};
            let processedCount = 0;
            
            data.forEach(row => {
                const dfuCode = this.toComparableString(row['DFU'] || row['DFU Code']);
                const partCode = this.toComparableString(row['Part Code'] || row['PartCode'] || row['Product Number']);
                const sos = row['SOS'] || row['Start of Supply'] || '';
                const eos = row['EOS'] || row['End of Supply'] || '';
                const comments = row['Comments'] || row['Comment'] || '';
                
                if (dfuCode && partCode) {
                    const dfuStr = dfuCode;
                    const partStr = partCode;
                    
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
            
            this.hasVariantCycleData = true;
            console.log(`Processed ${processedCount} cycle data records`);
            this.showNotification('Variant cycle data loaded successfully', 'success');
            this.render();
        } catch (error) {
            console.error('Error loading cycle data:', error);
            this.showNotification('Error loading cycle data: ' + error.message, 'error');
        }
    }
    
    getCycleDataForVariant(dfuCode, partCode) {
        const dfuStr = this.toComparableString(dfuCode);
        const partStr = this.toComparableString(partCode);
        
        if (this.variantCycleDates[dfuStr] && this.variantCycleDates[dfuStr][partStr]) {
            return this.variantCycleDates[dfuStr][partStr];
        }
        return null;
    }
    
    handleFileUpload(event) {
        const file = event.target.files[0];
        if (!file) return;
        
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
            const workbook = XLSX.read(arrayBuffer, { 
                cellStyles: true, 
                cellFormulas: true, 
                cellDates: true,
                cellNF: true,
                sheetStubs: true
            });
            
            let sheetName = 'Total Demand';
            if (!workbook.Sheets[sheetName]) {
                const possibleNames = ['Open Fcst', 'Demand', 'Sheet1'];
                sheetName = possibleNames.find(name => workbook.Sheets[name]) || workbook.SheetNames[0];
            }
            
            const worksheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(worksheet);
            
            if (data.length > 0) {
                this.rawData = data;
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
        if (data.length === 0) return;
        
        const dfuColumn = 'DFU';
        const partNumberColumn = 'Product Number';
        const demandColumn = 'weekly fcst';
        const partDescriptionColumn = 'PartDescription';
        const plantLocationColumn = 'Production Plant';
        const productionLineColumn = 'Production Line';
        
        const uniquePlants = new Set();
        const uniqueLines = new Set();
        
        data.forEach(record => {
            const plant = this.toComparableString(record[plantLocationColumn]);
            const line = this.toComparableString(record[productionLineColumn]);
            if (plant) uniquePlants.add(plant);
            if (line) uniqueLines.add(line);
        });
        
        this.availablePlantLocations = Array.from(uniquePlants).sort();
        this.availableProductionLines = Array.from(uniqueLines).sort();
        
        let filteredData = data;
        if (this.selectedPlantLocation) {
            filteredData = filteredData.filter(r => 
                this.toComparableString(r[plantLocationColumn]) === this.selectedPlantLocation
            );
        }
        if (this.selectedProductionLine) {
            filteredData = filteredData.filter(r => 
                this.toComparableString(r[productionLineColumn]) === this.selectedProductionLine
            );
        }
        
        const groupedByDFU = {};
        filteredData.forEach(record => {
            const dfuCode = this.toComparableString(record[dfuColumn]);
            if (dfuCode) {
                if (!groupedByDFU[dfuCode]) {
                    groupedByDFU[dfuCode] = [];
                }
                groupedByDFU[dfuCode].push(record);
            }
        });
        
        const allDFUs = {};
        Object.keys(groupedByDFU).forEach(dfuCode => {
            const records = groupedByDFU[dfuCode];
            const uniquePartCodes = [...new Set(records.map(r => this.toComparableString(r[partNumberColumn])))].filter(Boolean);
            const uniquePlants = [...new Set(records.map(r => this.toComparableString(r[plantLocationColumn])))].filter(Boolean);
            const uniqueProductionLines = [...new Set(records.map(r => this.toComparableString(r[productionLineColumn])))].filter(Boolean);
            const isCompleted = this.completedTransfers[dfuCode];
            
            const variantDemand = {};
            uniquePartCodes.forEach(partCode => {
                const partCodeRecords = records.filter(r => this.toComparableString(r[partNumberColumn]) === partCode);
                const totalDemand = partCodeRecords.reduce((sum, r) => sum + (parseFloat(r[demandColumn]) || 0), 0);
                const partDescription = partCodeRecords[0] ? (partCodeRecords[0][partDescriptionColumn] || '') : '';
                
                variantDemand[partCode] = {
                    totalDemand,
                    recordCount: partCodeRecords.length,
                    partDescription
                };
            });
            
            allDFUs[dfuCode] = {
                variants: uniquePartCodes,
                recordCount: records.length,
                variantDemand,
                isSingleVariant: uniquePartCodes.length === 1,
                plantLocations: uniquePlants,
                productionLines: uniqueProductionLines,
                isCompleted: isCompleted ? true : false
            };
            
            if (isCompleted) {
                allDFUs[dfuCode].completionInfo = this.completedTransfers[dfuCode];
            }
        });
        
        this.multiVariantDFUs = allDFUs;
        this.applySearchFilter();
    }
    
    applySearchFilter() {
        const searchLower = this.searchTerm.toLowerCase();
        this.filteredDFUs = {};
        
        Object.keys(this.multiVariantDFUs).forEach(dfuCode => {
            const dfuData = this.multiVariantDFUs[dfuCode];
            const dfuMatches = dfuCode.toLowerCase().includes(searchLower);
            const variantMatches = dfuData.variants.some(v => v.toLowerCase().includes(searchLower));
            
            if (dfuMatches || variantMatches || searchLower === '') {
                this.filteredDFUs[dfuCode] = dfuData;
            }
        });
    }
    
    selectDFU(dfuCode) {
        this.selectedDFU = dfuCode;
        this.render();
        setTimeout(() => {
            this.ensureGranularContainers();
        }, 100);
    }
    
    selectBulkTarget(dfuCode, targetVariant) {
        this.bulkTransfers[dfuCode] = targetVariant;
        this.render();
    }
    
    setIndividualTransfer(dfuCode, sourceVariant, targetVariant) {
        if (!targetVariant || targetVariant === '' || targetVariant === sourceVariant) {
            // Clear the transfer if empty or self-transfer
            if (this.transfers[dfuCode]) {
                delete this.transfers[dfuCode][sourceVariant];
                
                // Clean up empty objects
                if (Object.keys(this.transfers[dfuCode]).length === 0) {
                    delete this.transfers[dfuCode];
                }
            }
            
            // Clear granular transfers for this variant
            if (this.granularTransfers[dfuCode]?.[sourceVariant]) {
                delete this.granularTransfers[dfuCode][sourceVariant];
                
                // Clean up empty objects
                if (Object.keys(this.granularTransfers[dfuCode]).length === 0) {
                    delete this.granularTransfers[dfuCode];
                }
            }
        } else {
            // Set the transfer
            if (!this.transfers[dfuCode]) {
                this.transfers[dfuCode] = {};
            }
            this.transfers[dfuCode][sourceVariant] = targetVariant;
            
            // Initialize granular transfers structure
            if (!this.granularTransfers[dfuCode]) {
                this.granularTransfers[dfuCode] = {};
            }
            if (!this.granularTransfers[dfuCode][sourceVariant]) {
                this.granularTransfers[dfuCode][sourceVariant] = {};
            }
            if (!this.granularTransfers[dfuCode][sourceVariant][targetVariant]) {
                this.granularTransfers[dfuCode][sourceVariant][targetVariant] = {};
            }
        }
        
        this.render();
        setTimeout(() => {
            this.ensureGranularContainers();
        }, 100);
    }
    
    toggleGranularWeek(dfuCode, sourceVariant, targetVariant, weekKey) {
        if (!this.granularTransfers[dfuCode]) this.granularTransfers[dfuCode] = {};
        if (!this.granularTransfers[dfuCode][sourceVariant]) this.granularTransfers[dfuCode][sourceVariant] = {};
        if (!this.granularTransfers[dfuCode][sourceVariant][targetVariant]) this.granularTransfers[dfuCode][sourceVariant][targetVariant] = {};
        
        const weekData = this.granularTransfers[dfuCode][sourceVariant][targetVariant][weekKey];
        if (weekData) {
            weekData.selected = !weekData.selected;
        }
        
        // Update the Selected Demand box
        this.updateSelectedDemandDisplay(sourceVariant);
    }
    
    updateGranularQuantity(dfuCode, sourceVariant, targetVariant, weekKey, quantity) {
        if (!this.granularTransfers[dfuCode]?.[sourceVariant]?.[targetVariant]?.[weekKey]) return;
        
        const weekData = this.granularTransfers[dfuCode][sourceVariant][targetVariant][weekKey];
        weekData.customQuantity = quantity === '' ? null : parseFloat(quantity);
        
        // Update the Selected Demand box
        this.updateSelectedDemandDisplay(sourceVariant);
    }
    
    updateSelectedDemandDisplay(sourceVariant) {
        const selectedDemandEl = document.getElementById(`selected-demand-${sourceVariant}`);
        if (selectedDemandEl) {
            const selectedDemand = this.calculateSelectedDemand(sourceVariant);
            selectedDemandEl.textContent = this.formatNumber(selectedDemand);
        }
    }
    
    ensureGranularContainers() {
        if (!this.selectedDFU) return;
        
        const dfuData = this.multiVariantDFUs[this.selectedDFU];
        if (!dfuData) return;
        
        dfuData.variants.forEach(variant => {
            const container = document.getElementById(`granular-${variant}`);
            const targetVariant = this.transfers[this.selectedDFU]?.[variant];
            
            if (container) {
                // Only show granular section if there's a valid target selected
                if (targetVariant && targetVariant !== variant && targetVariant !== '') {
                    const granularHtml = this.renderGranularTransferSection(variant, targetVariant);
                    container.innerHTML = granularHtml;
                    this.attachGranularEventListeners();
                } else {
                    // Clear the container if no valid target
                    container.innerHTML = '';
                }
            }
        });
    }
    
    renderGranularTransferSection(sourceVariant, targetVariant) {
        if (!this.selectedDFU) return '';
        
        const demandColumn = 'weekly fcst';
        const weekNumberColumn = 'Week Number';
        const sourceLocationColumn = 'Source Location';
        const calendarWeekColumn = 'Calendar.week';
        const partNumberColumn = 'Product Number';
        
        const sourceRecords = this.rawData.filter(r => 
            this.toComparableString(r['DFU']) === this.selectedDFU &&
            this.toComparableString(r[partNumberColumn]) === sourceVariant
        );
        
        if (!this.granularTransfers[this.selectedDFU]) {
            this.granularTransfers[this.selectedDFU] = {};
        }
        if (!this.granularTransfers[this.selectedDFU][sourceVariant]) {
            this.granularTransfers[this.selectedDFU][sourceVariant] = {};
        }
        if (!this.granularTransfers[this.selectedDFU][sourceVariant][targetVariant]) {
            this.granularTransfers[this.selectedDFU][sourceVariant][targetVariant] = {};
        }
        
        const granularData = this.granularTransfers[this.selectedDFU][sourceVariant][targetVariant];
        
        sourceRecords.forEach(record => {
            const weekNumber = this.toComparableString(record[weekNumberColumn]);
            const sourceLocation = this.toComparableString(record[sourceLocationColumn]);
            const weekKey = `${weekNumber}-${sourceLocation}`;
            const demand = parseFloat(record[demandColumn]) || 0;
            
            if (!granularData[weekKey]) {
                granularData[weekKey] = {
                    selected: false,
                    originalQuantity: demand,
                    customQuantity: null,
                    calendarWeek: record[calendarWeekColumn],
                    weekNumber: weekNumber,
                    sourceLocation: sourceLocation
                };
            }
        });
        
        const sortedWeeks = Object.entries(granularData).sort((a, b) => {
            const weekA = parseInt(a[1].weekNumber);
            const weekB = parseInt(b[1].weekNumber);
            return weekA - weekB;
        });
        
        return `
            <div class="mt-4 p-3 bg-blue-50 rounded-lg border border-blue-200">
                <h5 class="font-semibold text-blue-800 mb-3">Granular Transfer (Select Specific Weeks)</h5>
                <div class="space-y-2 max-h-64 overflow-y-auto">
                    ${sortedWeeks.map(([weekKey, weekData]) => {
                        const isSelected = weekData.selected;
                        const customQty = weekData.customQuantity;
                        
                        return `
                            <div class="flex items-center gap-3 p-2 bg-white rounded border ${isSelected ? 'border-blue-500' : 'border-gray-200'}">
                                <input 
                                    type="checkbox" 
                                    ${isSelected ? 'checked' : ''}
                                    class="w-4 h-4"
                                    data-granular-toggle
                                    data-dfu="${this.selectedDFU}"
                                    data-source="${sourceVariant}"
                                    data-target="${targetVariant}"
                                    data-week="${weekKey}"
                                >
                                <div class="flex-1 grid grid-cols-3 gap-2 text-sm">
                                    <div>
                                        <span class="text-gray-600">Week:</span>
                                        <span class="font-medium">${weekData.weekNumber}</span>
                                    </div>
                                    <div>
                                        <span class="text-gray-600">Location:</span>
                                        <span class="font-medium">${weekData.sourceLocation}</span>
                                    </div>
                                    <div>
                                        <span class="text-gray-600">Date:</span>
                                        <span class="font-medium">${weekData.calendarWeek || 'N/A'}</span>
                                    </div>
                                </div>
                                <div class="flex items-center gap-2">
                                    <span class="text-sm text-gray-600">Demand:</span>
                                    <input 
                                        type="number" 
                                        class="w-24 px-2 py-1 border rounded text-sm"
                                        placeholder="${this.formatNumber(weekData.originalQuantity)}"
                                        value="${customQty !== null && customQty !== '' ? customQty : ''}"
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
            </div>
        `;
    }
    
    async executeTransfer() {
        if (!this.selectedDFU) {
            this.showNotification('Please select a DFU first', 'error');
            return;
        }
        
        const timestamp = new Date().toISOString().split('T')[0] + ' ' + new Date().toTimeString().split(' ')[0];
        const demandColumn = 'weekly fcst';
        const partNumberColumn = 'Product Number';
        const calendarWeekColumn = 'Calendar.week';
        const sourceLocationColumn = 'Source Location';
        const weekNumberColumn = 'Week Number';
        
        const dfuStr = this.selectedDFU;
        const dfuRecords = this.rawData.filter(r => this.toComparableString(r['DFU']) === dfuStr);
        
        let executionType = '';
        let executionMessage = '';
        let executionDetails = '';
        
        if (this.bulkTransfers[dfuStr]) {
            const targetVariant = this.bulkTransfers[dfuStr];
            const originalVariants = [...new Set(dfuRecords.map(r => this.toComparableString(r[partNumberColumn])))];
            const variantsToTransfer = originalVariants.filter(v => v !== targetVariant);
            
            variantsToTransfer.forEach(sourceVariant => {
                const sourceRecords = dfuRecords.filter(r => 
                    this.toComparableString(r[partNumberColumn]) === sourceVariant
                );
                
                sourceRecords.forEach(record => {
                    const transferDemand = parseFloat(record[demandColumn]) || 0;
                    const targetRecord = dfuRecords.find(r => 
                        this.toComparableString(r[partNumberColumn]) === targetVariant && 
                        this.toComparableString(r[calendarWeekColumn]) === this.toComparableString(record[calendarWeekColumn]) &&
                        this.toComparableString(r[sourceLocationColumn]) === this.toComparableString(record[sourceLocationColumn])
                    );
                    
                    if (targetRecord) {
                        const oldDemand = parseFloat(targetRecord[demandColumn]) || 0;
                        targetRecord[demandColumn] = oldDemand + transferDemand;
                        const existingHistory = targetRecord['Transfer History'] || '';
                        const newHistoryEntry = `[${sourceVariant} → ${transferDemand} @ ${timestamp}]`;
                        const pipoPrefix = existingHistory.startsWith('PIPO') ? '' : 'PIPO ';
                        targetRecord['Transfer History'] = existingHistory ? 
                            `${existingHistory} ${newHistoryEntry}` : `${pipoPrefix}${newHistoryEntry}`;
                        record[demandColumn] = 0;
                    } else {
                        const originalVariant = this.toComparableString(record[partNumberColumn]);
                        record[partNumberColumn] = isNaN(targetVariant) ? targetVariant : Number(targetVariant);
                        record['Transfer History'] = `PIPO [${originalVariant} → ${transferDemand} @ ${timestamp}]`;
                    }
                });
            });
            
            this.completedTransfers[dfuStr] = {
                type: 'bulk',
                targetVariant: targetVariant,
                timestamp: timestamp,
                originalVariantCount: originalVariants.length
            };
            
            delete this.bulkTransfers[dfuStr];
            executionType = 'Bulk Transfer';
            executionMessage = `All variants consolidated to ${targetVariant}`;
            this.showNotification(executionMessage, 'success');
            
        } else if (this.transfers[dfuStr] && Object.keys(this.transfers[dfuStr]).length > 0) {
            const individualTransfers = this.transfers[dfuStr];
            let transferCount = 0;
            const transferHistory = [];
            
            Object.keys(individualTransfers).forEach(sourceVariant => {
                const targetVariant = individualTransfers[sourceVariant];
                
                if (sourceVariant !== targetVariant) {
                    const sourceRecords = dfuRecords.filter(r => 
                        this.toComparableString(r[partNumberColumn]) === sourceVariant
                    );
                    
                    sourceRecords.forEach(record => {
                        const transferDemand = parseFloat(record[demandColumn]) || 0;
                        const targetRecord = dfuRecords.find(r => 
                            this.toComparableString(r[partNumberColumn]) === targetVariant && 
                            this.toComparableString(r[calendarWeekColumn]) === this.toComparableString(record[calendarWeekColumn]) &&
                            this.toComparableString(r[sourceLocationColumn]) === this.toComparableString(record[sourceLocationColumn])
                        );
                        
                        if (targetRecord) {
                            const oldDemand = parseFloat(targetRecord[demandColumn]) || 0;
                            targetRecord[demandColumn] = oldDemand + transferDemand;
                            const existingHistory = targetRecord['Transfer History'] || '';
                            const newHistoryEntry = `[${sourceVariant} → ${transferDemand} @ ${timestamp}]`;
                            const pipoPrefix = existingHistory.startsWith('PIPO') ? '' : 'PIPO ';
                            targetRecord['Transfer History'] = existingHistory ? 
                                `${existingHistory} ${newHistoryEntry}` : `${pipoPrefix}${newHistoryEntry}`;
                            record[demandColumn] = 0;
                        } else {
                            const originalVariant = this.toComparableString(record[partNumberColumn]);
                            record[partNumberColumn] = isNaN(targetVariant) ? targetVariant : Number(targetVariant);
                            record['Transfer History'] = `PIPO [${originalVariant} → ${transferDemand} @ ${timestamp}]`;
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
            this.completedTransfers[dfuStr] = {
                type: 'individual',
                transfers: individualTransfers,
                timestamp: timestamp,
                transferCount: transferCount,
                transferHistory
            };
            
            executionType = 'Individual Transfers';
            executionMessage = `${transferCount} variant transfers executed`;
            this.showNotification(`Individual transfers completed for DFU ${dfuStr}: ${executionMessage}`);
            
        } else if (this.granularTransfers[dfuStr] && Object.keys(this.granularTransfers[dfuStr]).length > 0) {
            const granularTransfers = this.granularTransfers[dfuStr];
            let granularTransferCount = 0;
            
            Object.keys(granularTransfers).forEach(sourceVariant => {
                const sourceTargets = granularTransfers[sourceVariant];
                
                Object.keys(sourceTargets).forEach(targetVariant => {
                    const weekTransfers = sourceTargets[targetVariant];
                    
                    Object.keys(weekTransfers).forEach(weekKey => {
                        const weekTransfer = weekTransfers[weekKey];
                        if (!weekTransfer.selected) return;
                        
                        const [weekNumber, sourceLocation] = weekKey.split('-');
                        const sourceRecord = dfuRecords.find(r => 
                            this.toComparableString(r[partNumberColumn]) === sourceVariant &&
                            this.toComparableString(r[weekNumberColumn]) === weekNumber &&
                            this.toComparableString(r[sourceLocationColumn]) === sourceLocation
                        );
                        
                        if (sourceRecord) {
                            const originalDemand = parseFloat(sourceRecord[demandColumn]) || 0;
                            const transferAmount = weekTransfer.customQuantity !== null ?
                                parseFloat(weekTransfer.customQuantity) : originalDemand;
                            
                            const targetRecord = dfuRecords.find(r => 
                                this.toComparableString(r[partNumberColumn]) === targetVariant &&
                                this.toComparableString(r[weekNumberColumn]) === weekNumber &&
                                this.toComparableString(r[sourceLocationColumn]) === sourceLocation
                            );
                            
                            if (targetRecord) {
                                const oldDemand = parseFloat(targetRecord[demandColumn]) || 0;
                                targetRecord[demandColumn] = oldDemand + transferAmount;
                                const existingHistory = targetRecord['Transfer History'] || '';
                                const newHistoryEntry = `[${sourceVariant} → ${transferAmount} @ ${timestamp}]`;
                                const pipoPrefix = existingHistory.startsWith('PIPO') ? '' : 'PIPO ';
                                targetRecord['Transfer History'] = existingHistory ? 
                                    `${existingHistory} ${newHistoryEntry}` : `${pipoPrefix}${newHistoryEntry}`;
                            }
                            
                            sourceRecord[demandColumn] = originalDemand - transferAmount;
                            if (!sourceRecord['Transfer History']) {
                                sourceRecord['Transfer History'] = '';
                            }
                            sourceRecord['Transfer History'] += ` PIPO [→ ${targetVariant}: -${transferAmount} @ ${timestamp}]`;
                            
                            granularTransferCount++;
                        }
                    });
                });
            });
            
            this.completedTransfers[dfuStr] = {
                type: 'granular',
                timestamp: timestamp,
                transferCount: granularTransferCount
            };
            
            delete this.granularTransfers[dfuStr];
            executionType = 'Granular Transfer';
            executionMessage = `${granularTransferCount} week-specific transfers executed`;
            this.showNotification(executionMessage, 'success');
        }
        
        this.lastExecutionSummary[dfuStr] = {
            type: executionType,
            message: executionMessage,
            details: executionDetails,
            timestamp: timestamp
        };
        
        this.processMultiVariantDFUs(this.rawData);
        this.render();
    }
    
    async undoTransfer(dfuCode) {
        const dfuStr = this.toComparableString(dfuCode);
        
        if (!this.originalRawData || this.originalRawData.length === 0) {
            this.showNotification('No original data available to restore', 'error');
            return;
        }
        
        delete this.completedTransfers[dfuStr];
        delete this.transfers[dfuStr];
        delete this.bulkTransfers[dfuStr];
        delete this.granularTransfers[dfuStr];
        delete this.lastExecutionSummary[dfuStr];
        
        const originalDfuRecords = this.originalRawData.filter(record => 
            this.toComparableString(record['DFU']) === dfuStr
        );
        
        this.rawData = this.rawData.filter(record => 
            this.toComparableString(record['DFU']) !== dfuStr
        );
        
        const restoredRecords = originalDfuRecords.map(record => ({...record}));
        this.rawData.push(...restoredRecords);
        
        this.multiVariantDFUs = {};
        this.filteredDFUs = {};
        this.processMultiVariantDFUs(this.rawData);
        
        this.showNotification(`Transfer undone for DFU ${dfuStr}. All original variants and quantities restored.`, 'success');
        
        const currentSelection = this.selectedDFU;
        this.selectedDFU = null;
        setTimeout(() => {
            this.selectedDFU = currentSelection;
            this.render();
        }, 50);
    }
    
    async addManualVariant(dfuCode) {
        const variantCode = prompt('Enter the new variant/part code:');
        if (!variantCode) return;
        
        const dfuRecords = this.rawData.filter(r => 
            this.toComparableString(r['DFU']) === dfuCode
        );
        
        if (dfuRecords.length === 0) {
            this.showNotification('No records found for this DFU', 'error');
            return;
        }
        
        const templateRecord = dfuRecords[0];
        const uniqueWeekLocations = new Set();
        dfuRecords.forEach(r => {
            const key = `${r['Week Number']}-${r['Source Location']}-${r['Calendar.week']}`;
            uniqueWeekLocations.add(key);
        });
        
        const newRecords = [];
        uniqueWeekLocations.forEach(key => {
            const [weekNum, sourceLoc, calWeek] = key.split('-');
            const newRecord = { ...templateRecord };
            newRecord['Product Number'] = variantCode;
            newRecord['weekly fcst'] = 0;
            newRecord['Week Number'] = weekNum;
            newRecord['Source Location'] = sourceLoc;
            newRecord['Calendar.week'] = calWeek;
            newRecord['Transfer History'] = 'PIPO [Manually added variant]';
            newRecords.push(newRecord);
        });
        
        this.rawData.push(...newRecords);
        this.processMultiVariantDFUs(this.rawData);
        this.showNotification(`Added variant ${variantCode} to DFU ${dfuCode}`, 'success');
        this.render();
    }
    
    async exportData() {
        try {
            const wb = XLSX.utils.book_new();
            
            const formattedData = this.rawData.map(record => {
                const formattedRecord = { ...record };
                
                if (formattedRecord['Week Number']) {
                    const weekNumStr = String(formattedRecord['Week Number']).trim();
                    const parts = weekNumStr.split('_');
                    
                    if (parts.length === 2) {
                        const weekNum = parseInt(parts[0]);
                        const year = parseInt(parts[1]);
                        
                        if (!isNaN(weekNum) && !isNaN(year) && weekNum >= 1 && weekNum <= 53) {
                            const targetDate = this.getDateFromWeekNumber(year, weekNum);
                            formattedRecord['Calendar.week'] = targetDate.toISOString().split('T')[0];
                        }
                    }
                }
                
                return formattedRecord;
            });
            
            const ws = XLSX.utils.json_to_sheet(formattedData);
            XLSX.utils.book_append_sheet(wb, ws, 'Updated Demand');
            XLSX.writeFile(wb, 'DFU_Transfer_Updated.xlsx');
            
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
                            <p class="text-gray-600 mb-4">Upload your Excel file with the "Total Demand" format</p>
                            
                            ${this.isLoading ? `
                                <div class="text-blue-600">
                                    <div class="loading-spinner mb-2"></div>
                                    <p>Processing file...</p>
                                </div>
                            ` : `
                                <div class="space-y-4">
                                    <div>
                                        <input type="file" accept=".xlsx,.xls" class="file-input" id="fileInput">
                                        <p class="text-sm text-gray-500 mt-2">Supported formats: .xlsx, .xls</p>
                                    </div>
                                    
                                    <!-- Supply Chain File Uploads -->
                                    <div class="border-t pt-4 mt-4 space-y-3">
                                        <h3 class="text-sm font-medium text-gray-700 mb-2">Optional: Upload Supply Chain Data</h3>
                                        
                                        <div class="space-y-2">
                                            <label class="block">
                                                <span class="text-sm text-gray-600">Stock RRP4 (SOH):</span>
                                                <input type="file" accept=".xlsx,.xls" class="file-input text-sm" id="stockFileInput">
                                                ${this.hasStockData ? '<span class="text-xs text-green-600">✓ Loaded</span>' : ''}
                                            </label>
                                            
                                            <label class="block">
                                                <span class="text-sm text-gray-600">Production RRP4 (Open Supply):</span>
                                                <input type="file" accept=".xlsx,.xls" class="file-input text-sm" id="openSupplyFileInput">
                                                ${this.hasOpenSupplyData ? '<span class="text-xs text-green-600">✓ Loaded</span>' : ''}
                                            </label>
                                            
                                            <label class="block">
                                                <span class="text-sm text-gray-600">Transport Receipts (In Transit):</span>
                                                <input type="file" accept=".xlsx,.xls" class="file-input text-sm" id="transitFileInput">
                                                ${this.hasTransitData ? '<span class="text-xs text-green-600">✓ Loaded</span>' : ''}
                                            </label>
                                        </div>
                                    </div>
                                    
                                    <div class="border-t pt-4 mt-4">
                                        <h3 class="text-sm font-medium text-gray-700 mb-2">Optional: Upload Variant Cycle Dates</h3>
                                        <input type="file" accept=".xlsx,.xls" class="file-input" id="cycleFileInput">
                                        <p class="text-xs text-gray-500 mt-1">Upload file with DFU, Part Code, SOS, and EOS columns</p>
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
                
                const stockFileInput = document.getElementById('stockFileInput');
                if (stockFileInput) {
                    stockFileInput.addEventListener('change', (e) => {
                        const file = e.target.files[0];
                        if (file) this.loadStockFile(file);
                    });
                }
                
                const openSupplyFileInput = document.getElementById('openSupplyFileInput');
                if (openSupplyFileInput) {
                    openSupplyFileInput.addEventListener('change', (e) => {
                        const file = e.target.files[0];
                        if (file) this.loadOpenSupplyFile(file);
                    });
                }
                
                const transitFileInput = document.getElementById('transitFileInput');
                if (transitFileInput) {
                    transitFileInput.addEventListener('change', (e) => {
                        const file = e.target.files[0];
                        if (file) this.loadTransitFile(file);
                    });
                }
                
                const cycleFileInput = document.getElementById('cycleFileInput');
                if (cycleFileInput) {
                    cycleFileInput.addEventListener('change', (e) => {
                        const file = e.target.files[0];
                        if (file) this.loadVariantCycleData(file);
                    });
                }
            }
            
            return;
        }
        
        const totalDFUs = Object.keys(this.filteredDFUs).length;
        const multiVariantCount = Object.keys(this.filteredDFUs).filter(dfu => !this.filteredDFUs[dfu].isSingleVariant).length;
        
        app.innerHTML = `
            <div>
                <h1 class="text-3xl font-bold text-gray-800 mb-2">DFU Demand Transfer Management</h1>
                <p class="text-gray-600">Managing ${totalDFUs} DFUs (${multiVariantCount} with multiple variants, ${totalDFUs - multiVariantCount} single variant)</p>
            </div>

            <div class="flex gap-4 mb-6 flex-responsive">
                <div class="relative flex-1">
                    <svg class="absolute left-3 top-3 h-4 w-4 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
                    </svg>
                    <input type="text" placeholder="Search DFU codes or part codes..." value="${this.searchTerm}" class="search-input" id="searchInput">
                </div>
                <div class="relative">
                    <select class="px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent" id="plantLocationFilter">
                        <option value="">All Plant Locations</option>
                        ${this.availablePlantLocations.map(location => `
                            <option value="${location}" ${this.selectedPlantLocation === location ? 'selected' : ''}>Plant ${location}</option>
                        `).join('')}
                    </select>
                </div>
                <div class="relative">
                    <select class="px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent" id="productionLineFilter">
                        <option value="">All Production Lines</option>
                        ${this.availableProductionLines.map(line => `
                            <option value="${line}" ${this.selectedProductionLine === line ? 'selected' : ''}>Line ${line}</option>
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
                        <svg class="w-5 h-5 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2" />
                        </svg>
                        All DFUs (${totalDFUs})
                    </h3>
                    <div class="relative" style="height: 580px;">
                        <div class="absolute inset-x-0 top-0 h-4 bg-gradient-to-b from-gray-50 to-transparent pointer-events-none z-10"></div>
                        <div class="absolute inset-x-0 bottom-0 h-4 bg-gradient-to-t from-gray-50 to-transparent pointer-events-none z-10"></div>
                        <div class="space-y-2 h-full overflow-y-auto pr-1 scrollbar-custom" style="padding-top: 16px; padding-bottom: 16px;">
                            ${Object.keys(this.filteredDFUs).map(dfuCode => {
                            const dfuData = this.filteredDFUs[dfuCode];
                            if (!dfuData || !dfuData.variants) return '';
                            
                            return `
                                <div class="dfu-card ${this.selectedDFU === dfuCode ? 'selected' : ''}" data-dfu="${dfuCode}">
                                    <div class="flex justify-between items-start">
                                        <div>
                                            <h4 class="font-medium text-gray-800">DFU: ${dfuCode}</h4>
                                            <p class="text-sm text-gray-600">
                                                ${dfuData.plantLocations && dfuData.plantLocations.length > 0 ? `Plant${dfuData.plantLocations.length > 1 ? 's' : ''}: ${dfuData.plantLocations.join(', ')} • ` : ''}
                                                ${dfuData.productionLines && dfuData.productionLines.length > 0 ? `Line${dfuData.productionLines.length > 1 ? 's' : ''}: ${dfuData.productionLines.join(', ')} • ` : ''}
                                                ${dfuData.variants.length} variant${dfuData.variants.length > 1 ? 's' : ''}
                                                ${dfuData.isSingleVariant ? ' (single)' : ''}
                                                ${dfuData.isCompleted ? ' (transfer completed)' : ''}
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
                                            ` : (this.transfers[dfuCode] && Object.keys(this.transfers[dfuCode]).length > 0 && Object.values(this.transfers[dfuCode]).some(t => t && t !== '')) || this.bulkTransfers[dfuCode] || (this.granularTransfers[dfuCode] && Object.keys(this.granularTransfers[dfuCode]).length > 0) ? `
                                                <span class="inline-flex items-center gap-1 text-green-600 text-sm">
                                                    <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                                                    </svg>
                                                    Ready
                                                </span>
                                            ` : dfuData.isSingleVariant ? `
                                                <span class="text-blue-600 text-sm">Single</span>
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
                    ${this.selectedDFU && this.multiVariantDFUs[this.selectedDFU] ? this.renderDetailsSection() : `
                        <div class="text-center py-12 text-gray-400">
                            <svg class="w-16 h-16 mx-auto mb-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-6 9l2 2 4-4" />
                            </svg>
                            <p class="text-lg">Select a DFU to view details</p>
                        </div>
                    `}
                </div>
            </div>
        `;
        
        this.attachEventListeners();
        this.ensureGranularContainers();
    }
    
    renderDetailsSection() {
        if (!this.selectedDFU || !this.multiVariantDFUs[this.selectedDFU]) return '';
        
        return `
            <div>
                <div class="flex justify-between items-center mb-4">
                    <h3 class="font-semibold text-gray-800">
                        DFU: ${this.selectedDFU}
                        ${this.multiVariantDFUs[this.selectedDFU].plantLocations && this.multiVariantDFUs[this.selectedDFU].plantLocations.length > 0 ? 
                            ` (Plant${this.multiVariantDFUs[this.selectedDFU].plantLocations.length > 1 ? 's' : ''}: ${this.multiVariantDFUs[this.selectedDFU].plantLocations.join(', ')})` : ''}
                        ${this.multiVariantDFUs[this.selectedDFU].productionLines && this.multiVariantDFUs[this.selectedDFU].productionLines.length > 0 ? 
                            ` (Line${this.multiVariantDFUs[this.selectedDFU].productionLines.length > 1 ? 's' : ''}: ${this.multiVariantDFUs[this.selectedDFU].productionLines.join(', ')})` : ''}
                        - Variant Details
                        ${this.multiVariantDFUs[this.selectedDFU].isCompleted ? `
                            <span class="ml-2 px-2 py-1 text-xs bg-green-100 text-green-800 rounded-full">✓ Transfer Complete</span>
                        ` : ''}
                    </h3>
                    ${!this.multiVariantDFUs[this.selectedDFU].isCompleted ? `
                        <button class="btn btn-primary text-sm" id="addVariantBtn">
                            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 4v16m8-8H4" />
                            </svg>
                            Add Variant
                        </button>
                    ` : ''}
                </div>
                
                ${this.multiVariantDFUs[this.selectedDFU].isCompleted ? `
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
                                        <p><strong>Week-specific Transfers:</strong> ${this.multiVariantDFUs[this.selectedDFU].completionInfo.transferCount}</p>
                                    ` : `
                                        <p><strong>Variant Transfers:</strong> ${this.multiVariantDFUs[this.selectedDFU].completionInfo.transferCount}</p>
                                    `}
                                </div>
                            </div>
                            <button class="btn btn-secondary text-sm" onclick="appInstance.undoTransfer('${this.selectedDFU}')">
                                <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 10h10a8 8 0 018 8v2M3 10l6 6m-6-6l6-6" />
                                </svg>
                                Undo
                            </button>
                        </div>
                    </div>
                    
                    <div class="mb-6">
                        <h4 class="font-semibold text-gray-800 mb-3">Current Variant Status</h4>
                        <div class="space-y-3">
                            ${this.multiVariantDFUs[this.selectedDFU].variants.map(variant => {
                            const demandData = this.multiVariantDFUs[this.selectedDFU].variantDemand[variant];
                            const supplyChain = this.getSupplyChainData(variant);
                            
                            return `
                                <div class="border rounded-lg p-3 bg-white">
                                    <div class="flex justify-between items-center">
                                        <div class="flex-1">
                                            <h5 class="font-medium text-gray-800">Part: ${variant}</h5>
                                            <p class="text-xs text-gray-500 mb-1 max-w-md break-words">${demandData?.partDescription || 'Description not available'}</p>
                                            <p class="text-sm text-gray-600">${demandData?.recordCount || 0} records</p>
                                            ${supplyChain.hasData ? `
                                                <div class="mt-2 text-xs space-y-1">
                                                    <p class="text-blue-600"><strong>SOH:</strong> ${this.formatNumber(supplyChain.soh)}</p>
                                                    <p class="text-green-600"><strong>Open Supply:</strong> ${this.formatNumber(supplyChain.openSupply)}</p>
                                                    <p class="text-purple-600"><strong>In Transit:</strong> ${this.formatNumber(supplyChain.transit)}</p>
                                                    <p class="text-gray-800 font-semibold"><strong>Total Network Stock:</strong> ${this.formatNumber(supplyChain.total)}</p>
                                                </div>
                                            ` : ''}
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
                            <div class="mt-3 p-2 bg-white rounded border border-purple-200">
                                <p class="text-sm text-purple-800">
                                    <strong>Selected:</strong> All variants will be transferred to <strong>${this.bulkTransfers[this.selectedDFU]}</strong>
                                </p>
                            </div>
                        ` : ''}
                    </div>
                    
                    ${this.renderIndividualTransferSection()}
                    ${this.renderActionButtons()}
                `}
                
                <div class="text-center text-sm text-gray-500 mt-6">
                    <p><strong>How to Use:</strong></p>
                    <ul class="text-left list-disc list-inside mt-2 space-y-1">
                        <li><strong>Bulk:</strong> Click a variant button above to transfer all to that variant</li>
                        <li><strong>Individual:</strong> Use dropdowns below to set specific transfers per variant</li>
                        <li><strong>Granular:</strong> After setting individual transfer, select specific weeks to transfer partial demand</li>
                        <li><strong>Execute:</strong> Click "Execute Transfer" to apply your chosen transfers</li>
                        <li><strong>Export:</strong> Export the updated data when you're done with all transfers</li>
                    </ul>
                </div>
            </div>
        `;
    }
    
    renderIndividualTransferSection() {
        if (!this.selectedDFU || !this.multiVariantDFUs[this.selectedDFU]) return '';
        
        return `
            <div class="mb-6">
                <h4 class="font-semibold text-gray-800 mb-3">Individual Transfers (Variant → Specific Target)</h4>
                <div class="space-y-4">
                    ${this.multiVariantDFUs[this.selectedDFU].variants.map(variant => {
                    const demandData = this.multiVariantDFUs[this.selectedDFU].variantDemand[variant];
                    const currentTransfer = this.transfers[this.selectedDFU]?.[variant];
                    const supplyChain = this.getSupplyChainData(variant);
                    const networkCalc = this.calculateNetworkStockPlusDemand(variant);
                    const selectedDemand = this.calculateSelectedDemand(variant);
                    
                    return `
                        <div class="border rounded-lg p-4 bg-gray-50">
                            <div class="flex justify-between items-center mb-3">
                                <div class="flex-1">
                                    <h5 class="font-medium text-gray-800">Part: ${variant}</h5>
                                    <p class="text-xs text-gray-500 mb-1 max-w-md break-words">${demandData?.partDescription || 'Description not available'}</p>
                                    <p class="text-sm text-gray-600">${demandData?.recordCount || 0} records • ${this.formatNumber(demandData?.totalDemand || 0)} total demand</p>
                                    ${supplyChain.hasData ? `
                                        <div class="mt-2 text-xs space-y-1">
                                            <p class="text-blue-600"><strong>SOH:</strong> ${this.formatNumber(supplyChain.soh)}</p>
                                            <p class="text-green-600"><strong>Open Supply:</strong> ${this.formatNumber(supplyChain.openSupply)}</p>
                                            <p class="text-purple-600"><strong>In Transit:</strong> ${this.formatNumber(supplyChain.transit)}</p>
                                            <p class="text-gray-800 font-semibold"><strong>Total Network Stock:</strong> ${this.formatNumber(supplyChain.total)}</p>
                                        </div>
                                    ` : ''}
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
                            
                            ${currentTransfer && currentTransfer !== variant && currentTransfer !== '' ? `
                                <!-- Network Stock + Total Demand Calculation Box -->
                                <div class="grid grid-cols-2 gap-3 mb-3">
                                    <div class="p-3 bg-orange-50 border border-orange-300 rounded-lg">
                                        <p class="text-xs text-orange-700 mb-1 font-semibold">Network Stock + Total Demand</p>
                                        <p class="text-2xl font-bold ${networkCalc && networkCalc.hasSupplyData ? (networkCalc.result < 0 ? 'text-red-600' : 'text-green-600') : 'text-orange-900'}" id="network-calc-${variant}">
                                            ${networkCalc && networkCalc.hasSupplyData ? this.formatNumber(networkCalc.result) : '-'}
                                        </p>
                                        ${networkCalc && networkCalc.hasSupplyData ? `
                                            <p class="text-xs text-orange-600 mt-1">
                                                ${this.formatNumber(networkCalc.networkStock)} + (${this.formatNumber(networkCalc.totalDemand)})
                                            </p>
                                        ` : '<p class="text-xs text-orange-600 mt-1">No supply data</p>'}
                                    </div>
                                    
                                    <div class="p-3 bg-teal-50 border border-teal-300 rounded-lg">
                                        <p class="text-xs text-teal-700 mb-1 font-semibold">Selected Demand</p>
                                        <p class="text-2xl font-bold text-teal-900" id="selected-demand-${variant}">
                                            ${this.formatNumber(selectedDemand)}
                                        </p>
                                        <p class="text-xs text-teal-600 mt-1">Updates as you select weeks</p>
                                    </div>
                                </div>
                            ` : ''}
                            
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
                                ${currentTransfer && currentTransfer !== variant && currentTransfer !== '' ? `
                                    <span class="text-green-600 text-sm">→ ${currentTransfer}</span>
                                ` : ''}
                            </div>
                            
                            <div id="granular-${variant}"></div>
                        </div>
                    `;
                    }).join('')}
                </div>
            </div>
        `;
    }
    
    renderActionButtons() {
        if (!this.selectedDFU) return '';
        
        const hasValidTransfers = this.transfers[this.selectedDFU] && 
            Object.values(this.transfers[this.selectedDFU]).some(t => t && t !== '');
        
        const hasTransfers = (hasValidTransfers || 
                             this.bulkTransfers[this.selectedDFU] || 
                             (this.granularTransfers[this.selectedDFU] && Object.keys(this.granularTransfers[this.selectedDFU]).length > 0));
        
        if (hasTransfers) {
            return `
                <div class="p-3 bg-blue-50 rounded-lg">
                    <div class="text-sm text-blue-800 mb-3">
                        ${this.bulkTransfers[this.selectedDFU] ? 
                            `<p>✓ Bulk transfer to <strong>${this.bulkTransfers[this.selectedDFU]}</strong> is ready</p>` : 
                            this.granularTransfers[this.selectedDFU] && Object.keys(this.granularTransfers[this.selectedDFU]).length > 0 ?
                            `<p>✓ Granular transfers configured</p>` :
                            `<p>✓ Individual transfers configured</p>`
                        }
                    </div>
                    <button class="btn btn-primary w-full" id="executeBtn">
                        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                        </svg>
                        Execute Transfer
                    </button>
                </div>
            `;
        }
        
        return '';
    }
    
    attachEventListeners() {
        const searchInput = document.getElementById('searchInput');
        if (searchInput) {
            searchInput.addEventListener('input', (e) => {
                clearTimeout(this.searchDebounceTimer);
                this.searchDebounceTimer = setTimeout(() => {
                    this.searchTerm = e.target.value;
                    this.applySearchFilter();
                    this.render();
                }, 300);
            });
        }
        
        const plantLocationFilter = document.getElementById('plantLocationFilter');
        if (plantLocationFilter) {
            plantLocationFilter.addEventListener('change', (e) => {
                this.selectedPlantLocation = e.target.value;
                this.processMultiVariantDFUs(this.rawData);
                this.render();
            });
        }
        
        const productionLineFilter = document.getElementById('productionLineFilter');
        if (productionLineFilter) {
            productionLineFilter.addEventListener('change', (e) => {
                this.selectedProductionLine = e.target.value;
                this.processMultiVariantDFUs(this.rawData);
                this.render();
            });
        }
        
        const exportBtn = document.getElementById('exportBtn');
        if (exportBtn) {
            exportBtn.addEventListener('click', () => this.exportData());
        }
        
        const executeBtn = document.getElementById('executeBtn');
        if (executeBtn) {
            executeBtn.addEventListener('click', () => this.executeTransfer());
        }
        
        const addVariantBtn = document.getElementById('addVariantBtn');
        if (addVariantBtn) {
            addVariantBtn.addEventListener('click', () => this.addManualVariant(this.selectedDFU));
        }
        
        const cycleFileInput = document.getElementById('cycleFileInput');
        if (cycleFileInput) {
            cycleFileInput.addEventListener('change', (e) => {
                const file = e.target.files[0];
                if (file) this.loadVariantCycleData(file);
            });
        }
        
        document.querySelectorAll('.dfu-card').forEach(card => {
            card.addEventListener('click', (e) => {
                const dfuCode = e.currentTarget.dataset.dfu;
                this.selectDFU(dfuCode);
            });
        });
        
        document.querySelectorAll('[data-bulk-target]').forEach(button => {
            button.addEventListener('click', (e) => {
                const targetVariant = e.target.dataset.bulkTarget;
                this.selectBulkTarget(this.selectedDFU, targetVariant);
            });
        });
        
        document.querySelectorAll('[data-source-variant]').forEach(select => {
            select.addEventListener('change', (e) => {
                const sourceVariant = e.target.dataset.sourceVariant;
                const targetVariant = e.target.value;
                
                // Always call setIndividualTransfer, even if empty (to clear)
                this.setIndividualTransfer(this.selectedDFU, sourceVariant, targetVariant);
            });
        });
    }
    
    attachGranularEventListeners() {
        document.querySelectorAll('[data-granular-toggle]').forEach(checkbox => {
            const newCheckbox = checkbox.cloneNode(true);
            checkbox.parentNode.replaceChild(newCheckbox, checkbox);
            
            newCheckbox.addEventListener('change', (e) => {
                const dfuCode = e.target.dataset.dfu;
                const sourceVariant = e.target.dataset.source;
                const targetVariant = e.target.dataset.target;
                const weekKey = e.target.dataset.week;
                
                this.toggleGranularWeek(dfuCode, sourceVariant, targetVariant, weekKey);
                
                const qtyInput = document.querySelector(`[data-granular-qty][data-week="${weekKey}"][data-source="${sourceVariant}"][data-target="${targetVariant}"]`);
                if (qtyInput) {
                    qtyInput.disabled = !e.target.checked;
                }
            });
        });
        
        document.querySelectorAll('[data-granular-qty]').forEach(input => {
            const newInput = input.cloneNode(true);
            input.parentNode.replaceChild(newInput, input);
            
            newInput.addEventListener('input', (e) => {
                const dfuCode = e.target.dataset.dfu;
                const sourceVariant = e.target.dataset.source;
                const targetVariant = e.target.dataset.target;
                const weekKey = e.target.dataset.week;
                const quantity = e.target.value;
                
                this.updateGranularQuantity(dfuCode, sourceVariant, targetVariant, weekKey, quantity);
            });
        });
    }
}

// Initialize the application when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    if (!window.preventAutoInit) {
        window.appInstance = new DemandTransferApp();
    }
});