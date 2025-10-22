// DFU Demand Transfer Management Application
// Version: 2.22.0 - Build: 2025-10-22-supply-chain-metrics
// Added SOH, Open Supply, and Stock in Transit display with totals

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
        
        // Supply chain data
        this.stockData = {}; // SOH: { productNumber: totalStock }
        this.hasStockData = false;
        this.openSupplyData = {}; // Open Supply: { productNumber: totalQuantity }
        this.hasOpenSupplyData = false;
        this.stockInTransitData = {}; // Stock in Transit: { productNumber: totalQuantity }
        this.hasStockInTransitData = false;
        
        this.init();
    }
    
    init() {
        console.log('🚀 DFU Demand Transfer App v2.22.0 - Build: 2025-10-22-supply-chain-metrics');
        console.log('📋 Added SOH, Open Supply, and Stock in Transit with totals');
        this.render();
        this.attachEventListeners();
    }
    
    toComparableString(value) {
        if (value === null || value === undefined) return '';
        return String(value).trim();
    }
    
    getDateFromWeekNumber(year, weekNumber) {
        const jan1 = new Date(year, 0, 1);
        const jan1DayOfWeek = jan1.getDay();
        const daysToFirstMonday = jan1DayOfWeek === 0 ? 1 : (8 - jan1DayOfWeek);
        const firstMonday = new Date(year, 0, jan1.getDate() + daysToFirstMonday);
        const targetDate = new Date(firstMonday);
        targetDate.setDate(firstMonday.getDate() + (weekNumber - 1) * 7);
        return targetDate;
    }
    
    formatNumber(num) {
        if (num === null || num === undefined) return '0';
        return Math.round(num).toLocaleString();
    }
    
    showNotification(message, type = 'success') {
        console.log(`[${type.toUpperCase()}] ${message}`);
        
        const container = document.getElementById('notifications') || this.createNotificationContainer();
        const notification = document.createElement('div');
        
        const colors = {
            success: 'bg-green-500',
            error: 'bg-red-500',
            info: 'bg-blue-500',
            warning: 'bg-yellow-500'
        };
        
        notification.className = `${colors[type] || colors.info} text-white px-6 py-3 rounded-lg shadow-lg mb-2`;
        notification.textContent = message;
        
        container.appendChild(notification);
        
        setTimeout(() => {
            notification.style.opacity = '0';
            setTimeout(() => notification.remove(), 300);
        }, 5000);
    }
    
    createNotificationContainer() {
        const container = document.createElement('div');
        container.id = 'notifications';
        container.style.cssText = 'position: fixed; top: 20px; right: 20px; z-index: 9999;';
        document.body.appendChild(container);
        return container;
    }

    async handleStockFile(file) {
        console.log('Processing Stock RRP4 file...');
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
            
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            
            console.log(`Loaded ${jsonData.length} stock records`);
            
            this.stockData = {};
            
            jsonData.forEach(record => {
                const productNumber = this.toComparableString(record['Product Number']);
                const stock = parseFloat(record['Stock']) || 0;
                
                if (productNumber) {
                    if (!this.stockData[productNumber]) {
                        this.stockData[productNumber] = 0;
                    }
                    this.stockData[productNumber] += stock;
                }
            });
            
            console.log(`Aggregated stock for ${Object.keys(this.stockData).length} unique product numbers`);
            this.hasStockData = true;
            this.showNotification('Stock (SOH) data loaded successfully!', 'success');
            
        } catch (error) {
            console.error('Error processing stock file:', error);
            this.showNotification('Error loading stock file: ' + error.message, 'error');
            this.hasStockData = false;
        } finally {
            this.isLoading = false;
            this.render();
        }
    }

    async handleOpenSupplyFile(file) {
        console.log('Processing Production RRP4 (Open Supply) file...');
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
            
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            
            console.log(`Loaded ${jsonData.length} open supply records`);
            
            this.openSupplyData = {};
            
            jsonData.forEach(record => {
                const productNumber = this.toComparableString(record['Product Number']);
                const quantity = parseFloat(record['Receipt Quantity / Requirements Quantity']) || 0;
                
                if (productNumber) {
                    if (!this.openSupplyData[productNumber]) {
                        this.openSupplyData[productNumber] = 0;
                    }
                    this.openSupplyData[productNumber] += quantity;
                }
            });
            
            console.log(`Aggregated open supply for ${Object.keys(this.openSupplyData).length} unique product numbers`);
            this.hasOpenSupplyData = true;
            this.showNotification('Open Supply data loaded successfully!', 'success');
            
        } catch (error) {
            console.error('Error processing open supply file:', error);
            this.showNotification('Error loading open supply file: ' + error.message, 'error');
            this.hasOpenSupplyData = false;
        } finally {
            this.isLoading = false;
            this.render();
        }
    }

    async handleStockInTransitFile(file) {
        console.log('Processing Transport Receipts (Stock in Transit) file...');
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
            
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            
            console.log(`Loaded ${jsonData.length} stock in transit records`);
            
            this.stockInTransitData = {};
            
            jsonData.forEach(record => {
                const productNumber = this.toComparableString(record['PartCode']);
                const quantity = parseFloat(record['SupplierOrderQuantityOrdered']) || 0;
                
                if (productNumber) {
                    if (!this.stockInTransitData[productNumber]) {
                        this.stockInTransitData[productNumber] = 0;
                    }
                    this.stockInTransitData[productNumber] += quantity;
                }
            });
            
            console.log(`Aggregated stock in transit for ${Object.keys(this.stockInTransitData).length} unique product numbers`);
            this.hasStockInTransitData = true;
            this.showNotification('Stock in Transit data loaded successfully!', 'success');
            
        } catch (error) {
            console.error('Error processing stock in transit file:', error);
            this.showNotification('Error loading stock in transit file: ' + error.message, 'error');
            this.hasStockInTransitData = false;
        } finally {
            this.isLoading = false;
            this.render();
        }
    }

    calculateTotal(variant) {
        const soh = this.stockData[variant] || 0;
        const openSupply = this.openSupplyData[variant] || 0;
        const inTransit = this.stockInTransitData[variant] || 0;
        return soh + openSupply + inTransit;
    }

    async loadVariantCycleData(file) {
        console.log('Loading variant cycle data...');
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
            
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(worksheet);
            
            console.log(`Loaded ${data.length} cycle data records`);
            
            this.processCycleData(data);
            
            this.hasVariantCycleData = true;
            this.showNotification('Variant cycle data loaded successfully!', 'success');
            
        } catch (error) {
            console.error('Error loading cycle data:', error);
            this.showNotification('Error loading cycle data: ' + error.message, 'error');
            this.hasVariantCycleData = false;
        } finally {
            this.isLoading = false;
            this.render();
        }
    }
    
    processCycleData(data) {
        console.log('Processing cycle data...');
        
        this.variantCycleDates = {};
        
        const dfuColumn = 'DFU';
        const partCodeColumn = 'Part Code';
        const sosColumn = 'SOS';
        const eosColumn = 'EOS';
        const commentsColumn = 'Comments';
        
        if (data.length > 0) {
            const sampleRecord = data[0];
            const columns = Object.keys(sampleRecord);
            console.log('Cycle data columns:', columns);
            
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
            
            let processedCount = 0;
            data.forEach(record => {
                const dfuCode = record[actualDfuColumn];
                const partCode = record[actualPartCodeColumn];
                const sos = record[actualSosColumn];
                const eos = record[actualEosColumn];
                const comments = record[actualCommentsColumn];
                
                if (dfuCode && partCode) {
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
            
            let sheetName = 'Total Demand';
            if (!workbook.Sheets[sheetName]) {
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
        console.log('Processing data:', data.length, 'records');
        
        if (data.length === 0) {
            this.showNotification('No data found in the file', 'error');
            return;
        }
        
        const sampleRecord = data[0];
        console.log('Sample record:', sampleRecord);
        console.log('Available columns:', Object.keys(sampleRecord));
        
        const dfuColumn = 'DFU';
        const partNumberColumn = 'Product Number';
        const demandColumn = 'weekly fcst';
        const partDescriptionColumn = 'PartDescription';
        const plantLocationColumn = 'Production Plant';
        const productionLineColumn = 'Production Line';
        const calendarWeekColumn = 'Calendar.week';
        const sourceLocationColumn = 'Source Location';
        const weekNumberColumn = 'Week Number';
        
        this.availablePlantLocations = [...new Set(data.map(r => this.toComparableString(r[plantLocationColumn])))].filter(Boolean).sort();
        this.availableProductionLines = [...new Set(data.map(r => this.toComparableString(r[productionLineColumn])))].filter(Boolean).sort();
        
        let filteredData = data;
        
        if (this.selectedPlantLocation) {
            filteredData = filteredData.filter(record => 
                this.toComparableString(record[plantLocationColumn]) === this.selectedPlantLocation
            );
        }
        
        if (this.selectedProductionLine) {
            filteredData = filteredData.filter(record => 
                this.toComparableString(record[productionLineColumn]) === this.selectedProductionLine
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
                
                const totalDemand = partCodeRecords.reduce((sum, r) => {
                    const demand = parseFloat(r[demandColumn]) || 0;
                    return sum + demand;
                }, 0);
                
                const partDescription = partCodeRecords[0] ? partCodeRecords[0][partDescriptionColumn] : '';
                
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
            });
            
            const activeVariants = Object.keys(variantDemand);
            
            allDFUs[dfuCode] = {
                variants: activeVariants,
                variantDemand,
                totalRecords: records.length,
                dfuColumn,
                partNumberColumn,
                demandColumn,
                partDescriptionColumn,
                plantLocationColumn,
                productionLineColumn,
                calendarWeekColumn,
                sourceLocationColumn,
                weekNumberColumn,
                isCompleted: !!isCompleted,
                completionInfo: isCompleted || null,
                plantLocations: uniquePlants,
                productionLines: uniqueProductionLines,
                plantLocation: records[0] ? this.toComparableString(records[0][plantLocationColumn]) : null,
                productionLine: records[0] ? this.toComparableString(records[0][productionLineColumn]) : null
            };
        });
        
        this.multiVariantDFUs = allDFUs;
        this.filterDFUs();
    }
    
    filterDFUs() {
        const searchTerm = this.searchTerm.toLowerCase().trim();
        
        if (!searchTerm) {
            this.filteredDFUs = { ...this.multiVariantDFUs };
        } else {
            this.filteredDFUs = {};
            
            Object.keys(this.multiVariantDFUs).forEach(dfuCode => {
                const dfu = this.multiVariantDFUs[dfuCode];
                const dfuCodeLower = dfuCode.toLowerCase();
                const variantsLower = dfu.variants.map(v => v.toLowerCase());
                
                if (dfuCodeLower.includes(searchTerm) || 
                    variantsLower.some(v => v.includes(searchTerm))) {
                    this.filteredDFUs[dfuCode] = dfu;
                }
            });
        }
        
        this.render();
    }
    
    async exportData() {
        try {
            console.log('Exporting data...');
            
            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.json_to_sheet(this.rawData);
            XLSX.utils.book_append_sheet(wb, ws, 'Updated Data');
            XLSX.writeFile(wb, `DFU_Updated_${new Date().toISOString().slice(0,10)}.xlsx`);
            
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
                                            <li><strong>Production Plant</strong> - Plant location codes</li>
                                            <li><strong>Production Line</strong> - Production line codes</li>
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
                                    
                                    <div class="border-t pt-4 mt-4">
                                        <h3 class="text-sm font-medium text-gray-700 mb-2">Optional: Upload Supply Chain Files</h3>
                                        
                                        <div class="space-y-3">
                                            <div>
                                                <label class="text-xs font-medium text-gray-600">Stock RRP4 (SOH)</label>
                                                <input type="file" accept=".xlsx,.xls" class="file-input" id="stockFileInput">
                                                <p class="text-xs text-gray-500 mt-1">
                                                    ${this.hasStockData ? '✓ SOH Data Loaded' : 'Upload file with Product Number and Stock columns'}
                                                </p>
                                            </div>
                                            
                                            <div>
                                                <label class="text-xs font-medium text-gray-600">Production RRP4 (Open Supply)</label>
                                                <input type="file" accept=".xlsx,.xls" class="file-input" id="openSupplyFileInput">
                                                <p class="text-xs text-gray-500 mt-1">
                                                    ${this.hasOpenSupplyData ? '✓ Open Supply Data Loaded' : 'Upload file with Product Number and Receipt Quantity columns'}
                                                </p>
                                            </div>
                                            
                                            <div>
                                                <label class="text-xs font-medium text-gray-600">Transport Receipts (Stock in Transit)</label>
                                                <input type="file" accept=".xlsx,.xls" class="file-input" id="stockInTransitFileInput">
                                                <p class="text-xs text-gray-500 mt-1">
                                                    ${this.hasStockInTransitData ? '✓ Stock in Transit Data Loaded' : 'Upload file with PartCode and SupplierOrderQuantityOrdered columns'}
                                                </p>
                                            </div>
                                        </div>
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
                        if (file) this.loadVariantCycleData(file);
                    });
                }
                
                const stockFileInput = document.getElementById('stockFileInput');
                if (stockFileInput) {
                    stockFileInput.addEventListener('change', (e) => {
                        const file = e.target.files[0];
                        if (file) this.handleStockFile(file);
                    });
                }
                
                const openSupplyFileInput = document.getElementById('openSupplyFileInput');
                if (openSupplyFileInput) {
                    openSupplyFileInput.addEventListener('change', (e) => {
                        const file = e.target.files[0];
                        if (file) this.handleOpenSupplyFile(file);
                    });
                }
                
                const stockInTransitFileInput = document.getElementById('stockInTransitFileInput');
                if (stockInTransitFileInput) {
                    stockInTransitFileInput.addEventListener('change', (e) => {
                        const file = e.target.files[0];
                        if (file) this.handleStockInTransitFile(file);
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
                <p class="text-gray-600">
                    Managing ${totalDFUs} DFUs (${multiVariantCount} with multiple variants, ${totalDFUs - multiVariantCount} single variant)
                </p>
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
                    <div class="relative">
                        <select class="px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent" id="productionLineFilter">
                            <option value="">All Production Lines</option>
                            ${this.availableProductionLines.map(line => `
                                <option value="${line}" ${this.selectedProductionLine === line ? 'selected' : ''}>
                                    Line ${line}
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
                            <svg class="w-5 h-5 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2" />
                            </svg>
                            All DFUs (${Object.keys(this.filteredDFUs).length})
                        </h3>
                        <div class="space-y-2 h-full overflow-y-auto" style="max-height: calc(100vh - 300px);">
                            ${Object.keys(this.filteredDFUs).map(dfuCode => {
                                const dfu = this.filteredDFUs[dfuCode];
                                const isSelected = this.selectedDFU === dfuCode;
                                
                                return `
                                    <div class="dfu-card ${isSelected ? 'selected' : ''}" data-dfu="${dfuCode}">
                                        <div class="flex justify-between items-center">
                                            <div>
                                                <h4 class="font-medium">DFU: ${dfuCode}</h4>
                                                <p class="text-sm text-gray-600">
                                                    ${dfu.variants.length} variant${dfu.variants.length > 1 ? 's' : ''} • ${dfu.totalRecords} records
                                                    ${dfu.isCompleted ? '• ✓ Complete' : ''}
                                                </p>
                                            </div>
                                        </div>
                                    </div>
                                `;
                            }).join('')}
                        </div>
                    </div>
                    
                    <div class="bg-gray-50 rounded-lg p-6">
                        ${this.renderDFUDetails()}
                    </div>
                </div>
        `;
        
        this.attachEventListeners();
        this.ensureGranularContainers();
    }
    
    renderDFUDetails() {
        if (!this.selectedDFU || !this.multiVariantDFUs[this.selectedDFU]) {
            return `
                <div class="text-center py-12 text-gray-500">
                    <svg class="w-12 h-12 mx-auto mb-2 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 15l-2 5L9 9l11 4-5 2zm0 0l5 5M7.188 2.239l.777 2.897M5.136 7.965l-2.898-.777M13.95 4.05l-2.122 2.122m-5.657 5.656l-2.12 2.122" />
                    </svg>
                    <p>Select a DFU from the list to view details</p>
                </div>
            `;
        }

        const hasSupplyChainData = this.hasStockData || this.hasOpenSupplyData || this.hasStockInTransitData;
        
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
                            <span class="ml-2 px-2 py-1 text-xs bg-green-100 text-green-800 rounded-full">
                                ✓ Transfer Complete
                            </span>
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
                    <!-- Current Variant Status -->
                    <div class="mb-6">
                        <h4 class="font-semibold text-gray-800 mb-3">Current Variant Status</h4>
                        <div class="space-y-3">
                            ${this.multiVariantDFUs[this.selectedDFU].variants.map(variant => {
                                const demandData = this.multiVariantDFUs[this.selectedDFU].variantDemand[variant];
                                const soh = this.stockData[variant] || 0;
                                const openSupply = this.openSupplyData[variant] || 0;
                                const inTransit = this.stockInTransitData[variant] || 0;
                                const total = soh + openSupply + inTransit;
                                
                                return `
                                    <div class="border rounded-lg p-3 bg-white">
                                        <div class="flex justify-between items-start">
                                            <div class="flex-1">
                                                <h5 class="font-medium text-gray-800">Part: ${variant}</h5>
                                                <p class="text-xs text-gray-500 mb-1 max-w-md break-words">${demandData?.partDescription || 'Description not available'}</p>
                                                <p class="text-sm text-gray-600">${demandData?.recordCount || 0} records</p>
                                            </div>
                                            <div class="text-right">
                                                <p class="font-medium text-gray-800">${this.formatNumber(demandData?.totalDemand || 0)}</p>
                                                <p class="text-sm text-gray-600">consolidated demand</p>
                                                ${hasSupplyChainData ? `
                                                    <div class="mt-3 pt-3 border-t space-y-1">
                                                        ${this.hasStockData ? `
                                                            <div class="flex justify-between text-xs">
                                                                <span class="text-blue-600">SOH:</span>
                                                                <span class="font-medium text-blue-600">${this.formatNumber(soh)}</span>
                                                            </div>
                                                        ` : ''}
                                                        ${this.hasOpenSupplyData ? `
                                                            <div class="flex justify-between text-xs">
                                                                <span class="text-green-600">Open Supply:</span>
                                                                <span class="font-medium text-green-600">${this.formatNumber(openSupply)}</span>
                                                            </div>
                                                        ` : ''}
                                                        ${this.hasStockInTransitData ? `
                                                            <div class="flex justify-between text-xs">
                                                                <span class="text-purple-600">In Transit:</span>
                                                                <span class="font-medium text-purple-600">${this.formatNumber(inTransit)}</span>
                                                            </div>
                                                        ` : ''}
                                                        <div class="flex justify-between text-sm font-semibold pt-1 border-t">
                                                            <span class="text-gray-800">Total:</span>
                                                            <span class="text-gray-800">${this.formatNumber(total)}</span>
                                                        </div>
                                                    </div>
                                                ` : ''}
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
                            <p class="text-sm text-purple-800 mt-3">
                                Selected: All variants will transfer to <strong>${this.bulkTransfers[this.selectedDFU]}</strong>
                            </p>
                        ` : ''}
                    </div>
                    
                    <!-- Individual Transfer Section -->
                    ${this.renderIndividualTransferSection()}
                    
                    <!-- Action Buttons Container -->
                    <div class="action-buttons-container">
                        ${this.renderActionButtons()}
                    </div>
                `}
                
                ${this.multiVariantDFUs[this.selectedDFU].isCompleted && this.multiVariantDFUs[this.selectedDFU].completionInfo ? `
                    <div class="mt-4 p-4 bg-green-50 rounded-lg border border-green-200">
                        <h4 class="font-semibold text-green-800 mb-2">Transfer Details</h4>
                        <div class="text-sm text-green-700 space-y-1">
                            <p><strong>Type:</strong> ${this.multiVariantDFUs[this.selectedDFU].completionInfo.type}</p>
                            <p><strong>Time:</strong> ${this.multiVariantDFUs[this.selectedDFU].completionInfo.timestamp}</p>
                            ${this.multiVariantDFUs[this.selectedDFU].completionInfo.targetVariant ? `
                                <p><strong>Target:</strong> ${this.multiVariantDFUs[this.selectedDFU].completionInfo.targetVariant}</p>
                            ` : ''}
                        </div>
                        <button class="btn btn-secondary mt-3 text-sm" id="undoTransferBtn">
                            Undo Transfer
                        </button>
                    </div>
                ` : ''}
                
                <div class="mt-6 p-4 bg-blue-50 rounded-lg border border-blue-200 text-sm text-gray-700">
                    <h4 class="font-semibold text-blue-800 mb-2">How to Use:</h4>
                    <ul class="list-disc list-inside space-y-1">
                        <li><strong>Bulk Transfer:</strong> Select one target variant to transfer all variants to</li>
                        <li><strong>Individual Transfer:</strong> Set specific targets for each variant using dropdowns</li>
                        <li><strong>Granular Transfer:</strong> Expand any variant and select specific weeks to transfer partial demand</li>
                        <li><strong>Execute:</strong> Click "Execute Transfer" to apply your chosen transfers</li>
                        <li><strong>Export:</strong> Export the updated data when you're done with all transfers</li>
                    </ul>
                </div>
            </div>
        `;
        
        this.attachEventListeners();
        this.ensureGranularContainers();
    }
    
    renderIndividualTransferSection() {
        if (!this.selectedDFU || !this.multiVariantDFUs[this.selectedDFU]) return '';

        const hasSupplyChainData = this.hasStockData || this.hasOpenSupplyData || this.hasStockInTransitData;
        
        return `
            <div class="mb-6">
                <h4 class="font-semibold text-gray-800 mb-3">Individual Transfers (Variant → Specific Target)</h4>
                <div class="space-y-4">
                    ${this.multiVariantDFUs[this.selectedDFU].variants.map(variant => {
                        const demandData = this.multiVariantDFUs[this.selectedDFU].variantDemand[variant];
                        const currentTransfer = this.transfers[this.selectedDFU]?.[variant];
                        const soh = this.stockData[variant] || 0;
                        const openSupply = this.openSupplyData[variant] || 0;
                        const inTransit = this.stockInTransitData[variant] || 0;
                        const total = soh + openSupply + inTransit;
                        
                        return `
                            <div class="border rounded-lg p-4 bg-gray-50">
                                <div class="flex justify-between items-center mb-3">
                                    <div class="flex-1">
                                        <h5 class="font-medium text-gray-800">Part: ${variant}</h5>
                                        <p class="text-xs text-gray-500 mb-1 max-w-md break-words">${demandData?.partDescription || 'Description not available'}</p>
                                        <p class="text-sm text-gray-600">
                                            ${demandData?.recordCount || 0} records • ${this.formatNumber(demandData?.totalDemand || 0)} total demand
                                        </p>
                                        ${hasSupplyChainData ? `
                                            <div class="mt-2 text-xs space-y-0.5">
                                                ${this.hasStockData ? `<p class="text-blue-600">SOH: ${this.formatNumber(soh)}</p>` : ''}
                                                ${this.hasOpenSupplyData ? `<p class="text-green-600">Open Supply: ${this.formatNumber(openSupply)}</p>` : ''}
                                                ${this.hasStockInTransitData ? `<p class="text-purple-600">In Transit: ${this.formatNumber(inTransit)}</p>` : ''}
                                                <p class="font-semibold text-gray-800">Total: ${this.formatNumber(total)}</p>
                                            </div>
                                        ` : ''}
                                        ${(() => {
                                            const cycleData = this.getCycleDataForVariant(this.selectedDFU, variant);
                                            if (cycleData) {
                                                return `
                                                    <div class="mt-2 text-xs space-y-0.5 pt-2 border-t">
                                                        <p class="text-blue-600"><strong>SOS:</strong> ${cycleData.sos}</p>
                                                        <p class="text-red-600"><strong>EOS:</strong> ${cycleData.eos}</p>
                                                        ${cycleData.comments ? `<p class="text-gray-600"><strong>Comments:</strong> ${cycleData.comments}</p>` : ''}
                                                    </div>
                                                `;
                                            }
                                            return '';
                                        })()}
                                    </div>
                                    <div class="ml-4">
                                        <select 
                                            class="px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                                            data-source-variant="${variant}"
                                        >
                                            <option value="">Transfer to...</option>
                                            ${this.multiVariantDFUs[this.selectedDFU].variants.map(targetVariant => 
                                                `<option value="${targetVariant}" ${currentTransfer === targetVariant ? 'selected' : ''}>${targetVariant}</option>`
                                            ).join('')}
                                        </select>
                                    </div>
                                </div>
                                
                                <!-- Granular Transfer Container -->
                                <div id="granular-${variant}" class="granular-section"></div>
                            </div>
                        `;
                    }).join('')}
                </div>
            </div>
        `;
    }
    
    renderActionButtons() {
        if (!this.selectedDFU) return '';
        
        const hasTransfers = ((this.transfers[this.selectedDFU] && Object.keys(this.transfers[this.selectedDFU]).length > 0) || 
                             this.bulkTransfers[this.selectedDFU] || 
                             (this.granularTransfers[this.selectedDFU] && Object.keys(this.granularTransfers[this.selectedDFU]).length > 0));
        
        const hasExecutionSummary = this.lastExecutionSummary[this.selectedDFU];
        
        if (hasTransfers) {
            return `
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
                    </div>
                    <div class="flex gap-2">
                        <button class="btn btn-success" id="executeBtn">Execute Transfer</button>
                        <button class="btn btn-secondary" id="cancelBtn">Cancel</button>
                    </div>
                </div>
            `;
        } else if (hasExecutionSummary) {
            return `
                <div class="p-3 bg-gray-50 rounded-lg">
                    <div class="text-sm text-gray-700">
                        <h5 class="font-semibold mb-2 text-gray-800">Last Execution Summary:</h5>
                        <p><strong>Type:</strong> ${this.lastExecutionSummary[this.selectedDFU].type}</p>
                        <p><strong>Time:</strong> ${this.lastExecutionSummary[this.selectedDFU].timestamp}</p>
                        <p><strong>Result:</strong> ${this.lastExecutionSummary[this.selectedDFU].message}</p>
                    </div>
                </div>
            `;
        }
        
        return '';
    }

    renderGranularSection(sourceVariant, targetVariant) {
        if (!this.selectedDFU || !targetVariant) return '';
        
        const demandData = this.multiVariantDFUs[this.selectedDFU].variantDemand[sourceVariant];
        const weeklyRecords = demandData?.weeklyRecords || {};
        
        if (Object.keys(weeklyRecords).length === 0) {
            return '<p class="text-sm text-gray-500 mt-2">No weekly data available</p>';
        }
        
        return `
            <div class="mt-3 pt-3 border-t">
                <h6 class="text-sm font-medium text-gray-700 mb-2">Granular Transfer (Select Specific Weeks)</h6>
                <div class="space-y-2 max-h-48 overflow-y-auto">
                    ${Object.keys(weeklyRecords).sort().map(weekKey => {
                        const weekData = weeklyRecords[weekKey];
                        const isSelected = this.granularTransfers[this.selectedDFU]?.[sourceVariant]?.[targetVariant]?.[weekKey]?.selected || false;
                        const customQty = this.granularTransfers[this.selectedDFU]?.[sourceVariant]?.[targetVariant]?.[weekKey]?.customQuantity;
                        
                        return `
                            <div class="flex items-center gap-2 text-sm p-2 bg-white rounded">
                                <input 
                                    type="checkbox" 
                                    ${isSelected ? 'checked' : ''} 
                                    data-granular-toggle
                                    data-dfu="${this.selectedDFU}"
                                    data-source="${sourceVariant}"
                                    data-target="${targetVariant}"
                                    data-week="${weekKey}"
                                    class="rounded"
                                >
                                <span class="flex-1">Week ${weekData.weekNumber} (${weekData.sourceLocation}): ${this.formatNumber(weekData.demand)}</span>
                                <input 
                                    type="number" 
                                    placeholder="Custom qty"
                                    class="w-24 px-2 py-1 border rounded text-xs"
                                    value="${customQty !== null && customQty !== undefined ? customQty : ''}"
                                    ${!isSelected ? 'disabled' : ''}
                                    data-granular-qty
                                    data-dfu="${this.selectedDFU}"
                                    data-source="${sourceVariant}"
                                    data-target="${targetVariant}"
                                    data-week="${weekKey}"
                                >
                            </div>
                        `;
                    }).join('')}
                </div>
            </div>
        `;
    }
    
    ensureGranularContainers() {
        if (this.selectedDFU && this.multiVariantDFUs[this.selectedDFU]) {
            this.multiVariantDFUs[this.selectedDFU].variants.forEach(variant => {
                const selectElement = document.querySelector(`[data-source-variant="${variant}"]`);
                
                if (selectElement) {
                    const parentDiv = selectElement.closest('.border.rounded-lg');
                    
                    if (parentDiv) {
                        let granularContainer = document.getElementById(`granular-${variant}`);
                        
                        if (!granularContainer) {
                            granularContainer = document.createElement('div');
                            granularContainer.id = `granular-${variant}`;
                            granularContainer.className = 'border-t pt-3 mt-3 granular-section';
                            granularContainer.style.minHeight = '10px';
                            
                            parentDiv.appendChild(granularContainer);
                        }
                    }
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
        
        const productionLineFilter = document.getElementById('productionLineFilter');
        if (productionLineFilter) {
            productionLineFilter.addEventListener('change', (e) => {
                this.filterByProductionLine(e.target.value);
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
        
        const addVariantBtn = document.getElementById('addVariantBtn');
        if (addVariantBtn) {
            addVariantBtn.addEventListener('click', () => this.addManualVariant(this.selectedDFU));
        }
        
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
        
        const stockFileInput = document.getElementById('stockFileInput');
        if (stockFileInput) {
            stockFileInput.addEventListener('change', (e) => {
                const file = e.target.files[0];
                if (file) {
                    console.log('Stock file selected:', file.name);
                    this.handleStockFile(file);
                }
            });
        }

        const openSupplyFileInput = document.getElementById('openSupplyFileInput');
        if (openSupplyFileInput) {
            openSupplyFileInput.addEventListener('change', (e) => {
                const file = e.target.files[0];
                if (file) {
                    console.log('Open Supply file selected:', file.name);
                    this.handleOpenSupplyFile(file);
                }
            });
        }

        const stockInTransitFileInput = document.getElementById('stockInTransitFileInput');
        if (stockInTransitFileInput) {
            stockInTransitFileInput.addEventListener('change', (e) => {
                const file = e.target.files[0];
                if (file) {
                    console.log('Stock in Transit file selected:', file.name);
                    this.handleStockInTransitFile(file);
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
                
                if (targetVariant && targetVariant !== sourceVariant) {
                    this.setIndividualTransfer(this.selectedDFU, sourceVariant, targetVariant);
                    
                    const granularSection = document.getElementById(`granular-${sourceVariant}`);
                    
                    if (granularSection) {
                        granularSection.innerHTML = this.renderGranularSection(sourceVariant, targetVariant);
                        this.attachGranularEventListeners();
                    }
                } else if (!targetVariant) {
                    this.setIndividualTransfer(this.selectedDFU, sourceVariant, '');
                }
            });
        });
        
        this.attachGranularEventListeners();
        
        setTimeout(() => {
            const searchInput = document.getElementById('searchInput');
            if (searchInput && document.activeElement !== searchInput && this.searchTerm) {
                searchInput.focus();
                searchInput.setSelectionRange(searchInput.value.length, searchInput.value.length);
            }
        }, 10);
    }
    
    filterByPlantLocation(plantLocation) {
        this.selectedPlantLocation = plantLocation;
        this.multiVariantDFUs = {};
        this.filteredDFUs = {};
        this.processMultiVariantDFUs(this.rawData);
        this.render();
    }
    
    filterByProductionLine(productionLine) {
        this.selectedProductionLine = productionLine;
        this.multiVariantDFUs = {};
        this.filteredDFUs = {};
        this.processMultiVariantDFUs(this.rawData);
        this.render();
    }
    
    selectDFU(dfuCode) {
        this.selectedDFU = this.toComparableString(dfuCode);
        this.render();
    }
    
    selectBulkTarget(dfuCode, targetVariant) {
        const dfuStr = this.toComparableString(dfuCode);
        const targetStr = this.toComparableString(targetVariant);
        
        this.bulkTransfers[dfuStr] = targetStr;
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
        
        if (this.granularTransfers[dfuStr] && this.granularTransfers[dfuStr][sourceStr]) {
            delete this.granularTransfers[dfuStr][sourceStr];
        }
        
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
        
        const current = this.granularTransfers[dfuStr][sourceStr][targetStr][weekKey];
        if (current && current.selected) {
            delete this.granularTransfers[dfuStr][sourceStr][targetStr][weekKey];
        } else {
            this.granularTransfers[dfuStr][sourceStr][targetStr][weekKey] = {
                selected: true,
                customQuantity: null
            };
        }
        
        if (this.transfers[dfuStr] && this.transfers[dfuStr][sourceStr]) {
            delete this.transfers[dfuStr][sourceStr];
        }
        
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
            
            this.updateActionButtonsOnly();
        }
    }
    
    updateActionButtonsOnly() {
        const actionButtonsContainer = document.querySelector('.action-buttons-container');
        if (actionButtonsContainer && this.selectedDFU) {
            const hasTransfers = ((this.transfers[this.selectedDFU] && Object.keys(this.transfers[this.selectedDFU]).length > 0) || 
                                 this.bulkTransfers[this.selectedDFU] || 
                                 (this.granularTransfers[this.selectedDFU] && Object.keys(this.granularTransfers[this.selectedDFU]).length > 0));
            
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
                        </div>
                        <div class="flex gap-2">
                            <button class="btn btn-success" id="executeBtn">Execute Transfer</button>
                            <button class="btn btn-secondary" id="cancelBtn">Cancel</button>
                        </div>
                    </div>
                `;
                
                const executeBtn = document.getElementById('executeBtn');
                if (executeBtn) {
                    executeBtn.addEventListener('click', () => this.executeTransfer(this.selectedDFU));
                }
                
                const cancelBtn = document.getElementById('cancelBtn');
                if (cancelBtn) {
                    cancelBtn.addEventListener('click', () => this.cancelTransfer(this.selectedDFU));
                }
            } else if (hasExecutionSummary) {
                const summary = this.lastExecutionSummary[this.selectedDFU];
                actionButtonsContainer.innerHTML = `
                    <div class="p-3 bg-gray-50 rounded-lg">
                        <div class="text-sm text-gray-700">
                            <h5 class="font-semibold mb-2 text-gray-800">Last Execution:</h5>
                            <p><strong>Type:</strong> ${summary.type}</p>
                            <p><strong>Time:</strong> ${summary.timestamp}</p>
                            <p><strong>Status:</strong> ${summary.message}</p>
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
    
    addManualVariant(dfuCode) {
        const variantCode = prompt('Enter the new variant code:');
        if (!variantCode || !variantCode.trim()) return;
        
        const dfuStr = this.toComparableString(dfuCode);
        const dfuData = this.multiVariantDFUs[dfuStr];
        
        if (!dfuData) {
            this.showNotification('DFU not found', 'error');
            return;
        }
        
        const trimmedVariant = variantCode.trim();
        if (dfuData.variants.some(v => this.toComparableString(v) === this.toComparableString(trimmedVariant))) {
            this.showNotification('Variant already exists', 'error');
            return;
        }
        
        const dfuRecords = this.rawData.filter(r => 
            this.toComparableString(r['DFU']) === dfuStr
        );
        
        const newRecords = [];
        const processedCombos = new Set();
        
        dfuRecords.forEach(record => {
            const weekNum = record['Week Number'];
            const sourceLoc = record['Source Location'];
            const comboKey = `${weekNum}_${sourceLoc}`;
            
            if (!processedCombos.has(comboKey)) {
                processedCombos.add(comboKey);
                
                const newRecord = { ...record };
                newRecord['Product Number'] = trimmedVariant;
                newRecord['weekly fcst'] = 0;
                newRecord['PartDescription'] = 'Manually added variant';
                newRecord['Transfer History'] = `Manually added on ${new Date().toLocaleString()}`;
                
                newRecords.push(newRecord);
            }
        });
        
        if (newRecords.length === 0 && dfuRecords.length > 0) {
            const templateRecord = dfuRecords[0];
            const newRecord = { ...templateRecord };
            newRecord['Product Number'] = trimmedVariant;
            newRecord['weekly fcst'] = 0;
            newRecord['PartDescription'] = 'Manually added variant';
            newRecord['Transfer History'] = `Manually added on ${new Date().toLocaleString()}`;
            newRecords.push(newRecord);
        }
        
        this.rawData.push(...newRecords);
        this.processMultiVariantDFUs(this.rawData);
        
        this.showNotification(`Variant ${trimmedVariant} added successfully with ${newRecords.length} records`);
        this.render();
    }
    
    executeTransfer(dfuCode) {
        const dfuStr = this.toComparableString(dfuCode);
        
        if (!this.multiVariantDFUs[dfuStr]) {
            this.showNotification('DFU not found', 'error');
            return;
        }
        
        const dfuData = this.multiVariantDFUs[dfuStr];
        const { dfuColumn, partNumberColumn, demandColumn, weekNumberColumn, sourceLocationColumn } = dfuData;
        
        const dfuRecords = this.rawData.filter(record => this.toComparableString(record[dfuColumn]) === dfuStr);
        const originalVariants = new Set(dfuData.variants);
        
        if (this.bulkTransfers[dfuStr]) {
            const targetVariant = this.bulkTransfers[dfuStr];
            
            dfuRecords.forEach(record => {
                const currentPartNumber = this.toComparableString(record[partNumberColumn]);
                if (currentPartNumber !== targetVariant) {
                    record['Transfer History'] = `Bulk transferred from ${currentPartNumber} to ${targetVariant} on ${new Date().toLocaleString()}`;
                    record[partNumberColumn] = targetVariant;
                }
            });
            
            this.consolidateRecords(dfuStr, originalVariants);
            
            this.completedTransfers[dfuStr] = {
                type: 'bulk',
                targetVariant: targetVariant,
                timestamp: new Date().toLocaleString()
            };
            
            this.lastExecutionSummary[dfuStr] = {
                type: 'Bulk Transfer',
                message: `All variants transferred to ${targetVariant}`,
                timestamp: new Date().toLocaleString()
            };
            
            this.showNotification(`Bulk transfer completed for DFU ${dfuStr}: All variants → ${targetVariant}`);
        }
        
        else if (this.transfers[dfuStr] && Object.keys(this.transfers[dfuStr]).length > 0) {
            const individualTransfers = this.transfers[dfuStr];
            
            Object.keys(individualTransfers).forEach(sourceVariant => {
                const targetVariant = individualTransfers[sourceVariant];
                
                if (sourceVariant !== targetVariant) {
                    dfuRecords.forEach(record => {
                        const currentPartNumber = this.toComparableString(record[partNumberColumn]);
                        if (currentPartNumber === sourceVariant) {
                            record['Transfer History'] = `Transferred from ${sourceVariant} to ${targetVariant} on ${new Date().toLocaleString()}`;
                            record[partNumberColumn] = targetVariant;
                        }
                    });
                }
            });
            
            this.consolidateRecords(dfuStr, originalVariants);
            
            this.completedTransfers[dfuStr] = {
                type: 'individual',
                transfers: individualTransfers,
                timestamp: new Date().toLocaleString()
            };
            
            const executionMessage = Object.keys(individualTransfers)
                .filter(src => src !== individualTransfers[src])
                .map(src => `${src} → ${individualTransfers[src]}`)
                .join(', ');
            
            this.lastExecutionSummary[dfuStr] = {
                type: 'Individual Transfer',
                message: executionMessage,
                timestamp: new Date().toLocaleString()
            };
            
            this.showNotification(`Individual transfers completed for DFU ${dfuStr}: ${executionMessage}`);
        }
        
        else if (this.granularTransfers[dfuStr] && Object.keys(this.granularTransfers[dfuStr]).length > 0) {
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
                                weekTransfer.customQuantity : originalDemand;
                            
                            let targetRecord = dfuRecords.find(r => 
                                this.toComparableString(r[partNumberColumn]) === targetVariant &&
                                this.toComparableString(r[weekNumberColumn]) === weekNumber &&
                                this.toComparableString(r[sourceLocationColumn]) === sourceLocation
                            );
                            
                            if (targetRecord) {
                                const currentTargetDemand = parseFloat(targetRecord[demandColumn]) || 0;
                                targetRecord[demandColumn] = currentTargetDemand + transferAmount;
                                targetRecord['Transfer History'] = `Received ${transferAmount} from ${sourceVariant} (granular) on ${new Date().toLocaleString()}`;
                            } else {
                                targetRecord = { ...sourceRecord };
                                targetRecord[partNumberColumn] = targetVariant;
                                targetRecord[demandColumn] = transferAmount;
                                targetRecord['Transfer History'] = `Received ${transferAmount} from ${sourceVariant} (granular) on ${new Date().toLocaleString()}`;
                                this.rawData.push(targetRecord);
                            }
                            
                            sourceRecord[demandColumn] = originalDemand - transferAmount;
                            sourceRecord['Transfer History'] = `Transferred ${transferAmount} to ${targetVariant} (granular) on ${new Date().toLocaleString()}`;
                            
                            granularTransferCount++;
                        }
                    });
                });
            });
            
            this.completedTransfers[dfuStr] = {
                type: 'granular',
                transferCount: granularTransferCount,
                timestamp: new Date().toLocaleString()
            };
            
            this.lastExecutionSummary[dfuStr] = {
                type: 'Granular Transfer',
                message: `${granularTransferCount} week-specific transfers completed`,
                timestamp: new Date().toLocaleString()
            };
            
            this.showNotification(`Granular transfer completed for DFU ${dfuStr}: ${granularTransferCount} weeks transferred`);
        }
        
        delete this.transfers[dfuStr];
        delete this.bulkTransfers[dfuStr];
        delete this.granularTransfers[dfuStr];
        
        this.processMultiVariantDFUs(this.rawData);
        
        const currentSelection = this.selectedDFU;
        this.selectedDFU = null;
        this.render();
        
        setTimeout(() => {
            this.selectedDFU = currentSelection;
            this.forceUIRefresh();
        }, 300);
    }
    
    consolidateRecords(dfuCode, originalVariants = null) {
        const dfuStr = this.toComparableString(dfuCode);
        
        const currentDFUData = this.multiVariantDFUs[dfuStr] || Object.values(this.multiVariantDFUs)[0];
        if (!currentDFUData) return;
        
        const { dfuColumn, partNumberColumn, demandColumn, weekNumberColumn, sourceLocationColumn } = currentDFUData;
        
        const allRecords = this.rawData;
        const dfuRecords = allRecords.filter(record => this.toComparableString(record[dfuColumn]) === dfuStr);
        
        const allPartNumbers = originalVariants || new Set();
        if (!originalVariants) {
            dfuRecords.forEach(record => {
                allPartNumbers.add(this.toComparableString(record[partNumberColumn]));
            });
        }
        
        const consolidatedMap = new Map();
        
        dfuRecords.forEach(record => {
            const partNumber = this.toComparableString(record[partNumberColumn]);
            const weekNumber = this.toComparableString(record[weekNumberColumn]);
            const sourceLocation = this.toComparableString(record[sourceLocationColumn]);
            const key = `${partNumber}_${weekNumber}_${sourceLocation}`;
            
            if (!consolidatedMap.has(key)) {
                consolidatedMap.set(key, { ...record });
            } else {
                const existing = consolidatedMap.get(key);
                const currentDemand = parseFloat(existing[demandColumn]) || 0;
                const additionalDemand = parseFloat(record[demandColumn]) || 0;
                existing[demandColumn] = currentDemand + additionalDemand;
                
                if (record['Transfer History']) {
                    existing['Transfer History'] = (existing['Transfer History'] || '') + '; ' + record['Transfer History'];
                }
            }
        });
        
        this.rawData = this.rawData.filter(record => 
            this.toComparableString(record[dfuColumn]) !== dfuStr
        );
        
        consolidatedMap.forEach(record => {
            this.rawData.push(record);
        });
    }
    
    forceUIRefresh() {
        const app = document.getElementById('app');
        const currentSearch = this.searchTerm;
        
        app.innerHTML = '<div class="max-w-6xl mx-auto p-6 bg-white min-h-screen"><div class="text-center py-12"><div class="loading-spinner mb-2"></div><p>Refreshing interface...</p></div></div>';
        
        setTimeout(() => {
            this.searchTerm = currentSearch;
            this.render();
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
        
        this.showNotification(`Transfer undone for DFU ${dfuStr}. Original data restored.`);
        this.render();
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
        new DemandTransferApp();
    }
});