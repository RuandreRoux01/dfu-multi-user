// DFU Demand Transfer Management Application
// Version: 2.21.1 - Build: 2025-10-27-transfer-fix
// Fixed: Granular and bulk transfers now execute correctly

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
        
        this.stockData = {};
        this.openSupplyData = {};
        this.transitData = {};
        this.hasStockData = false;
        this.hasOpenSupplyData = false;
        this.hasTransitData = false;
        
        this.init();
    }
    
    init() {
        console.log('🚀 DFU Transfer App v2.21.1 - Transfers Fixed');
        this.loadSupplyChainDataFromServer();
        this.setupSupplyChainSocketListeners();
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
    
    async loadStockFile(file) {
        console.log('Loading Stock (SOH) file...');
        try {
            const arrayBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(worksheet);
            
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
            await this.saveSupplyChainDataToServer('stock', this.stockData);
            this.showNotification('Stock (SOH) data loaded successfully', 'success');
            this.render();
        } catch (error) {
            console.error('Error loading stock file:', error);
            this.showNotification('Error loading stock file: ' + error.message, 'error');
        }
    }
    
    async loadOpenSupplyFile(file) {
        console.log('Loading Open Supply file...');
        try {
            const arrayBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(worksheet);
            
            if (data.length > 0) {
                console.log('Available columns:', Object.keys(data[0]));
            }
            
            this.openSupplyData = {};
            let processedCount = 0;
            
            data.forEach(row => {
                const partCode = this.toComparableString(
                    row['Product Number'] || row['ProductNumber'] || row['PartCode'] || 
                    row['Part Code'] || row['Material'] || row['Material Number']
                );
                
                const openSupply = parseFloat(
                    row['Receipt Quantity / Requirements Quantity'] ||
                    row['Receipt Quantity'] || row['Requirements Quantity'] || 
                    row['ReceiptQuantity'] || row['RequirementsQuantity'] ||
                    row['OpenSupply'] || row['Open Supply'] || row['Quantity'] ||
                    row['Qty'] || row['Order Quantity'] || row['OrderQuantity'] || 0
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
            console.log('Open Supply processed:', Object.keys(this.openSupplyData).length, 'parts');
            await this.saveSupplyChainDataToServer('openSupply', this.openSupplyData);
            this.showNotification(`Open Supply data loaded: ${processedCount} records`, 'success');
            this.render();
        } catch (error) {
            console.error('Error loading Open Supply file:', error);
            this.showNotification('Error: ' + error.message, 'error');
        }
    }
    
    async loadTransitFile(file) {
        console.log('Loading Transit file...');
        try {
            const arrayBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(worksheet);
            
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
            console.log('Transit data processed:', Object.keys(this.transitData).length, 'parts');
            await this.saveSupplyChainDataToServer('transit', this.transitData);
            this.showNotification('Transit data loaded successfully', 'success');
            this.render();
        } catch (error) {
            console.error('Error loading Transit file:', error);
            this.showNotification('Error: ' + error.message, 'error');
        }
    }
    
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
    
    async saveSupplyChainDataToServer(type, data) {
        try {
            const response = await fetch('/api/upload-supply-chain', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    type,
                    data,
                    userName: window.userName || 'Unknown'
                })
            });
            
            const result = await response.json();
            if (result.success) {
                console.log(`[SYNC] ${type} saved:`, result.itemCount, 'items');
            }
        } catch (error) {
            console.error('[SYNC] Error:', error);
        }
    }
    
    async loadSupplyChainDataFromServer() {
        try {
            const response = await fetch('/api/session/data');
            const result = await response.json();
            
            if (result.success && result.session.supplyChainData) {
                const { stockData, openSupplyData, transitData } = result.session.supplyChainData;
                
                if (stockData && Object.keys(stockData).length > 0) {
                    this.stockData = stockData;
                    this.hasStockData = true;
                }
                
                if (openSupplyData && Object.keys(openSupplyData).length > 0) {
                    this.openSupplyData = openSupplyData;
                    this.hasOpenSupplyData = true;
                }
                
                if (transitData && Object.keys(transitData).length > 0) {
                    this.transitData = transitData;
                    this.hasTransitData = true;
                }
                
                if (this.hasStockData || this.hasOpenSupplyData || this.hasTransitData) {
                    this.render();
                }
            }
        } catch (error) {
            console.error('[SYNC] Load error:', error);
        }
    }
    
    setupSupplyChainSocketListeners() {
        if (typeof io === 'undefined' || !window.socket) return;
        
        window.socket.on('supplyChainDataUpdated', (data) => {
            console.log(`[SOCKET] Updated by ${data.uploadedBy}:`, data.type);
            
            if (data.type === 'stock') {
                this.stockData = data.data;
                this.hasStockData = true;
                this.showNotification(`Stock updated by ${data.uploadedBy}`, 'info');
            } else if (data.type === 'openSupply') {
                this.openSupplyData = data.data;
                this.hasOpenSupplyData = true;
                this.showNotification(`Open Supply updated by ${data.uploadedBy}`, 'info');
            } else if (data.type === 'transit') {
                this.transitData = data.data;
                this.hasTransitData = true;
                this.showNotification(`Transit updated by ${data.uploadedBy}`, 'info');
            }
            
            this.render();
        });
    }
    
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
        notification.className = `notification notification-${type}`;
        
        const icon = type === 'success' ? '✓' : type === 'error' ? '✕' : 'ℹ';
        const bgColor = type === 'success' ? 'bg-green-500' : type === 'error' ? 'bg-red-500' : 'bg-blue-500';
        
        notification.innerHTML = `
            <div class="${bgColor} text-white px-4 py-3 rounded-lg shadow-lg flex items-center gap-3">
                <span class="text-xl font-bold">${icon}</span>
                <span>${message}</span>
            </div>
        `;
        
        container.appendChild(notification);
        setTimeout(() => notification.remove(), 4000);
    }
    
    async loadVariantCycleData(file) {
        try {
            const arrayBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const data = XLSX.utils.sheet_to_json(worksheet);
            
            this.variantCycleDates = {};
            let count = 0;
            
            data.forEach(row => {
                const dfuCode = this.toComparableString(row['DFU'] || row['DFU Code']);
                const partCode = this.toComparableString(row['Part Code'] || row['PartCode'] || row['Product Number']);
                const sos = row['SOS'] || row['Start of Supply'] || '';
                const eos = row['EOS'] || row['End of Supply'] || '';
                const comments = row['Comments'] || row['Comment'] || '';
                
                if (dfuCode && partCode) {
                    if (!this.variantCycleDates[dfuCode]) {
                        this.variantCycleDates[dfuCode] = {};
                    }
                    
                    this.variantCycleDates[dfuCode][partCode] = {
                        sos: sos || 'N/A',
                        eos: eos || 'N/A',
                        comments: comments || ''
                    };
                    count++;
                }
            });
            
            this.hasVariantCycleData = true;
            console.log(`Processed ${count} cycle records`);
            this.showNotification('Cycle data loaded', 'success');
            this.render();
        } catch (error) {
            console.error('Error loading cycle data:', error);
            this.showNotification('Error: ' + error.message, 'error');
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
            this.showNotification('Please select an Excel file', 'error');
            return;
        }
        
        this.loadData(file);
    }
    
    async loadData(file) {
        this.isLoading = true;
        this.render();
        
        try {
            const arrayBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer);
            
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
                this.showNotification(`Loaded ${data.length} records`);
            } else {
                this.showNotification('No data found', 'error');
            }
            
        } catch (error) {
            console.error('Error loading data:', error);
            this.showNotification('Error: ' + error.message, 'error');
        } finally {
            this.isLoading = false;
            this.render();
        }
    }
    
    processMultiVariantDFUs(data) {
        if (data.length === 0) return;
        
        const uniquePlants = new Set();
        const uniqueLines = new Set();
        
        data.forEach(record => {
            const plant = this.toComparableString(record['Production Plant']);
            const line = this.toComparableString(record['Production Line']);
            if (plant) uniquePlants.add(plant);
            if (line) uniqueLines.add(line);
        });
        
        this.availablePlantLocations = Array.from(uniquePlants).sort();
        this.availableProductionLines = Array.from(uniqueLines).sort();
        
        let filteredData = data;
        if (this.selectedPlantLocation) {
            filteredData = filteredData.filter(r => 
                this.toComparableString(r['Production Plant']) === this.selectedPlantLocation
            );
        }
        if (this.selectedProductionLine) {
            filteredData = filteredData.filter(r => 
                this.toComparableString(r['Production Line']) === this.selectedProductionLine
            );
        }
        
        const groupedByDFU = {};
        filteredData.forEach(record => {
            const dfuCode = this.toComparableString(record['DFU']);
            if (dfuCode) {
                if (!groupedByDFU[dfuCode]) groupedByDFU[dfuCode] = [];
                groupedByDFU[dfuCode].push(record);
            }
        });
        
        const allDFUs = {};
        Object.keys(groupedByDFU).forEach(dfuCode => {
            const records = groupedByDFU[dfuCode];
            const uniquePartCodes = [...new Set(records.map(r => this.toComparableString(r['Product Number'])))].filter(Boolean);
            const uniquePlants = [...new Set(records.map(r => this.toComparableString(r['Production Plant'])))].filter(Boolean);
            const uniqueLines = [...new Set(records.map(r => this.toComparableString(r['Production Line'])))].filter(Boolean);
            const isCompleted = this.completedTransfers[dfuCode];
            
            const variantDemand = {};
            uniquePartCodes.forEach(partCode => {
                const partRecords = records.filter(r => this.toComparableString(r['Product Number']) === partCode);
                const totalDemand = partRecords.reduce((sum, r) => sum + (parseFloat(r['weekly fcst']) || 0), 0);
                const partDescription = partRecords[0] ? (partRecords[0]['PartDescription'] || '') : '';
                
                variantDemand[partCode] = {
                    totalDemand,
                    recordCount: partRecords.length,
                    partDescription
                };
            });
            
            allDFUs[dfuCode] = {
                variants: uniquePartCodes,
                recordCount: records.length,
                variantDemand,
                isSingleVariant: uniquePartCodes.length === 1,
                plantLocations: uniquePlants,
                productionLines: uniqueLines,
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
        // Save scroll position before render
        const dfuList = document.querySelector('.bg-gray-50.rounded-lg.p-6 .overflow-y-auto');
        const scrollPos = dfuList ? dfuList.scrollTop : 0;
        
        this.selectedDFU = dfuCode;
        this.render();
        
        // Restore scroll position after render
        setTimeout(() => {
            const newDfuList = document.querySelector('.bg-gray-50.rounded-lg.p-6 .overflow-y-auto');
            if (newDfuList) newDfuList.scrollTop = scrollPos;
            this.ensureGranularContainers();
        }, 10);
    }
    
    selectBulkTarget(dfuCode, targetVariant) {
        this.bulkTransfers[dfuCode] = targetVariant;
        this.render();
    }
    
    setIndividualTransfer(dfuCode, sourceVariant, targetVariant) {
        if (!targetVariant || targetVariant === '' || targetVariant === sourceVariant) {
            if (this.transfers[dfuCode]) {
                delete this.transfers[dfuCode][sourceVariant];
                if (Object.keys(this.transfers[dfuCode]).length === 0) {
                    delete this.transfers[dfuCode];
                }
            }
            
            if (this.granularTransfers[dfuCode]?.[sourceVariant]) {
                delete this.granularTransfers[dfuCode][sourceVariant];
                if (Object.keys(this.granularTransfers[dfuCode]).length === 0) {
                    delete this.granularTransfers[dfuCode];
                }
            }
        } else {
            if (!this.transfers[dfuCode]) this.transfers[dfuCode] = {};
            this.transfers[dfuCode][sourceVariant] = targetVariant;
            
            if (!this.granularTransfers[dfuCode]) this.granularTransfers[dfuCode] = {};
            if (!this.granularTransfers[dfuCode][sourceVariant]) this.granularTransfers[dfuCode][sourceVariant] = {};
            if (!this.granularTransfers[dfuCode][sourceVariant][targetVariant]) {
                this.granularTransfers[dfuCode][sourceVariant][targetVariant] = {};
            }
        }
        
        this.render();
        setTimeout(() => this.ensureGranularContainers(), 100);
    }
    
    toggleGranularWeek(dfuCode, sourceVariant, targetVariant, weekKey) {
        if (!this.granularTransfers[dfuCode]) this.granularTransfers[dfuCode] = {};
        if (!this.granularTransfers[dfuCode][sourceVariant]) this.granularTransfers[dfuCode][sourceVariant] = {};
        if (!this.granularTransfers[dfuCode][sourceVariant][targetVariant]) {
            this.granularTransfers[dfuCode][sourceVariant][targetVariant] = {};
        }
        
        const weekData = this.granularTransfers[dfuCode][sourceVariant][targetVariant][weekKey];
        if (weekData) weekData.selected = !weekData.selected;
        
        this.updateSelectedDemandDisplay(sourceVariant);
    }
    
    updateGranularQuantity(dfuCode, sourceVariant, targetVariant, weekKey, quantity) {
        if (!this.granularTransfers[dfuCode]?.[sourceVariant]?.[targetVariant]?.[weekKey]) return;
        
        const weekData = this.granularTransfers[dfuCode][sourceVariant][targetVariant][weekKey];
        weekData.customQuantity = quantity === '' ? null : parseFloat(quantity);
        
        this.updateSelectedDemandDisplay(sourceVariant);
    }
    
    updateSelectedDemandDisplay(sourceVariant) {
        const el = document.getElementById(`selected-demand-${sourceVariant}`);
        if (el) {
            const selectedDemand = this.calculateSelectedDemand(sourceVariant);
            el.textContent = this.formatNumber(selectedDemand);
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
                if (targetVariant && targetVariant !== variant && targetVariant !== '') {
                    const html = this.renderGranularTransferSection(variant, targetVariant);
                    container.innerHTML = html;
                    this.attachGranularEventListeners();
                } else {
                    container.innerHTML = '';
                }
            }
        });
    }
    
    renderGranularTransferSection(sourceVariant, targetVariant) {
        if (!this.selectedDFU) return '';
        
        const sourceRecords = this.rawData.filter(r => 
            this.toComparableString(r['DFU']) === this.selectedDFU &&
            this.toComparableString(r['Product Number']) === sourceVariant
        );
        
        if (!this.granularTransfers[this.selectedDFU]) this.granularTransfers[this.selectedDFU] = {};
        if (!this.granularTransfers[this.selectedDFU][sourceVariant]) {
            this.granularTransfers[this.selectedDFU][sourceVariant] = {};
        }
        if (!this.granularTransfers[this.selectedDFU][sourceVariant][targetVariant]) {
            this.granularTransfers[this.selectedDFU][sourceVariant][targetVariant] = {};
        }
        
        const granularData = this.granularTransfers[this.selectedDFU][sourceVariant][targetVariant];
        
        sourceRecords.forEach(record => {
            const weekNumber = this.toComparableString(record['Week Number']);
            const sourceLocation = this.toComparableString(record['Source Location']);
            const calendarWeek = record['Calendar.week'];
            const year = calendarWeek ? new Date(calendarWeek).getFullYear() : 2025;
            const paddedWeek = String(weekNumber).padStart(2, '0');
            const weekKey = `${year}-${paddedWeek}-${sourceLocation}`;
            const demand = parseFloat(record['weekly fcst']) || 0;
            
            if (!granularData[weekKey]) {
                granularData[weekKey] = {
                    selected: false,
                    originalQuantity: demand,
                    customQuantity: null,
                    calendarWeek: record['Calendar.week'],
                    weekNumber: weekNumber,
                    sourceLocation: sourceLocation
                };
            }
        });
        
        const sortedWeeks = Object.entries(granularData).sort((a, b) => {
            const aDate = new Date(a[1].calendarWeek || '2025-01-01');
            const bDate = new Date(b[1].calendarWeek || '2025-01-01');
            const aYear = aDate.getFullYear();
            const bYear = bDate.getFullYear();
            
            if (aYear !== bYear) return aYear - bYear;
            
            return parseInt(a[1].weekNumber) - parseInt(b[1].weekNumber);
        });
        
        return `
            <div class="mt-4 p-3 bg-blue-50 rounded-lg border border-blue-200">
                <h5 class="font-semibold text-blue-800 mb-3">Granular Transfer (Select Specific Weeks)</h5>
                <div class="space-y-2 max-h-64 overflow-y-auto">
                    ${sortedWeeks.map(([weekKey, weekData]) => `
                        <div class="flex items-center gap-3 p-2 bg-white rounded border ${weekData.selected ? 'border-blue-500' : 'border-gray-200'}">
                            <input 
                                type="checkbox" 
                                ${weekData.selected ? 'checked' : ''}
                                class="w-4 h-4"
                                data-granular-toggle
                                data-dfu="${this.selectedDFU}"
                                data-source="${sourceVariant}"
                                data-target="${targetVariant}"
                                data-week="${weekKey}"
                            >
                            <div class="flex-1 grid grid-cols-3 gap-2 text-sm">
                                <div><span class="text-gray-600">Week:</span> <span class="font-medium">${weekData.weekNumber}</span></div>
                                <div><span class="text-gray-600">Location:</span> <span class="font-medium">${weekData.sourceLocation}</span></div>
                                <div><span class="text-gray-600">Demand:</span> <span class="font-medium">${this.formatNumber(weekData.originalQuantity)}</span></div>
                            </div>
                            <div class="flex items-center gap-2">
                                <span class="text-sm text-gray-600">Transfer:</span>
                                <input 
                                    type="number" 
                                    class="w-24 px-2 py-1 border rounded text-sm"
                                    placeholder="${this.formatNumber(weekData.originalQuantity)}"
                                    value="${weekData.customQuantity !== null && weekData.customQuantity !== '' ? weekData.customQuantity : ''}"
                                    ${!weekData.selected ? 'disabled' : ''}
                                    data-granular-qty
                                    data-dfu="${this.selectedDFU}"
                                    data-source="${sourceVariant}"
                                    data-target="${targetVariant}"
                                    data-week="${weekKey}"
                                >
                            </div>
                        </div>
                    `).join('')}
                </div>
            </div>
        `;
    }
    
    async executeTransfer() {
        if (!this.selectedDFU) {
            this.showNotification('Select a DFU first', 'error');
            return;
        }
        
        const timestamp = new Date().toISOString().split('T')[0] + ' ' + new Date().toTimeString().split(' ')[0];
        const dfuStr = this.selectedDFU;
        const dfuRecords = this.rawData.filter(r => this.toComparableString(r['DFU']) === dfuStr);
        
        let executionType = '';
        let executionMessage = '';
        let transfersExecuted = false;
        
        // BULK TRANSFER - Process independently
        if (this.bulkTransfers[dfuStr]) {
            console.log('[EXECUTE] Bulk transfer');
            const targetVariant = this.bulkTransfers[dfuStr];
            const originalVariants = [...new Set(dfuRecords.map(r => this.toComparableString(r['Product Number'])))];
            const variantsToTransfer = originalVariants.filter(v => v !== targetVariant);
            
            variantsToTransfer.forEach(sourceVariant => {
                const sourceRecords = dfuRecords.filter(r => 
                    this.toComparableString(r['Product Number']) === sourceVariant
                );
                
                sourceRecords.forEach(record => {
                    const transferDemand = parseFloat(record['weekly fcst']) || 0;
                    const targetRecord = dfuRecords.find(r => 
                        this.toComparableString(r['Product Number']) === targetVariant && 
                        this.toComparableString(r['Calendar.week']) === this.toComparableString(record['Calendar.week']) &&
                        this.toComparableString(r['Source Location']) === this.toComparableString(record['Source Location'])
                    );
                    
                    if (targetRecord) {
                        const oldDemand = parseFloat(targetRecord['weekly fcst']) || 0;
                        targetRecord['weekly fcst'] = oldDemand + transferDemand;
                        const existingHistory = targetRecord['Transfer History'] || '';
                        const newEntry = `[${sourceVariant} → ${transferDemand} @ ${timestamp}]`;
                        const pipoPrefix = existingHistory.startsWith('PIPO') ? '' : 'PIPO ';
                        targetRecord['Transfer History'] = existingHistory ? 
                            `${existingHistory} ${newEntry}` : `${pipoPrefix}${newEntry}`;
                        
                        record['weekly fcst'] = 0;
                        const sourceHistory = record['Transfer History'] || '';
                        const sourceEntry = `[→ ${targetVariant}: -${transferDemand} @ ${timestamp}]`;
                        record['Transfer History'] = sourceHistory ? 
                            `${sourceHistory} ${sourceEntry}` : `PIPO ${sourceEntry}`;
                    } else {
                        const newRecord = { ...record };
                        newRecord['Product Number'] = isNaN(targetVariant) ? targetVariant : Number(targetVariant);
                        newRecord['Transfer History'] = `PIPO [from ${sourceVariant}: ${transferDemand} @ ${timestamp}]`;
                        this.rawData.push(newRecord);
                        
                        record['weekly fcst'] = 0;
                        record['Transfer History'] = `PIPO [→ ${targetVariant}: -${transferDemand} @ ${timestamp}]`;
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
            executionMessage = `All variants → ${targetVariant}`;
            this.showNotification(executionMessage, 'success');
            transfersExecuted = true;
        }
        
        // GRANULAR TRANSFERS - Process independently
        if (this.granularTransfers[dfuStr] && Object.keys(this.granularTransfers[dfuStr]).length > 0) {
            console.log('[EXECUTE] Granular transfers');
            const granularTransfers = this.granularTransfers[dfuStr];
            let count = 0;
            
            Object.keys(granularTransfers).forEach(sourceVariant => {
                const sourceTargets = granularTransfers[sourceVariant];
                
                Object.keys(sourceTargets).forEach(targetVariant => {
                    const weekTransfers = sourceTargets[targetVariant];
                    
                    Object.keys(weekTransfers).forEach(weekKey => {
                        const weekTransfer = weekTransfers[weekKey];
                        if (!weekTransfer.selected) return;
                        
                        const [year, weekNum, sourceLocation] = weekKey.split('-');
                        const sourceRecord = dfuRecords.find(r => 
                            this.toComparableString(r['Product Number']) === sourceVariant &&
                            this.toComparableString(r['Week Number']) === weekNum &&
                            this.toComparableString(r['Source Location']) === sourceLocation
                        );
                        
                        if (sourceRecord) {
                            const originalDemand = parseFloat(sourceRecord['weekly fcst']) || 0;
                            const transferAmount = weekTransfer.customQuantity !== null ?
                                parseFloat(weekTransfer.customQuantity) : originalDemand;
                            
                            const targetRecord = dfuRecords.find(r => 
                                this.toComparableString(r['Product Number']) === targetVariant &&
                                this.toComparableString(r['Week Number']) === weekNum &&
                                this.toComparableString(r['Source Location']) === sourceLocation
                            );
                            
                            if (targetRecord) {
                                const oldDemand = parseFloat(targetRecord['weekly fcst']) || 0;
                                targetRecord['weekly fcst'] = oldDemand + transferAmount;
                                const existingHistory = targetRecord['Transfer History'] || '';
                                const newEntry = `[${sourceVariant} → ${transferAmount} @ ${timestamp}]`;
                                const pipoPrefix = existingHistory.startsWith('PIPO') ? '' : 'PIPO ';
                                targetRecord['Transfer History'] = existingHistory ? 
                                    `${existingHistory} ${newEntry}` : `${pipoPrefix}${newEntry}`;
                            } else {
                                const newRecord = { ...sourceRecord };
                                newRecord['Product Number'] = isNaN(targetVariant) ? targetVariant : Number(targetVariant);
                                newRecord['weekly fcst'] = transferAmount;
                                newRecord['Transfer History'] = `PIPO [from ${sourceVariant}: ${transferAmount} @ ${timestamp}]`;
                                this.rawData.push(newRecord);
                            }
                            
                            sourceRecord['weekly fcst'] = originalDemand - transferAmount;
                            if (!sourceRecord['Transfer History']) sourceRecord['Transfer History'] = '';
                            const sourceEntry = `[→ ${targetVariant}: -${transferAmount} @ ${timestamp}]`;
                            sourceRecord['Transfer History'] = sourceRecord['Transfer History'] ? 
                                `${sourceRecord['Transfer History']} PIPO ${sourceEntry}` : `PIPO ${sourceEntry}`;
                            
                            count++;
                        }
                    });
                });
            });
            
            if (count > 0) {
                this.completedTransfers[dfuStr] = {
                    type: 'granular',
                    timestamp: timestamp,
                    transferCount: count
                };
                
                executionType = 'Granular Transfer';
                executionMessage = `${count} week transfers`;
                this.showNotification(executionMessage, 'success');
                delete this.granularTransfers[dfuStr];
                transfersExecuted = true;
            }
        }
        
        // INDIVIDUAL TRANSFERS - Only if no granular for that variant
        if (!transfersExecuted && this.transfers[dfuStr] && Object.keys(this.transfers[dfuStr]).length > 0) {
            console.log('[EXECUTE] Individual transfers');
            const individualTransfers = this.transfers[dfuStr];
            const granularTransfers = this.granularTransfers[dfuStr] || {};
            let count = 0;
            
            Object.keys(individualTransfers).forEach(sourceVariant => {
                const targetVariant = individualTransfers[sourceVariant];
                
                // Skip if granular configured
                // Skip ONLY if granular weeks are actually SELECTED
                let hasSelectedGranularWeeks = false;
                if (granularTransfers[sourceVariant]) {
                    Object.keys(granularTransfers[sourceVariant]).forEach(tv => {
                        const weeks = granularTransfers[sourceVariant][tv];
                        if (Object.values(weeks).some(w => w.selected)) {
                            hasSelectedGranularWeeks = true;
                        }
                    });
                }

                if (hasSelectedGranularWeeks) {
                    console.log(`Skip ${sourceVariant} - has selected granular weeks`);
                    return;
            }
                
                if (sourceVariant !== targetVariant) {
                    const sourceRecords = dfuRecords.filter(r => 
                        this.toComparableString(r['Product Number']) === sourceVariant
                    );
                    
                    sourceRecords.forEach(record => {
                        const transferDemand = parseFloat(record['weekly fcst']) || 0;
                        const targetRecord = dfuRecords.find(r => 
                            this.toComparableString(r['Product Number']) === targetVariant && 
                            this.toComparableString(r['Calendar.week']) === this.toComparableString(record['Calendar.week']) &&
                            this.toComparableString(r['Source Location']) === this.toComparableString(record['Source Location'])
                        );
                        
                        if (targetRecord) {
                            const oldDemand = parseFloat(targetRecord['weekly fcst']) || 0;
                            targetRecord['weekly fcst'] = oldDemand + transferDemand;
                            const existingHistory = targetRecord['Transfer History'] || '';
                            const newEntry = `[${sourceVariant} → ${transferDemand} @ ${timestamp}]`;
                            const pipoPrefix = existingHistory.startsWith('PIPO') ? '' : 'PIPO ';
                            targetRecord['Transfer History'] = existingHistory ? 
                                `${existingHistory} ${newEntry}` : `${pipoPrefix}${newEntry}`;
                            
                            record['weekly fcst'] = 0;
                            const sourceHistory = record['Transfer History'] || '';
                            const sourceEntry = `[→ ${targetVariant}: -${transferDemand} @ ${timestamp}]`;
                            record['Transfer History'] = sourceHistory ? 
                                `${sourceHistory} ${sourceEntry}` : `PIPO ${sourceEntry}`;
                        } else {
                            const newRecord = { ...record };
                            newRecord['Product Number'] = isNaN(targetVariant) ? targetVariant : Number(targetVariant);
                            newRecord['Transfer History'] = `PIPO [from ${sourceVariant}: ${transferDemand} @ ${timestamp}]`;
                            this.rawData.push(newRecord);
                            
                            record['weekly fcst'] = 0;
                            record['Transfer History'] = `PIPO [→ ${targetVariant}: -${transferDemand} @ ${timestamp}]`;
                        }
                    });
                    
                    count++;
                }
            });
            
            if (count > 0) {
                this.transfers[dfuStr] = {};
                this.completedTransfers[dfuStr] = {
                    type: 'individual',
                    transfers: individualTransfers,
                    timestamp: timestamp,
                    transferCount: count
                };
                
                executionType = 'Individual Transfers';
                executionMessage = `${count} variant transfers`;
                this.showNotification(executionMessage);
                transfersExecuted = true;
            }
        }
        
        if (transfersExecuted) {
            this.lastExecutionSummary[dfuStr] = {
                type: executionType,
                message: executionMessage,
                timestamp: timestamp
            };
            
            this.processMultiVariantDFUs(this.rawData);
            this.render();
        } else {
            this.showNotification('No transfers configured', 'error');
        }
    }
    
    async undoTransfer(dfuCode) {
        const dfuStr = this.toComparableString(dfuCode);
        
        if (!this.originalRawData || this.originalRawData.length === 0) {
            this.showNotification('No original data', 'error');
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
        
        this.showNotification(`Transfer undone for ${dfuStr}`, 'success');
        
        const current = this.selectedDFU;
        this.selectedDFU = null;
        setTimeout(() => {
            this.selectedDFU = current;
            this.render();
        }, 50);
    }
    
    async addManualVariant(dfuCode) {
        const variantCode = prompt('Enter new variant code:');
        if (!variantCode) return;
        
        const dfuRecords = this.rawData.filter(r => 
            this.toComparableString(r['DFU']) === dfuCode
        );
        
        if (dfuRecords.length === 0) {
            this.showNotification('No records found', 'error');
            return;
        }
        
        const template = dfuRecords[0];
        const uniqueWeekLocations = new Set();
        dfuRecords.forEach(r => {
            const key = `${r['Week Number']}-${r['Source Location']}-${r['Calendar.week']}`;
            uniqueWeekLocations.add(key);
        });
        
        const newRecords = [];
        uniqueWeekLocations.forEach(key => {
            const [weekNum, sourceLoc, calWeek] = key.split('-');
            const newRecord = { ...template };
            newRecord['Product Number'] = variantCode;
            newRecord['weekly fcst'] = 0;
            newRecord['Week Number'] = weekNum;
            newRecord['Source Location'] = sourceLoc;
            newRecord['Calendar.week'] = calWeek;
            newRecord['Transfer History'] = 'PIPO [Manually added]';
            newRecords.push(newRecord);
        });
        
        this.rawData.push(...newRecords);
        this.processMultiVariantDFUs(this.rawData);
        this.showNotification(`Added ${variantCode} to ${dfuCode}`, 'success');
        this.render();
    }
    
    async exportData() {
        try {
            const wb = XLSX.utils.book_new();
            
            const formattedData = this.rawData.map((record, index) => {
                const formatted = {
                    OrderNumber: index + 1,  // Add as first column
                    ...record
                };
                
                if (formatted['Week Number']) {
                    const weekNumStr = String(formatted['Week Number']).trim();
                    const parts = weekNumStr.split('_');
                    
                    if (parts.length === 2) {
                        const weekNum = parseInt(parts[0]);
                        const year = parseInt(parts[1]);
                        
                        if (!isNaN(weekNum) && !isNaN(year) && weekNum >= 1 && weekNum <= 53) {
                            const targetDate = this.getDateFromWeekNumber(year, weekNum);
                            formatted['Calendar.week'] = targetDate.toISOString().split('T')[0];
                        }
                    }
                }
                
                return formatted;
            });
            // Add this line to check
            console.log('First row with OrderNumber:', formattedData[0]);
            
            const ws = XLSX.utils.json_to_sheet(formattedData);
            XLSX.utils.book_append_sheet(wb, ws, 'Updated Demand');
            XLSX.writeFile(wb, 'DFU_Transfer_Updated.xlsx');
            
            this.showNotification('Data exported');
        } catch (error) {
            console.error('Export error:', error);
            this.showNotification('Export failed: ' + error.message, 'error');
        }
    }
    
    render() {
        const app = document.getElementById('app');
        
        if (!this.isProcessed) {
            app.innerHTML = this.renderUploadScreen();
            this.attachUploadListeners();
            return;
        }
        
        const totalDFUs = Object.keys(this.filteredDFUs).length;
        const multiVariantCount = Object.keys(this.filteredDFUs).filter(dfu => !this.filteredDFUs[dfu].isSingleVariant).length;
        
        app.innerHTML = `
            <div>
                <h1 class="text-3xl font-bold text-gray-800 mb-2">DFU Demand Transfer Management</h1>
                <p class="text-gray-600">Managing ${totalDFUs} DFUs (${multiVariantCount} multi-variant, ${totalDFUs - multiVariantCount} single)</p>
            </div>

            ${this.renderControls()}
            
            <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                ${this.renderDFUList()}
                ${this.renderDetailsPanel()}
            </div>
        `;
        
        this.attachEventListeners();
        this.ensureGranularContainers();
    }
    
    renderUploadScreen() {
        return `
            <div class="max-w-6xl mx-auto p-6 bg-white min-h-screen">
                <div class="text-center py-12">
                    <div class="bg-blue-50 rounded-lg p-8 inline-block">
                        <div class="w-12 h-12 mb-4 mx-auto bg-blue-600 rounded-full flex items-center justify-center">
                            <svg class="w-6 h-6 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                            </svg>
                        </div>
                        <h2 class="text-xl font-semibold mb-2">Upload Demand Data</h2>
                        <p class="text-gray-600 mb-4">Upload Excel file with "Total Demand" format</p>
                        
                        ${this.isLoading ? `
                            <div class="text-blue-600">
                                <div class="loading-spinner mb-2"></div>
                                <p>Processing...</p>
                            </div>
                        ` : `
                            <div class="space-y-4">
                                <div>
                                    <input type="file" accept=".xlsx,.xls" class="file-input" id="fileInput">
                                    <p class="text-sm text-gray-500 mt-2">Formats: .xlsx, .xls</p>
                                    <div class="mt-3 p-3 bg-amber-50 border border-amber-300 rounded-lg">
                                        <p class="text-sm text-amber-800">
                                            <strong>⚠️ Important:</strong> Upload all supply chain files (Stock, Open Supply, Transit) and Cycle Dates BEFORE uploading Total Demand. The system will process the data immediately.
                                        </p>
                                    </div>
                                </div>
                                
                                <div class="border-t pt-4 mt-4 space-y-3">
                                    <h3 class="text-sm font-medium text-gray-700 mb-2">Optional: Supply Chain Data</h3>
                                    
                                    <label class="block">
                                        <span class="text-sm text-gray-600">Upload Stock RRP4 (SOH):</span>
                                        <input type="file" accept=".xlsx,.xls" class="file-input text-sm" id="stockFileInput">
                                        ${this.hasStockData ? '<span class="text-xs text-green-600">✓ Loaded</span>' : ''}
                                    </label>
                                    
                                    <label class="block">
                                        <span class="text-sm text-gray-600">Production RRP4 (Open Supply):</span>
                                        <input type="file" accept=".xlsx,.xls" class="file-input text-sm" id="openSupplyFileInput">
                                        ${this.hasOpenSupplyData ? '<span class="text-xs text-green-600">✓ Loaded</span>' : ''}
                                    </label>
                                    
                                    <label class="block">
                                        <span class="text-sm text-gray-600">Transport receipts Confirmed RRP4 only orders (In Transit):</span>
                                        <input type="file" accept=".xlsx,.xls" class="file-input text-sm" id="transitFileInput">
                                        ${this.hasTransitData ? '<span class="text-xs text-green-600">✓ Loaded</span>' : ''}
                                    </label>
                                </div>
                                
                                <div class="border-t pt-4">
                                    <h3 class="text-sm font-medium text-gray-700 mb-2">Optional: Variant Cycle Dates</h3>
                                    <input type="file" accept=".xlsx,.xls" class="file-input" id="cycleFileInput">
                                    ${this.hasVariantCycleData ? '<span class="text-xs text-green-600 ml-2">✓ Loaded</span>' : ''}
                                    <p class="text-xs text-gray-500 mt-1">DFU, Part Code, SOS, EOS columns</p>
                                </div>
                            </div>
                        `}
                    </div>
                </div>
            </div>
        `;
    }
    
    renderControls() {
        return `
            <div class="flex gap-4 mb-6 flex-wrap">
                <div class="relative flex-1 min-w-[200px]">
                    <svg class="absolute left-3 top-3 h-4 w-4 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
                    </svg>
                    <input type="text" placeholder="Search..." value="${this.searchTerm}" class="search-input" id="searchInput">
                </div>
                <select class="px-4 py-2 border rounded-lg" id="plantLocationFilter">
                    <option value="">All Plants</option>
                    ${this.availablePlantLocations.map(loc => `
                        <option value="${loc}" ${this.selectedPlantLocation === loc ? 'selected' : ''}>Plant ${loc}</option>
                    `).join('')}
                </select>
                <select class="px-4 py-2 border rounded-lg" id="productionLineFilter">
                    <option value="">All Lines</option>
                    ${this.availableProductionLines.map(line => `
                        <option value="${line}" ${this.selectedProductionLine === line ? 'selected' : ''}>Line ${line}</option>
                    `).join('')}
                </select>
                ${this.hasVariantCycleData ? `
                    <span class="inline-flex items-center px-3 py-2 text-sm text-green-700 bg-green-100 rounded-lg">
                        ✓ Cycle Data
                    </span>
                ` : `
                    <label class="btn btn-secondary cursor-pointer">
                        Load Cycle Dates
                        <input type="file" accept=".xlsx,.xls" class="hidden" id="cycleFileInput">
                    </label>
                `}
                <button class="btn btn-success" id="exportBtn">Export</button>
            </div>
        `;
    }
    
    renderDFUList() {
        return `
            <div class="bg-gray-50 rounded-lg p-6">
                <h3 class="font-semibold text-gray-800 mb-4">All DFUs (${Object.keys(this.filteredDFUs).length})</h3>
                <div class="relative" style="height: 580px;">
                    <div class="space-y-2 h-full overflow-y-auto pr-1">
                        ${Object.keys(this.filteredDFUs).map(dfuCode => {
                            const dfuData = this.filteredDFUs[dfuCode];
                            if (!dfuData) return '';
                            
                            const hasTransfers = (this.transfers[dfuCode] && Object.values(this.transfers[dfuCode]).some(t => t && t !== '')) ||
                                this.bulkTransfers[dfuCode] ||
                                (this.granularTransfers[dfuCode] && Object.keys(this.granularTransfers[dfuCode]).some(sv => {
                                    const targets = this.granularTransfers[dfuCode][sv];
                                    return Object.keys(targets).some(tv => {
                                        return Object.values(targets[tv]).some(w => w.selected);
                                    });
                                }));
                            
                            return `
                                <div class="dfu-card ${this.selectedDFU === dfuCode ? 'selected' : ''}" data-dfu="${dfuCode}">
                                    <div class="flex justify-between items-start">
                                        <div>
                                            <h4 class="font-medium text-gray-800">DFU: ${dfuCode}</h4>
                                            <p class="text-sm text-gray-600">
                                                ${dfuData.variants.length} variant${dfuData.variants.length > 1 ? 's' : ''}
                                                ${dfuData.isCompleted ? ' (done)' : ''}
                                            </p>
                                        </div>
                                        <div class="text-right">
                                            ${dfuData.isCompleted ? `
                                                <span class="text-green-600 text-sm">✓ Done</span>
                                            ` : hasTransfers ? `
                                                <span class="text-green-600 text-sm">✓ Ready</span>
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
        `;
    }
    
    renderDetailsPanel() {
        return `
            <div class="bg-white border border-gray-200 rounded-lg p-6">
                ${this.selectedDFU && this.multiVariantDFUs[this.selectedDFU] ? 
                    this.renderDetailsSection() : 
                    '<div class="text-center py-12 text-gray-400"><p>Select a DFU</p></div>'
                }
            </div>
        `;
    }
    
    renderDetailsSection() {
        if (!this.selectedDFU) return '';
        
        const dfu = this.multiVariantDFUs[this.selectedDFU];
        if (!dfu) return '';
        
        return `
            <div>
                <div class="flex justify-between items-center mb-4">
                    <h3 class="font-semibold text-gray-800">
                        DFU: ${this.selectedDFU}
                        ${dfu.isCompleted ? '<span class="ml-2 px-2 py-1 text-xs bg-green-100 text-green-800 rounded-full">✓</span>' : ''}
                    </h3>
                    ${!dfu.isCompleted ? `
                        <button class="btn btn-primary text-sm" id="addVariantBtn">+ Add Variant</button>
                    ` : ''}
                </div>
                
                ${dfu.isCompleted ? this.renderCompletedSection() : this.renderTransferSection()}
                
                <div class="text-center text-sm text-gray-500 mt-6">
                    <p><strong>How to Use:</strong></p>
                    <ul class="text-left list-disc list-inside mt-2 space-y-1">
                        <li><strong>Bulk:</strong> Click variant to transfer all</li>
                        <li><strong>Individual:</strong> Use dropdowns for specific transfers</li>
                        <li><strong>Granular:</strong> Select specific weeks after setting target</li>
                        <li><strong>Execute:</strong> Click button to apply transfers</li>
                    </ul>
                </div>
            </div>
        `;
    }
    
    renderCompletedSection() {
        const dfu = this.multiVariantDFUs[this.selectedDFU];
        const info = dfu.completionInfo;
        
        return `
            <div class="mb-6 p-4 bg-green-50 rounded-lg border border-green-200">
                <div class="flex justify-between items-start">
                    <div class="flex-1">
                        <h4 class="font-semibold text-green-800 mb-3">✓ Transfer Complete</h4>
                        <div class="text-sm text-green-700">
                            <p><strong>Type:</strong> ${info.type === 'bulk' ? 'Bulk' : info.type === 'granular' ? 'Granular' : 'Individual'}</p>
                            <p><strong>Date:</strong> ${info.timestamp}</p>
                            ${info.type === 'bulk' ? `
                                <p><strong>Target:</strong> ${info.targetVariant}</p>
                            ` : `
                                <p><strong>Transfers:</strong> ${info.transferCount}</p>
                            `}
                        </div>
                    </div>
                    <button class="btn btn-secondary text-sm" onclick="appInstance.undoTransfer('${this.selectedDFU}')">Undo</button>
                </div>
            </div>
            
            <div class="space-y-3">
                ${dfu.variants.map(variant => {
                    const demandData = dfu.variantDemand[variant];
                    const supplyChain = this.getSupplyChainData(variant);
                    
                    return `
                        <div class="border rounded-lg p-3 bg-white">
                            <div class="flex justify-between items-center">
                                <div class="flex-1">
                                    <h5 class="font-medium text-gray-800">Part: ${variant}</h5>
                                    <p class="text-xs text-gray-500">${demandData?.partDescription || ''}</p>
                                    ${supplyChain.hasData ? `
                                        <div class="mt-2 text-xs">
                                            <p class="text-blue-600"><strong>SOH:</strong> ${this.formatNumber(supplyChain.soh)}</p>
                                            <p class="text-green-600"><strong>Open Supply:</strong> ${this.formatNumber(supplyChain.openSupply)}</p>
                                            <p class="text-purple-600"><strong>Transit:</strong> ${this.formatNumber(supplyChain.transit)}</p>
                                            <p class="text-gray-800 font-semibold"><strong>Total:</strong> ${this.formatNumber(supplyChain.total)}</p>
                                        </div>
                                    ` : ''}
                                </div>
                                <div class="text-right">
                                    <p class="font-medium text-gray-800">${this.formatNumber(demandData?.totalDemand || 0)}</p>
                                    <p class="text-sm text-gray-600">demand</p>
                                </div>
                            </div>
                        </div>
                    `;
                }).join('')}
            </div>
        `;
    }
    
    renderTransferSection() {
        const dfu = this.multiVariantDFUs[this.selectedDFU];
        
        return `
            <div class="mb-6 p-4 bg-purple-50 rounded-lg border">
                <h4 class="font-semibold text-purple-800 mb-3">Bulk Transfer (All → One)</h4>
                <p class="text-sm text-purple-600 mb-3">Transfer all variants to one target:</p>
                <div class="flex flex-wrap gap-2">
                    ${dfu.variants.map(variant => {
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
                            <strong>Selected:</strong> All → <strong>${this.bulkTransfers[this.selectedDFU]}</strong>
                        </p>
                    </div>
                ` : ''}
            </div>
            
            ${this.renderIndividualSection()}
            ${this.renderActionButtons()}
        `;
    }
    
    renderIndividualSection() {
        const dfu = this.multiVariantDFUs[this.selectedDFU];
        
        return `
            <div class="mb-6">
                <h4 class="font-semibold text-gray-800 mb-3">Individual Transfers</h4>
                <div class="space-y-4">
                    ${dfu.variants.map(variant => {
                        const demandData = dfu.variantDemand[variant];
                        const currentTransfer = this.transfers[this.selectedDFU]?.[variant];
                        const supplyChain = this.getSupplyChainData(variant);
                        const networkCalc = this.calculateNetworkStockPlusDemand(variant);
                        const selectedDemand = this.calculateSelectedDemand(variant);
                        
                        return `
                            <div class="border rounded-lg p-4 bg-gray-50">
                                <div class="flex justify-between items-center mb-3">
                                    <div class="flex-1">
                                        <h5 class="font-medium text-gray-800">Part: ${variant}</h5>
                                        <p class="text-xs text-gray-500 mb-1">${demandData?.partDescription || ''}</p>
                                        <p class="text-sm text-gray-600">${demandData?.recordCount || 0} records • ${this.formatNumber(demandData?.totalDemand || 0)} demand</p>
                                        ${supplyChain.hasData ? `
                                            <div class="mt-2 text-xs space-y-1">
                                                <p class="text-blue-600"><strong>SOH:</strong> ${this.formatNumber(supplyChain.soh)}</p>
                                                <p class="text-green-600"><strong>Open Supply:</strong> ${this.formatNumber(supplyChain.openSupply)}</p>
                                                <p class="text-purple-600"><strong>Transit:</strong> ${this.formatNumber(supplyChain.transit)}</p>
                                                <p class="text-gray-800 font-semibold"><strong>Total:</strong> ${this.formatNumber(supplyChain.total)}</p>
                                            </div>
                                        ` : ''}
                                        ${(() => {
                                            const cycleData = this.getCycleDataForVariant(this.selectedDFU, variant);
                                            if (cycleData) {
                                                return `
                                                    <div class="mt-1 text-xs space-y-0.5">
                                                        <p class="text-blue-600"><strong>SOS:</strong> ${cycleData.sos}</p>
                                                        <p class="text-red-600"><strong>EOS:</strong> ${cycleData.eos}</p>
                                                        ${cycleData.comments ? `<p class="text-gray-600 italic">${cycleData.comments}</p>` : ''}
                                                    </div>
                                                `;
                                            }
                                            return '';
                                        })()}
                                    </div>
                                </div>
                                
                                ${currentTransfer && currentTransfer !== variant && currentTransfer !== '' ? `
                                    <div class="grid grid-cols-2 gap-3 mb-3">
                                        <div class="p-3 bg-orange-50 border border-orange-300 rounded-lg">
                                            <p class="text-xs text-orange-700 mb-1 font-semibold">Network Stock + Demand</p>
                                            <p class="text-2xl font-bold ${networkCalc && networkCalc.hasSupplyData ? (networkCalc.result < 0 ? 'text-red-600' : 'text-green-600') : 'text-orange-900'}" id="network-calc-${variant}">
                                                ${networkCalc && networkCalc.hasSupplyData ? this.formatNumber(networkCalc.result) : '-'}
                                            </p>
                                            ${networkCalc && networkCalc.hasSupplyData ? `
                                                <p class="text-xs text-orange-600 mt-1">
                                                    ${this.formatNumber(networkCalc.networkStock)} + (${this.formatNumber(networkCalc.totalDemand)})
                                                </p>
                                            ` : '<p class="text-xs text-orange-600 mt-1">No data</p>'}
                                        </div>
                                        
                                        <div class="p-3 bg-teal-50 border border-teal-300 rounded-lg">
                                            <p class="text-xs text-teal-700 mb-1 font-semibold">Selected Demand</p>
                                            <p class="text-2xl font-bold text-teal-900" id="selected-demand-${variant}">
                                                ${this.formatNumber(selectedDemand)}
                                            </p>
                                            <p class="text-xs text-teal-600 mt-1">Live update</p>
                                        </div>
                                    </div>
                                ` : ''}
                                
                                <div class="flex items-center gap-2 text-sm mb-3">
                                    <span class="text-gray-600">Transfer to:</span>
                                    <select class="px-2 py-1 border rounded text-sm" data-source-variant="${variant}" id="select-${variant}">
                                        <option value="">Select...</option>
                                        ${dfu.variants.map(targetVariant => `
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
        
        let hasGranular = false;
        if (this.granularTransfers[this.selectedDFU]) {
            for (const sv of Object.keys(this.granularTransfers[this.selectedDFU])) {
                const targets = this.granularTransfers[this.selectedDFU][sv];
                for (const tv of Object.keys(targets)) {
                    const weeks = targets[tv];
                    if (Object.values(weeks).some(w => w.selected)) {
                        hasGranular = true;
                        break;
                    }
                }
                if (hasGranular) break;
            }
        }
        
        const hasTransfers = hasValidTransfers || this.bulkTransfers[this.selectedDFU] || hasGranular;
        
        if (hasTransfers) {
            return `
                <div class="p-3 bg-blue-50 rounded-lg">
                    <div class="text-sm text-blue-800 mb-3">
                        ${this.bulkTransfers[this.selectedDFU] ? 
                            `<p>✓ Bulk transfer ready</p>` : 
                            hasGranular ? `<p>✓ Granular transfers configured</p>` :
                            `<p>✓ Individual transfers configured</p>`
                        }
                    </div>
                    <button class="btn btn-primary w-full" id="executeBtn">Execute Transfer</button>
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
        
        const plantFilter = document.getElementById('plantLocationFilter');
        if (plantFilter) {
            plantFilter.addEventListener('change', (e) => {
                this.selectedPlantLocation = e.target.value;
                this.processMultiVariantDFUs(this.rawData);
                this.render();
            });
        }
        
        const lineFilter = document.getElementById('productionLineFilter');
        if (lineFilter) {
            lineFilter.addEventListener('change', (e) => {
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
                this.setIndividualTransfer(this.selectedDFU, sourceVariant, targetVariant);
            });
        });
    }
    
    attachUploadListeners() {
        const fileInput = document.getElementById('fileInput');
        if (fileInput) {
            fileInput.addEventListener('change', (e) => this.handleFileUpload(e));
        }
        
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

document.addEventListener('DOMContentLoaded', () => {
    if (!window.preventAutoInit) {
        window.appInstance = new DemandTransferApp();
    }
});