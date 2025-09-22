class MLUCADashboard {
    constructor() {
        this.data = [];
        this.filteredData = [];
        this.performanceChart = null;
        this.fundamentalsChart = null;

        this.initializeEventListeners();
        this.tryAutoLoadFile();
    }

    initializeEventListeners() {
        // Upload controls
        const uploadBtn = document.getElementById('uploadBtn');
        const fileInput = document.getElementById('fileInput');
        const selectFileBtn = document.getElementById('selectFileBtn');
        const uploadArea = document.getElementById('uploadArea');

        if (uploadBtn && fileInput) {
            uploadBtn.addEventListener('click', (e) => {
                e.preventDefault();
                fileInput.click();
            });
        }

        if (selectFileBtn && fileInput) {
            selectFileBtn.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                fileInput.click();
            });
        }

        if (fileInput) {
            fileInput.addEventListener('change', this.handleFileSelect.bind(this));
        }

        // Drag and drop
        if (uploadArea && fileInput) {
            uploadArea.addEventListener('dragover', this.handleDragOver.bind(this));
            uploadArea.addEventListener('dragleave', this.handleDragLeave.bind(this));
            uploadArea.addEventListener('drop', this.handleDrop.bind(this));
            uploadArea.addEventListener('click', (e) => {
                e.preventDefault();
                fileInput.click();
            });
        }

        // Other controls
        const refreshBtn = document.getElementById('refreshBtn');
        const periodSelect = document.getElementById('periodSelect');
        const exportPerformance = document.getElementById('exportPerformance');
        const exportFundamentals = document.getElementById('exportFundamentals');

        if (refreshBtn) {
            refreshBtn.addEventListener('click', this.refreshData.bind(this));
        }

        if (periodSelect) {
            periodSelect.addEventListener('change', this.filterByPeriod.bind(this));
        }

        if (exportPerformance) {
            exportPerformance.addEventListener('click', () => this.exportChart('performance'));
        }

        if (exportFundamentals) {
            exportFundamentals.addEventListener('click', () => this.exportChart('fundamentals'));
        }
    }

    async tryAutoLoadFile() {
        try {
            this.showLoading(true);
            this.showStatus('Tentando carregar mluca.xlsx automaticamente...', 'info');

            const response = await fetch('./mluca.xlsx');
            if (response.ok) {
                const arrayBuffer = await response.arrayBuffer();
                await this.processExcelData(arrayBuffer);
                this.showStatus('Arquivo mluca.xlsx carregado automaticamente!', 'success');
            } else {
                throw new Error('Arquivo não encontrado');
            }
        } catch (error) {
            console.log('Carregamento automático falhou:', error.message);
            this.showStatus('Arquivo mluca.xlsx não encontrado. Use o upload manual.', 'info');
        } finally {
            this.showLoading(false);
        }
    }

    handleDragOver(e) {
        e.preventDefault();
        document.getElementById('uploadArea').classList.add('dragover');
    }

    handleDragLeave(e) {
        e.preventDefault();
        document.getElementById('uploadArea').classList.remove('dragover');
    }

    handleDrop(e) {
        e.preventDefault();
        document.getElementById('uploadArea').classList.remove('dragover');

        const files = e.dataTransfer.files;
        if (files.length > 0) {
            this.processFile(files[0]);
        }
    }

    handleFileSelect(e) {
        const file = e.target.files[0];
        if (file) {
            this.processFile(file);
        }
    }

    async processFile(file) {
        if (!file.name.match(/\.(xlsx|xls)$/i)) {
            this.showStatus('Por favor, selecione um arquivo Excel (.xlsx ou .xls)', 'error');
            return;
        }

        try {
            this.showLoading(true);
            this.showStatus('Processando arquivo Excel...', 'info');

            const arrayBuffer = await file.arrayBuffer();
            await this.processExcelData(arrayBuffer);

            this.showStatus(`Arquivo ${file.name} processado com sucesso!`, 'success');
        } catch (error) {
            console.error('Erro ao processar arquivo:', error);
            this.showStatus('Erro ao processar o arquivo. Verifique o formato.', 'error');
        } finally {
            this.showLoading(false);
        }
    }

    async processExcelData(arrayBuffer) {
        return new Promise((resolve, reject) => {
            try {
                readXlsxFile(arrayBuffer).then((rows) => {
                    if (rows.length < 2) {
                        throw new Error('Arquivo Excel vazio ou sem dados');
                    }

                    // Header row
                    const headers = rows[0];
                    console.log('Headers encontrados:', headers);

                    // Validar estrutura necessária
                    const requiredColumns = ['Mês', 'MLUCA (acc)', 'IBOV (acc)', 'CDI (acc)', 'Vol (ano)'];
                    const missingColumns = requiredColumns.filter(col => !headers.includes(col));

                    if (missingColumns.length > 0) {
                        throw new Error(`Colunas obrigatórias não encontradas: ${missingColumns.join(', ')}`);
                    }

                    // Processar dados
                    this.data = [];
                    for (let i = 1; i < rows.length; i++) {
                        const row = rows[i];
                        const dataPoint = {};

                        headers.forEach((header, index) => {
                            dataPoint[header] = row[index];
                        });

                        // Validar e converter data
                        if (dataPoint['Mês']) {
                            const dateValue = dataPoint['Mês'];
                            if (dateValue instanceof Date) {
                                dataPoint['Mês'] = dateValue;
                            } else if (typeof dateValue === 'string') {
                                dataPoint['Mês'] = new Date(dateValue);
                            }
                        }

                        this.data.push(dataPoint);
                    }

                    // Ordenar por data
                    this.data.sort((a, b) => new Date(a['Mês']) - new Date(b['Mês']));

                    console.log('Dados processados:', this.data.length, 'registros');
                    console.log('Amostra dos dados:', this.data.slice(0, 3));

                    // Aplicar dados ao dashboard
                    this.filteredData = [...this.data];
                    this.updateDashboard();

                    resolve();
                }).catch(reject);
            } catch (error) {
                reject(error);
            }
        });
    }

    updateDashboard() {
        // Mostrar seção do dashboard
        document.getElementById('uploadSection').style.display = 'none';
        document.getElementById('dashboardContent').style.display = 'block';

        // Atualizar informações
        document.getElementById('dataInfo').textContent = `${this.data.length} registros carregados`;

        // Atualizar KPIs
        this.updateKPIs();

        // Criar gráficos
        this.createCharts();

        // Atualizar tabela de performance
        this.updatePerformanceTable();
    }

    updateKPIs() {
        if (this.filteredData.length === 0) return;

        const lastData = this.filteredData[this.filteredData.length - 1];

        // Performance MLUCA
        const mlucaPerf = lastData['MLUCA (acc)'] || 0;
        document.querySelector('#kpiPerformance .kpi-value').textContent = 
            this.formatPercentage(mlucaPerf);

        // vs IBOVESPA
        const ibovPerf = lastData['IBOV (acc)'] || 0;
        const diff = mlucaPerf - ibovPerf;
        const diffElement = document.querySelector('#kpiVsIbov .kpi-value');
        diffElement.textContent = this.formatPercentage(diff, true);
        diffElement.className = `kpi-value ${diff >= 0 ? 'text-success' : 'text-danger'}`;

        // Dividend Yield
        const dy = lastData['DY(%)'];
        document.querySelector('#kpiDividendYield .kpi-value').textContent = 
            dy ? this.formatPercentage(dy) : '--';

        // Volatilidade
        const vol = lastData['Vol (ano)'] || 0;
        document.querySelector('#kpiVolatilidade .kpi-value').textContent = 
            this.formatPercentage(vol);
    }

    updatePerformanceTable() {
        if (this.data.length === 0) return;

        const lastData = this.data[this.data.length - 1];
        const currentYear = new Date(lastData['Mês']).getFullYear();
        const previousYear = currentYear - 1;

        // Procurar dados de dezembro do ano anterior para cálculo YTD
        const decemberData = this.findDecemberData(previousYear);

        // Dados do mês (última linha)
        const monthMLUCA = (lastData['MLUCA (mês)'] || 0) * 100;
        const monthIBOV = (lastData['IBOV (mes)'] || 0) * 100;
        const monthCDI = (lastData['CDI (mês)'] || 0) * 100;

        // Dados acumulados (desde o início)
        const allMLUCA = (lastData['MLUCA (acc)'] || 0) * 100;
        const allIBOV = (lastData['IBOV (acc)'] || 0) * 100;
        const allCDI = (lastData['CDI (acc)'] || 0) * 100;

        // Calcular YTD
        let ytdMLUCA = 0, ytdIBOV = 0, ytdCDI = 0;
        if (decemberData) {
            ytdMLUCA = ((lastData['MLUCA (cota)'] - decemberData['MLUCA (cota)']) / decemberData['MLUCA (cota)']) * 100;
            ytdIBOV = ((lastData['IBOV (pts)'] - decemberData['IBOV (pts)']) / decemberData['IBOV (pts)']) * 100;
            ytdCDI = ((lastData['CDI (100)'] - decemberData['CDI (100)']) / decemberData['CDI (100)']) * 100;
        }

        // Atualizar tabela HTML
        this.updateTableCell('mluca-month', monthMLUCA);
        this.updateTableCell('mluca-ytd', ytdMLUCA);
        this.updateTableCell('mluca-all', allMLUCA);

        this.updateTableCell('ibov-month', monthIBOV);
        this.updateTableCell('ibov-ytd', ytdIBOV);
        this.updateTableCell('ibov-all', allIBOV);

        this.updateTableCell('cdi-month', monthCDI);
        this.updateTableCell('cdi-ytd', ytdCDI);
        this.updateTableCell('cdi-all', allCDI);

        console.log('Tabela de performance atualizada:', {
            month: { mluca: monthMLUCA, ibov: monthIBOV, cdi: monthCDI },
            ytd: { mluca: ytdMLUCA, ibov: ytdIBOV, cdi: ytdCDI },
            all: { mluca: allMLUCA, ibov: allIBOV, cdi: allCDI }
        });
    }

    findDecemberData(year) {
        // Procurar dezembro do ano especificado
        const decemberData = this.data.filter(item => {
            const itemDate = new Date(item['Mês']);
            return itemDate.getFullYear() === year && itemDate.getMonth() === 11;
        });

        if (decemberData.length > 0) {
            return decemberData[decemberData.length - 1];
        }

        // Se não encontrar dezembro, pegar o último registro do ano
        const yearData = this.data.filter(item => {
            return new Date(item['Mês']).getFullYear() === year;
        });

        return yearData.length > 0 ? yearData[yearData.length - 1] : null;
    }

    updateTableCell(cellId, value) {
        const cell = document.getElementById(cellId);
        if (cell) {
            const formattedValue = this.formatPercentage(value / 100, true);
            cell.textContent = formattedValue;

            // Aplicar classe de cor
            if (value >= 0.01) {
                cell.className = 'positive';
            } else if (value <= -0.01) {
                cell.className = 'negative';
            } else {
                cell.className = 'neutral';
            }
        }
    }

    createCharts() {
        this.createPerformanceChart();
        this.createFundamentalsChart();
    }

    createPerformanceChart() {
        const ctx = document.getElementById('performanceChart');
        if (!ctx) return;

        // Destruir gráfico anterior se existir
        if (this.performanceChart) {
            this.performanceChart.destroy();
        }

        const labels = this.filteredData.map(item => this.formatDate(item['Mês']));

        const mlucaData = this.filteredData.map(item => (item['MLUCA (acc)'] || 0) * 100);
        const ibovData = this.filteredData.map(item => (item['IBOV (acc)'] || 0) * 100);
        const cdiData = this.filteredData.map(item => (item['CDI (acc)'] || 0) * 100);
        const volData = this.filteredData.map(item => (item['Vol (ano)'] || 0) * 100);

        this.performanceChart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: labels,
                datasets: [
                    {
                        label: 'MLUCA (acc)',
                        data: mlucaData,
                        borderColor: '#2c5aa0',
                        backgroundColor: 'rgba(44, 90, 160, 0.1)',
                        borderWidth: 3,
                        fill: false,
                        tension: 0.2,
                        yAxisID: 'y'
                    },
                    {
                        label: 'IBOV (acc)',
                        data: ibovData,
                        borderColor: '#e53e3e',
                        backgroundColor: 'rgba(229, 62, 62, 0.1)',
                        borderWidth: 2,
                        fill: false,
                        tension: 0.2,
                        yAxisID: 'y'
                    },
                    {
                        label: 'CDI (acc)',
                        data: cdiData,
                        borderColor: '#38a169',
                        backgroundColor: 'rgba(56, 161, 105, 0.1)',
                        borderWidth: 2,
                        fill: false,
                        tension: 0.2,
                        yAxisID: 'y'
                    },
                    {
                        label: 'Volatilidade (ano)',
                        data: volData,
                        type: 'bar',
                        backgroundColor: 'rgba(26, 54, 93, 0.3)',
                        borderColor: '#1a365d',
                        borderWidth: 1,
                        yAxisID: 'y1'
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    title: {
                        display: false
                    },
                    legend: {
                        position: 'top',
                        labels: {
                            usePointStyle: true,
                            padding: 20
                        }
                    },
                    tooltip: {
                        mode: 'index',
                        intersect: false,
                        callbacks: {
                            label: function(context) {
                                const label = context.dataset.label || '';
                                const value = context.parsed.y;
                                return `${label}: ${value.toFixed(2)}%`;
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        grid: {
                            display: true,
                            color: 'rgba(0,0,0,0.05)'
                        }
                    },
                    y: {
                        type: 'linear',
                        display: true,
                        position: 'left',
                        title: {
                            display: true,
                            text: 'Performance Acumulada (%)'
                        },
                        grid: {
                            color: 'rgba(0,0,0,0.05)'
                        }
                    },
                    y1: {
                        type: 'linear',
                        display: true,
                        position: 'right',
                        title: {
                            display: true,
                            text: 'Volatilidade Anual (%)'
                        },
                        grid: {
                            drawOnChartArea: false,
                        },
                    }
                },
                interaction: {
                    mode: 'nearest',
                    axis: 'x',
                    intersect: false
                }
            }
        });
    }

    createFundamentalsChart() {
        const ctx = document.getElementById('fundamentalsChart');
        if (!ctx) return;

        // Destruir gráfico anterior se existir
        if (this.fundamentalsChart) {
            this.fundamentalsChart.destroy();
        }

        // Filtrar apenas dados com DY e GAP válidos
        const validData = this.filteredData.filter(item => 
            item['DY(%)'] != null && item['GAP (risco)'] != null &&
            !isNaN(item['DY(%)']) && !isNaN(item['GAP (risco)'])
        );

        if (validData.length === 0) {
            // Se não há dados válidos, criar gráfico vazio
            this.fundamentalsChart = new Chart(ctx, {
                type: 'line',
                data: {
                    labels: [],
                    datasets: []
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        title: {
                            display: true,
                            text: 'Dados fundamentalistas não disponíveis'
                        }
                    }
                }
            });
            return;
        }

        const labels = validData.map(item => this.formatDate(item['Mês']));
        const dyData = validData.map(item => (item['DY(%)'] || 0) * 100);
        const gapData = validData.map(item => (item['GAP (risco)'] || 0) * 100);

        this.fundamentalsChart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: labels,
                datasets: [
                    {
                        label: 'Dividend Yield (%)',
                        data: dyData,
                        borderColor: '#d69e2e',
                        backgroundColor: 'rgba(214, 158, 46, 0.1)',
                        borderWidth: 3,
                        fill: false,
                        tension: 0.2,
                        yAxisID: 'y'
                    },
                    {
                        label: 'GAP de Risco (%)',
                        data: gapData,
                        borderColor: '#1a365d',
                        backgroundColor: 'rgba(26, 54, 93, 0.1)',
                        borderWidth: 3,
                        fill: false,
                        tension: 0.2,
                        yAxisID: 'y1'
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    title: {
                        display: false
                    },
                    legend: {
                        position: 'top',
                        labels: {
                            usePointStyle: true,
                            padding: 20
                        }
                    },
                    tooltip: {
                        mode: 'index',
                        intersect: false,
                        callbacks: {
                            label: function(context) {
                                const label = context.dataset.label || '';
                                const value = context.parsed.y;
                                return `${label}: ${value.toFixed(2)}%`;
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        grid: {
                            display: true,
                            color: 'rgba(0,0,0,0.05)'
                        }
                    },
                    y: {
                        type: 'linear',
                        display: true,
                        position: 'left',
                        title: {
                            display: true,
                            text: 'Dividend Yield (%)'
                        },
                        grid: {
                            color: 'rgba(0,0,0,0.05)'
                        }
                    },
                    y1: {
                        type: 'linear',
                        display: true,
                        position: 'right',
                        title: {
                            display: true,
                            text: 'GAP de Risco (%)'
                        },
                        grid: {
                            drawOnChartArea: false,
                        }
                    }
                },
                interaction: {
                    mode: 'nearest',
                    axis: 'x',
                    intersect: false
                }
            }
        });
    }

    filterByPeriod() {
        const periodSelect = document.getElementById('periodSelect');
        if (!periodSelect) return;

        const period = periodSelect.value;

        if (period === 'all') {
            this.filteredData = [...this.data];
        } else {
            const months = parseInt(period);
            const cutoffDate = new Date();
            cutoffDate.setMonth(cutoffDate.getMonth() - months);

            this.filteredData = this.data.filter(item => 
                new Date(item['Mês']) >= cutoffDate
            );
        }

        this.updateKPIs();
        this.createCharts();
        // Não atualizar a tabela de performance pois ela sempre usa todos os dados
    }

    refreshData() {
        this.tryAutoLoadFile();
    }

    exportChart(chartType) {
        let chart, filename;

        if (chartType === 'performance') {
            chart = this.performanceChart;
            filename = 'mluca_performance_chart.png';
        } else if (chartType === 'fundamentals') {
            chart = this.fundamentalsChart;
            filename = 'mluca_fundamentals_chart.png';
        }

        if (chart) {
            const url = chart.toBase64Image();
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            a.click();
        }
    }

    // Utility functions
    formatDate(date) {
        if (!date) return '';

        const d = new Date(date);
        if (isNaN(d.getTime())) return '';

        const months = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun',
                       'jul', 'ago', 'set', 'out', 'nov', 'dez'];

        const month = months[d.getMonth()];
        const year = d.getFullYear().toString().substr(-2);

        return `${month}/${year}`;
    }

    formatPercentage(value, showSign = false) {
        if (value == null || isNaN(value)) return '--';

        const percentage = (value * 100).toFixed(2);
        const sign = showSign && value >= 0 ? '+' : '';

        return `${sign}${percentage}%`;
    }

    formatNumber(value, decimals = 2) {
        if (value == null || isNaN(value)) return '--';

        return value.toLocaleString('pt-BR', {
            minimumFractionDigits: decimals,
            maximumFractionDigits: decimals
        });
    }

    showLoading(show) {
        const overlay = document.getElementById('loadingOverlay');
        if (overlay) {
            overlay.style.display = show ? 'flex' : 'none';
        }
    }

    showStatus(message, type) {
        const statusDiv = document.getElementById('uploadStatus');
        if (statusDiv) {
            statusDiv.textContent = message;
            statusDiv.className = `upload-status ${type}`;
            statusDiv.style.display = 'block';

            // Auto-hide success messages after 5 seconds
            if (type === 'success') {
                setTimeout(() => {
                    statusDiv.style.display = 'none';
                }, 5000);
            }
        }
    }
}

// Inicializar dashboard quando DOM estiver pronto
document.addEventListener('DOMContentLoaded', function() {
    new MLUCADashboard();
});