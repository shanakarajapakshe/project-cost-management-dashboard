// Data Management Module
class DataManager {
    constructor() {
        this.currentMonth = null;
        this.currentYear = null;
        this.projects = [];
        this.monthlyData = {}; // Store data for trend analysis
        this.isElectron = typeof window.electronAPI !== 'undefined';
    }

    setCurrentPeriod(month, year) {
        this.currentMonth = month;
        this.currentYear = year;
    }

    getCurrentPeriod() {
        return {
            month: this.currentMonth,
            year: this.currentYear,
            key: `${this.currentYear}-${this.currentMonth}`
        };
    }

    async addProject(projectData) {
        const calculatedData = this.calculateProjectMetrics(projectData);
        
        if (this.isElectron) {
            // Save to Excel file via Electron
            const result = await window.electronAPI.saveProject(
                this.currentMonth, 
                this.currentYear, 
                calculatedData
            );
            
            if (result.success) {
                this.projects.push(calculatedData);
                this.updateMonthlyData();
            } else {
                throw new Error(result.error || 'Failed to save project');
            }
        } else {
            // Fallback to localStorage for web version
            this.projects.push(calculatedData);
            this.saveToLocalStorage();
        }
        
        return calculatedData;
    }

    calculateProjectMetrics(data) {
        const engineerCost = data.engineerSalary;
        const ceVisitCost = data.ceVisitCharge * data.visitsPerMonth;
        const directCost = engineerCost + ceVisitCost + data.transportCost;
        const overheadCost = directCost * (data.overheadAllocation / 100);
        const totalCost = directCost + overheadCost;
        const profit = data.clientPayment - totalCost;

        return {
            ...data,
            engineerCost: engineerCost,
            ceVisitCost: ceVisitCost,
            directCost: directCost,
            overheadCost: overheadCost,
            totalCost: totalCost,
            profit: profit,
            timestamp: new Date().toISOString()
        };
    }

    getProjects() {
        return this.projects;
    }

    async deleteProject(index) {
        const project = this.projects[index];
        
        if (this.isElectron) {
            const result = await window.electronAPI.deleteProject(
                this.currentMonth,
                this.currentYear,
                project.projectName
            );
            
            if (result.success) {
                this.projects.splice(index, 1);
                this.updateMonthlyData();
            } else {
                throw new Error(result.error || 'Failed to delete project');
            }
        } else {
            this.projects.splice(index, 1);
            this.saveToLocalStorage();
        }
    }

    getSummaryStats() {
        const totalProjects = this.projects.length;
        const totalRevenue = this.projects.reduce((sum, p) => sum + p.clientPayment, 0);
        const totalCosts = this.projects.reduce((sum, p) => sum + p.totalCost, 0);
        const netProfit = totalRevenue - totalCosts;

        return { totalProjects, totalRevenue, totalCosts, netProfit };
    }

    updateMonthlyData() {
        const period = this.getCurrentPeriod();
        this.monthlyData[period.key] = {
            projects: [...this.projects],
            stats: this.getSummaryStats()
        };
        
        if (!this.isElectron) {
            localStorage.setItem('profit_dashboard_monthly', JSON.stringify(this.monthlyData));
        }
    }

    saveToLocalStorage() {
        const period = this.getCurrentPeriod();
        const key = `profit_dashboard_${period.key}`;
        localStorage.setItem(key, JSON.stringify(this.projects));
        this.updateMonthlyData();
        this.createBackup();
    }

    async loadFromLocalStorage() {
        const period = this.getCurrentPeriod();
        
        if (this.isElectron) {
            // Load from Excel file via Electron
            const result = await window.electronAPI.loadOrCreateExcel(
                this.currentMonth,
                this.currentYear
            );
            
            if (result.success) {
                this.projects = result.data || [];
            } else {
                console.error('Error loading from Excel:', result.error);
                this.projects = [];
            }
        } else {
            // Fallback to localStorage for web version
            const key = `profit_dashboard_${period.key}`;
            const data = localStorage.getItem(key);
            
            if (data) {
                this.projects = JSON.parse(data);
            } else {
                this.projects = [];
            }
        }

        // Load monthly data for trends
        const monthlyData = localStorage.getItem('profit_dashboard_monthly');
        if (monthlyData) {
            this.monthlyData = JSON.parse(monthlyData);
        }
        
        this.updateMonthlyData();
        return this.projects;
    }

    async createBackup() {
        if (this.isElectron) {
            try {
                const result = await window.electronAPI.createBackup(
                    this.currentMonth,
                    this.currentYear
                );
                console.log('Backup created:', result.path);
            } catch (error) {
                console.error('Error creating backup:', error);
            }
        } else {
            // localStorage backup for web version
            const period = this.getCurrentPeriod();
            const backupKey = `profit_dashboard_backup_${period.key}_${Date.now()}`;
            localStorage.setItem(backupKey, JSON.stringify(this.projects));
            this.cleanOldBackups(period.key);
        }
    }

    cleanOldBackups(periodKey) {
        const allKeys = Object.keys(localStorage);
        const backupKeys = allKeys
            .filter(key => key.startsWith(`profit_dashboard_backup_${periodKey}`))
            .sort()
            .reverse();
        
        // Remove old backups, keep only 5
        if (backupKeys.length > 5) {
            backupKeys.slice(5).forEach(key => localStorage.removeItem(key));
        }
    }

    getMonthlyTrend() {
        // Get last 6 months of data for trend chart
        const sortedKeys = Object.keys(this.monthlyData).sort();
        const last6Months = sortedKeys.slice(-6);
        
        return last6Months.map(key => ({
            period: key,
            profit: this.monthlyData[key].stats.netProfit
        }));
    }

    async exportToExcel() {
        if (this.isElectron) {
            try {
                const result = await window.electronAPI.exportExcelFile(
                    this.currentMonth,
                    this.currentYear
                );
                
                if (result.success && !result.canceled) {
                    return { success: true, path: result.path };
                } else if (result.canceled) {
                    return { success: false, canceled: true };
                } else {
                    return { success: false, error: result.error };
                }
            } catch (error) {
                console.error('Error exporting Excel:', error);
                return { success: false, error: error.message };
            }
        } else {
            console.log('Excel export is only available in the desktop version');
            return { success: false, error: 'Not available in web version' };
        }
    }
}

// UI Controller
class UIController {
    constructor(dataManager) {
        this.dataManager = dataManager;
        this.charts = {};
        this.initializeElements();
        this.initializeCharts();
    }

    initializeElements() {
        this.elements = {
            monthSelect: document.getElementById('monthSelect'),
            yearSelect: document.getElementById('yearSelect'),
            loadMonthBtn: document.getElementById('loadMonthBtn'),
            currentPeriod: document.getElementById('currentPeriod'),
            mainContent: document.getElementById('mainContent'),
            projectForm: document.getElementById('projectForm'),
            resetFormBtn: document.getElementById('resetFormBtn'),
            projectsTableBody: document.getElementById('projectsTableBody'),
            totalProjects: document.getElementById('totalProjects'),
            totalRevenue: document.getElementById('totalRevenue'),
            totalCosts: document.getElementById('totalCosts'),
            netProfit: document.getElementById('netProfit'),
            toast: document.getElementById('toast'),
            toastMessage: document.getElementById('toastMessage')
        };
    }

    populateYearSelect() {
        const currentYear = new Date().getFullYear();
        const yearSelect = this.elements.yearSelect;
        
        for (let year = currentYear - 5; year <= currentYear + 5; year++) {
            const option = document.createElement('option');
            option.value = year;
            option.textContent = year;
            if (year === currentYear) {
                option.selected = true;
            }
            yearSelect.appendChild(option);
        }
    }

    setCurrentMonth() {
        const currentMonth = new Date().getMonth() + 1;
        this.elements.monthSelect.value = currentMonth.toString().padStart(2, '0');
    }

    showMainContent() {
        this.elements.mainContent.style.display = 'block';
        this.elements.mainContent.classList.add('fade-in');
    }

    hideMainContent() {
        this.elements.mainContent.style.display = 'none';
    }

    updateCurrentPeriodDisplay() {
        const period = this.dataManager.getCurrentPeriod();
        const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                          'July', 'August', 'September', 'October', 'November', 'December'];
        const monthName = monthNames[parseInt(period.month) - 1];
        
        this.elements.currentPeriod.innerHTML = `
            <strong>Current Period:</strong> ${monthName} ${period.year}
            <span class="ms-3 badge bg-primary">${this.dataManager.projects.length} Projects</span>
        `;
    }

    updateSummaryStats() {
        const stats = this.dataManager.getSummaryStats();
        
        this.elements.totalProjects.textContent = stats.totalProjects;
        this.elements.totalRevenue.textContent = this.formatCurrency(stats.totalRevenue);
        this.elements.totalCosts.textContent = this.formatCurrency(stats.totalCosts);
        this.elements.netProfit.textContent = this.formatCurrency(stats.netProfit);
        
        // Color code net profit
        if (stats.netProfit >= 0) {
            this.elements.netProfit.className = 'stat-value profit-positive';
        } else {
            this.elements.netProfit.className = 'stat-value profit-negative';
        }
    }

    updateProjectsTable() {
        const projects = this.dataManager.getProjects();
        const tbody = this.elements.projectsTableBody;
        
        if (projects.length === 0) {
            tbody.innerHTML = '<tr><td colspan="5" class="text-center text-muted">No projects added yet</td></tr>';
            return;
        }

        tbody.innerHTML = projects.map((project, index) => `
            <tr>
                <td><strong>${project.projectName}</strong></td>
                <td>${this.formatCurrency(project.totalCost)}</td>
                <td>${this.formatCurrency(project.clientPayment)}</td>
                <td class="${project.profit >= 0 ? 'profit-positive' : 'profit-negative'}">
                    ${this.formatCurrency(project.profit)}
                </td>
                <td>
                    <button class="btn btn-sm btn-info" onclick="uiController.viewProjectDetails(${index})">View</button>
                    <button class="btn btn-sm btn-danger" onclick="uiController.deleteProject(${index})">Delete</button>
                </td>
            </tr>
        `).join('');
    }

    viewProjectDetails(index) {
        const project = this.dataManager.projects[index];
        const details = `
            <strong>Project:</strong> ${project.projectName}<br>
            <strong>Engineers:</strong> ${project.numEngineers} × ${this.formatCurrency(project.engineerSalary)} = ${this.formatCurrency(project.engineerCost)}<br>
            <strong>CE Visits:</strong> ${project.visitsPerMonth} × ${this.formatCurrency(project.ceVisitCharge)} = ${this.formatCurrency(project.ceVisitCost)}<br>
            <strong>Transport:</strong> ${this.formatCurrency(project.transportCost)}<br>
            <strong>Direct Cost:</strong> ${this.formatCurrency(project.directCost)}<br>
            <strong>Overhead (${project.overheadAllocation}%):</strong> ${this.formatCurrency(project.overheadCost)}<br>
            <strong>Total Cost:</strong> ${this.formatCurrency(project.totalCost)}<br>
            <strong>Client Payment:</strong> ${this.formatCurrency(project.clientPayment)}<br>
            <strong>Profit:</strong> ${this.formatCurrency(project.profit)}
        `;
        this.showToast(details, 'info');
    }

    async deleteProject(index) {
        if (confirm('Are you sure you want to delete this project?')) {
            try {
                await this.dataManager.deleteProject(index);
                this.refreshDashboard();
                this.showToast('Project deleted successfully', 'success');
            } catch (error) {
                this.showToast('Error deleting project: ' + error.message, 'danger');
            }
        }
    }

    async exportExcel() {
        try {
            const result = await this.dataManager.exportToExcel();
            if (result.success) {
                this.showToast('Excel file exported successfully to: ' + result.path, 'success');
            } else if (result.canceled) {
                this.showToast('Export canceled', 'info');
            } else {
                this.showToast('Error exporting: ' + result.error, 'danger');
            }
        } catch (error) {
            this.showToast('Error exporting Excel: ' + error.message, 'danger');
        }
    }

    refreshDashboard() {
        this.updateSummaryStats();
        this.updateProjectsTable();
        this.updateCurrentPeriodDisplay();
        this.updateAllCharts();
    }

    initializeCharts() {
        const chartOptions = {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    display: true,
                    position: 'top'
                }
            }
        };

        // Cost vs Payment Chart
        this.charts.costPayment = new Chart(
            document.getElementById('costPaymentChart'),
            {
                type: 'bar',
                data: {
                    labels: [],
                    datasets: [
                        {
                            label: 'Total Cost',
                            data: [],
                            backgroundColor: 'rgba(255, 99, 132, 0.7)',
                            borderColor: 'rgba(255, 99, 132, 1)',
                            borderWidth: 1
                        },
                        {
                            label: 'Client Payment',
                            data: [],
                            backgroundColor: 'rgba(75, 192, 192, 0.7)',
                            borderColor: 'rgba(75, 192, 192, 1)',
                            borderWidth: 1
                        }
                    ]
                },
                options: chartOptions
            }
        );

        // Profit Chart
        this.charts.profit = new Chart(
            document.getElementById('profitChart'),
            {
                type: 'bar',
                data: {
                    labels: [],
                    datasets: [{
                        label: 'Profit',
                        data: [],
                        backgroundColor: [],
                        borderColor: [],
                        borderWidth: 1
                    }]
                },
                options: chartOptions
            }
        );

        // Trend Chart
        this.charts.trend = new Chart(
            document.getElementById('trendChart'),
            {
                type: 'line',
                data: {
                    labels: [],
                    datasets: [{
                        label: 'Monthly Profit',
                        data: [],
                        borderColor: 'rgba(54, 162, 235, 1)',
                        backgroundColor: 'rgba(54, 162, 235, 0.2)',
                        tension: 0.4,
                        fill: true
                    }]
                },
                options: chartOptions
            }
        );

        // Overhead Chart
        this.charts.overhead = new Chart(
            document.getElementById('overheadChart'),
            {
                type: 'scatter',
                data: {
                    datasets: [{
                        label: 'Overhead % vs Profit',
                        data: [],
                        backgroundColor: 'rgba(153, 102, 255, 0.7)',
                        borderColor: 'rgba(153, 102, 255, 1)',
                        pointRadius: 6
                    }]
                },
                options: {
                    ...chartOptions,
                    scales: {
                        x: {
                            title: {
                                display: true,
                                text: 'Overhead %'
                            }
                        },
                        y: {
                            title: {
                                display: true,
                                text: 'Profit'
                            }
                        }
                    }
                }
            }
        );
    }

    updateAllCharts() {
        const projects = this.dataManager.getProjects();
        
        // Update Cost vs Payment Chart
        this.charts.costPayment.data.labels = projects.map(p => p.projectName);
        this.charts.costPayment.data.datasets[0].data = projects.map(p => p.totalCost);
        this.charts.costPayment.data.datasets[1].data = projects.map(p => p.clientPayment);
        this.charts.costPayment.update();

        // Update Profit Chart
        this.charts.profit.data.labels = projects.map(p => p.projectName);
        this.charts.profit.data.datasets[0].data = projects.map(p => p.profit);
        this.charts.profit.data.datasets[0].backgroundColor = projects.map(p => 
            p.profit >= 0 ? 'rgba(75, 192, 192, 0.7)' : 'rgba(255, 99, 132, 0.7)'
        );
        this.charts.profit.data.datasets[0].borderColor = projects.map(p => 
            p.profit >= 0 ? 'rgba(75, 192, 192, 1)' : 'rgba(255, 99, 132, 1)'
        );
        this.charts.profit.update();

        // Update Trend Chart
        const trendData = this.dataManager.getMonthlyTrend();
        this.charts.trend.data.labels = trendData.map(d => d.period);
        this.charts.trend.data.datasets[0].data = trendData.map(d => d.profit);
        this.charts.trend.update();

        // Update Overhead Chart
        this.charts.overhead.data.datasets[0].data = projects.map(p => ({
            x: p.overheadAllocation,
            y: p.profit
        }));
        this.charts.overhead.update();
    }

    formatCurrency(amount) {
        return new Intl.NumberFormat('en-US', {
            style: 'currency',
            currency: 'LKR',
            minimumFractionDigits: 2
        }).format(amount);
    }

    showToast(message, type = 'info') {
        this.elements.toastMessage.textContent = message;
        const toast = new bootstrap.Toast(this.elements.toast);
        toast.show();
    }

    resetForm() {
        this.elements.projectForm.reset();
    }

    printDashboard() {
        // Save the current page title
        const originalTitle = document.title;
        const period = this.dataManager.getCurrentPeriod();
        
        // Set a descriptive title for the print
        document.title = `Profit Dashboard - ${period.month}/${period.year}`;
        
        // Trigger the print dialog
        window.print();
        
        // Restore the original title after printing
        document.title = originalTitle;
    }
}

// Application Controller
class AppController {
    constructor() {
        this.dataManager = new DataManager();
        this.uiController = new UIController(this.dataManager);
        this.initializeApp();
    }

    initializeApp() {
        this.uiController.populateYearSelect();
        this.uiController.setCurrentMonth();
        this.setupEventListeners();
    }

    setupEventListeners() {
        // Load Month Button
        this.uiController.elements.loadMonthBtn.addEventListener('click', () => {
            this.loadMonthData();
        });

        // Project Form Submit
        this.uiController.elements.projectForm.addEventListener('submit', (e) => {
            e.preventDefault();
            this.handleFormSubmit();
        });

        // Reset Form Button
        this.uiController.elements.resetFormBtn.addEventListener('click', () => {
            this.uiController.resetForm();
        });
    }

    async loadMonthData() {
        const month = this.uiController.elements.monthSelect.value;
        const year = this.uiController.elements.yearSelect.value;

        this.dataManager.setCurrentPeriod(month, year);
        
        try {
            await this.dataManager.loadFromLocalStorage();
            
            this.uiController.showMainContent();
            this.uiController.updateCurrentPeriodDisplay();
            this.uiController.refreshDashboard();
            
            this.uiController.showToast(`Loaded data for ${month}/${year}`, 'success');
        } catch (error) {
            this.uiController.showToast('Error loading data: ' + error.message, 'danger');
        }
    }

    async handleFormSubmit() {
        const formData = {
            projectName: document.getElementById('projectName').value.trim(),
            numEngineers: parseFloat(document.getElementById('numEngineers').value),
            engineerSalary: parseFloat(document.getElementById('engineerSalary').value),
            ceVisitCharge: parseFloat(document.getElementById('ceVisitCharge').value),
            visitsPerMonth: parseFloat(document.getElementById('visitsPerMonth').value),
            transportCost: parseFloat(document.getElementById('transportCost').value),
            clientPayment: parseFloat(document.getElementById('clientPayment').value),
            overheadAllocation: parseFloat(document.getElementById('overheadAllocation').value)
        };

        // Validation
        if (!this.validateFormData(formData)) {
            this.uiController.showToast('Please fill all fields with valid numbers', 'danger');
            return;
        }

        try {
            // Add project
            await this.dataManager.addProject(formData);
            this.uiController.refreshDashboard();
            this.uiController.resetForm();
            this.uiController.showToast('Project added successfully!', 'success');
        } catch (error) {
            this.uiController.showToast('Error saving project: ' + error.message, 'danger');
        }
    }

    validateFormData(data) {
        if (!data.projectName) return false;
        
        const numericFields = [
            'numEngineers', 'engineerSalary', 'ceVisitCharge',
            'visitsPerMonth', 'transportCost', 'clientPayment', 'overheadAllocation'
        ];

        for (let field of numericFields) {
            if (isNaN(data[field]) || data[field] < 0) {
                return false;
            }
        }

        return true;
    }
}

// Initialize Application
let appController;
let uiController;

document.addEventListener('DOMContentLoaded', () => {
    appController = new AppController();
    uiController = appController.uiController;
});
