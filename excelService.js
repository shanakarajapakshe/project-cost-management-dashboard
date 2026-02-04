const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

class ExcelService {
    constructor(basePath) {
        this.basePath = basePath;
        this.workbooks = {}; // Cache workbooks in memory
    }

    getFilePath(month, year) {
        return path.join(this.basePath, `Profit_Dashboard_${year}_${month}.xlsx`);
    }

    getBackupPath(month, year) {
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        return path.join(this.basePath, 'backups', `Backup_${year}_${month}_${timestamp}.xlsx`);
    }

    async loadOrCreateMonthFile(month, year) {
        const filePath = this.getFilePath(month, year);
        
        if (fs.existsSync(filePath)) {
            return await this.loadExistingFile(month, year);
        } else {
            return await this.createNewFile(month, year);
        }
    }

    async loadExistingFile(month, year) {
        const filePath = this.getFilePath(month, year);
        const workbook = new ExcelJS.Workbook();
        
        try {
            await workbook.xlsx.readFile(filePath);
            this.workbooks[`${year}-${month}`] = workbook;
            
            const projects = await this.readProjectsFromWorkbook(workbook);
            return projects;
        } catch (error) {
            console.error('Error loading Excel file:', error);
            throw error;
        }
    }

    async createNewFile(month, year) {
        const workbook = new ExcelJS.Workbook();
        
        // Create Projects Sheet
        const projectsSheet = workbook.addWorksheet('Projects');
        
        // Define headers with styling
        projectsSheet.columns = [
            { header: 'Project Name', key: 'projectName', width: 25 },
            { header: 'No. of Engineers', key: 'numEngineers', width: 15 },
            { header: 'Engineer Salary/Month', key: 'engineerSalary', width: 20 },
            { header: 'CE Visit Charge', key: 'ceVisitCharge', width: 18 },
            { header: 'Visits/Month', key: 'visitsPerMonth', width: 15 },
            { header: 'Transport Cost/Month', key: 'transportCost', width: 20 },
            { header: 'Client Payment/Month', key: 'clientPayment', width: 20 },
            { header: 'Overhead Allocation %', key: 'overheadAllocation', width: 20 },
            { header: 'Engineer Cost', key: 'engineerCost', width: 18 },
            { header: 'CE Visit Cost', key: 'ceVisitCost', width: 18 },
            { header: 'Direct Cost', key: 'directCost', width: 18 },
            { header: 'Overhead Cost', key: 'overheadCost', width: 18 },
            { header: 'Total Cost', key: 'totalCost', width: 18 },
            { header: 'Profit', key: 'profit', width: 18 },
            { header: 'Timestamp', key: 'timestamp', width: 22 }
        ];

        // Style the header row
        const headerRow = projectsSheet.getRow(1);
        headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        headerRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF0066CC' }
        };
        headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
        headerRow.height = 25;

        // Create Summary Sheet
        const summarySheet = workbook.addWorksheet('Summary');
        
        // Add title
        summarySheet.mergeCells('A1:D1');
        const titleCell = summarySheet.getCell('A1');
        titleCell.value = 'Monthly Summary';
        titleCell.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
        titleCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF00AA00' }
        };
        titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
        summarySheet.getRow(1).height = 30;

        // Add overall metrics
        summarySheet.getCell('A3').value = 'Metric';
        summarySheet.getCell('B3').value = 'Value';
        summarySheet.getRow(3).font = { bold: true, color: { argb: 'FFFFFFFF' } };
        summarySheet.getRow(3).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF00AA00' }
        };
        summarySheet.getRow(3).alignment = { vertical: 'middle', horizontal: 'center' };
        summarySheet.getRow(3).height = 25;

        summarySheet.getCell('A4').value = 'Total Projects';
        summarySheet.getCell('B4').value = { formula: '=COUNTA(Projects!A2:A1000)' };
        
        summarySheet.getCell('A5').value = 'Total Revenue';
        summarySheet.getCell('B5').value = { formula: '=SUM(Projects!G2:G1000)' };
        summarySheet.getCell('B5').numFmt = '"LKR "#,##0.00';
        
        summarySheet.getCell('A6').value = 'Total Costs';
        summarySheet.getCell('B6').value = { formula: '=SUM(Projects!M2:M1000)' };
        summarySheet.getCell('B6').numFmt = '"LKR "#,##0.00';
        
        summarySheet.getCell('A7').value = 'Net Profit';
        summarySheet.getCell('B7').value = { formula: '=B5-B6' };
        summarySheet.getCell('B7').numFmt = '"LKR "#,##0.00';

        // Set column widths for metrics section
        summarySheet.getColumn(1).width = 30;
        summarySheet.getColumn(2).width = 20;

        // Add spacing
        summarySheet.getRow(8).height = 10;

        // Add project details table header
        summarySheet.getCell('A9').value = 'Project';
        summarySheet.getCell('B9').value = 'Total Cost';
        summarySheet.getCell('C9').value = 'Client Payment';
        summarySheet.getCell('D9').value = 'Profit';
        
        const detailsHeaderRow = summarySheet.getRow(9);
        detailsHeaderRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        detailsHeaderRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF0066CC' }
        };
        detailsHeaderRow.alignment = { vertical: 'middle', horizontal: 'center' };
        detailsHeaderRow.height = 25;
        
        // Add borders to header cells
        for (let col = 1; col <= 4; col++) {
            detailsHeaderRow.getCell(col).border = {
                top: { style: 'medium' },
                left: { style: 'medium' },
                bottom: { style: 'medium' },
                right: { style: 'medium' }
            };
        }

        // Set column widths for details table
        summarySheet.getColumn(3).width = 20;
        summarySheet.getColumn(4).width = 20;

        // Add formulas to pull data from Projects sheet (will be populated when projects are added)
        // This is a placeholder - actual data will be added via updateSummarySheet method

        // Save the file
        const filePath = this.getFilePath(month, year);
        await workbook.xlsx.writeFile(filePath);
        
        this.workbooks[`${year}-${month}`] = workbook;
        
        return [];
    }

    async updateSummarySheet(month, year) {
        const key = `${year}-${month}`;
        let workbook = this.workbooks[key];

        if (!workbook) {
            await this.loadOrCreateMonthFile(month, year);
            workbook = this.workbooks[key];
        }

        const projectsSheet = workbook.getWorksheet('Projects');
        const summarySheet = workbook.getWorksheet('Summary');

        // Clear existing project details (rows 10 onwards)
        const maxRow = summarySheet.rowCount;
        if (maxRow > 9) {
            summarySheet.spliceRows(10, maxRow - 9);
        }

        // Get all projects
        const projects = [];
        projectsSheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; // Skip header
            if (row.getCell(1).value) {
                projects.push({
                    name: row.getCell(1).value,
                    rowNumber: rowNumber
                });
            }
        });

        // Add project details with formulas
        let currentRow = 10;
        projects.forEach((project) => {
            const row = summarySheet.getRow(currentRow);
            row.getCell(1).value = { formula: `=Projects!A${project.rowNumber}` }; // Project Name
            row.getCell(2).value = { formula: `=Projects!M${project.rowNumber}` }; // Total Cost
            row.getCell(2).numFmt = '"LKR "#,##0.00';
            row.getCell(3).value = { formula: `=Projects!G${project.rowNumber}` }; // Client Payment
            row.getCell(3).numFmt = '"LKR "#,##0.00';
            row.getCell(4).value = { formula: `=Projects!N${project.rowNumber}` }; // Profit
            row.getCell(4).numFmt = '"LKR "#,##0.00';
            
            // Add borders to cells
            for (let col = 1; col <= 4; col++) {
                row.getCell(col).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            }
            
            // Add conditional formatting to profit cell
            const profitCell = row.getCell(4);
            profitCell.font = { bold: true };
            
            currentRow++;
        });

        // Add conditional formatting rules for profit column
        if (projects.length > 0) {
            const lastDataRow = 9 + projects.length;
            
            // Add green fill for positive profits
            summarySheet.addConditionalFormatting({
                ref: `D10:D${lastDataRow}`,
                rules: [
                    {
                        type: 'cellIs',
                        operator: 'greaterThan',
                        formulae: ['0'],
                        style: {
                            fill: {
                                type: 'pattern',
                                pattern: 'solid',
                                bgColor: { argb: 'FFC6EFCE' }
                            },
                            font: {
                                bold: true,
                                color: { argb: 'FF006100' }
                            }
                        }
                    },
                    {
                        type: 'cellIs',
                        operator: 'lessThan',
                        formulae: ['0'],
                        style: {
                            fill: {
                                type: 'pattern',
                                pattern: 'solid',
                                bgColor: { argb: 'FFFFC7CE' }
                            },
                            font: {
                                bold: true,
                                color: { argb: 'FF9C0006' }
                            }
                        }
                    }
                ]
            });
        }

        // Add chart instructions and data ranges if there are projects
        if (projects.length > 0) {
            const lastDataRow = 9 + projects.length;

            // Add spacing before chart section
            currentRow = lastDataRow + 2;

            // Add section header for charts
            summarySheet.mergeCells(`A${currentRow}:D${currentRow}`);
            const chartsHeader = summarySheet.getCell(`A${currentRow}`);
            chartsHeader.value = '📊 CHARTS';
            chartsHeader.font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
            chartsHeader.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF0066CC' }
            };
            chartsHeader.alignment = { vertical: 'middle', horizontal: 'center' };
            summarySheet.getRow(currentRow).height = 30;
            currentRow += 2;

            // Chart 1: Total Cost vs Client Payment
            summarySheet.mergeCells(`A${currentRow}:D${currentRow}`);
            const chart1Title = summarySheet.getCell(`A${currentRow}`);
            chart1Title.value = 'Chart 1: Total Cost vs Client Payment (Clustered Column)';
            chart1Title.font = { bold: true, size: 12, color: { argb: 'FF0066CC' } };
            chart1Title.alignment = { horizontal: 'left' };
            chart1Title.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE7F3FF' }
            };
            currentRow++;

            summarySheet.getCell(`A${currentRow}`).value = '• Select range: A9:C' + lastDataRow;
            summarySheet.getCell(`A${currentRow}`).font = { bold: true };
            currentRow++;
            summarySheet.getCell(`A${currentRow}`).value = '• Insert > Charts > Clustered Column Chart';
            currentRow++;
            summarySheet.getCell(`A${currentRow}`).value = '• Excel will automatically create the chart comparing Total Cost and Client Payment';
            currentRow += 2;

            // Chart 2: Profit per Project
            summarySheet.mergeCells(`A${currentRow}:D${currentRow}`);
            const chart2Title = summarySheet.getCell(`A${currentRow}`);
            chart2Title.value = 'Chart 2: Profit per Project (Bar Chart)';
            chart2Title.font = { bold: true, size: 12, color: { argb: 'FF0066CC' } };
            chart2Title.alignment = { horizontal: 'left' };
            chart2Title.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE7F3FF' }
            };
            currentRow++;

            summarySheet.getCell(`A${currentRow}`).value = '• Select Project column (A9:A' + lastDataRow + ')';
            summarySheet.getCell(`A${currentRow}`).font = { bold: true };
            currentRow++;
            summarySheet.getCell(`A${currentRow}`).value = '• Hold Ctrl/Cmd and select Profit column (D9:D' + lastDataRow + ')';
            currentRow++;
            summarySheet.getCell(`A${currentRow}`).value = '• Insert > Charts > Bar Chart';
            currentRow++;
            summarySheet.getCell(`A${currentRow}`).value = '• The chart will display profit for each project (green bars for positive, red for negative)';
            currentRow += 2;

            // Add note about automatic charts
            summarySheet.mergeCells(`A${currentRow}:D${currentRow}`);
            const noteCell = summarySheet.getCell(`A${currentRow}`);
            noteCell.value = '💡 TIP: Both charts will update automatically when you add or modify projects!';
            noteCell.font = { italic: true, color: { argb: 'FF666666' } };
            noteCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFFFE0' }
            };
            noteCell.alignment = { horizontal: 'center' };
        }

        // Save the updated workbook
        const filePath = this.getFilePath(month, year);
        await workbook.xlsx.writeFile(filePath);
    }

    async readProjectsFromWorkbook(workbook) {
        const projectsSheet = workbook.getWorksheet('Projects');
        const projects = [];

        projectsSheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; // Skip header

            const project = {
                projectName: row.getCell(1).value,
                numEngineers: row.getCell(2).value,
                engineerSalary: row.getCell(3).value,
                ceVisitCharge: row.getCell(4).value,
                visitsPerMonth: row.getCell(5).value,
                transportCost: row.getCell(6).value,
                clientPayment: row.getCell(7).value,
                overheadAllocation: row.getCell(8).value,
                engineerCost: row.getCell(9).value,
                ceVisitCost: row.getCell(10).value,
                directCost: row.getCell(11).value,
                overheadCost: row.getCell(12).value,
                totalCost: row.getCell(13).value,
                profit: row.getCell(14).value,
                timestamp: row.getCell(15).value
            };

            if (project.projectName) {
                projects.push(project);
            }
        });

        return projects;
    }

    async saveProject(month, year, projectData) {
        const key = `${year}-${month}`;
        let workbook = this.workbooks[key];

        if (!workbook) {
            await this.loadOrCreateMonthFile(month, year);
            workbook = this.workbooks[key];
        }

        const projectsSheet = workbook.getWorksheet('Projects');
        
        // Get the next row number
        const nextRowNumber = projectsSheet.rowCount + 1;
        
        // Add the project row with formulas
        const newRow = projectsSheet.addRow({
            projectName: projectData.projectName,
            numEngineers: projectData.numEngineers,
            engineerSalary: projectData.engineerSalary,
            ceVisitCharge: projectData.ceVisitCharge,
            visitsPerMonth: projectData.visitsPerMonth,
            transportCost: projectData.transportCost,
            clientPayment: projectData.clientPayment,
            overheadAllocation: projectData.overheadAllocation,
            engineerCost: { formula: `=C${nextRowNumber}` },
            ceVisitCost: { formula: `=D${nextRowNumber}*E${nextRowNumber}` },
            directCost: { formula: `=I${nextRowNumber}+J${nextRowNumber}+F${nextRowNumber}` },
            overheadCost: { formula: `=K${nextRowNumber}*(H${nextRowNumber}/100)` },
            totalCost: { formula: `=K${nextRowNumber}+L${nextRowNumber}` },
            profit: { formula: `=G${nextRowNumber}-M${nextRowNumber}` },
            timestamp: new Date().toISOString()
        });

        // Format numeric cells
        for (let col = 2; col <= 14; col++) {
            const cell = newRow.getCell(col);
            if (col >= 3 && col !== 5 && col !== 8 && col !== 2) { // Format as currency except for counts and percentages
                cell.numFmt = '"LKR "#,##0.00';
            } else if (col === 8) { // Format percentage
                cell.numFmt = '0.00"%"';
            }
        }

        // Save the file
        const filePath = this.getFilePath(month, year);
        await workbook.xlsx.writeFile(filePath);
        
        // Update summary sheet with new project data
        await this.updateSummarySheet(month, year);
        
        return projectData;
    }

    async deleteProject(month, year, projectName) {
        const key = `${year}-${month}`;
        let workbook = this.workbooks[key];

        if (!workbook) {
            await this.loadOrCreateMonthFile(month, year);
            workbook = this.workbooks[key];
        }

        const projectsSheet = workbook.getWorksheet('Projects');
        let rowToDelete = null;

        projectsSheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; // Skip header
            if (row.getCell(1).value === projectName) {
                rowToDelete = rowNumber;
            }
        });

        if (rowToDelete) {
            projectsSheet.spliceRows(rowToDelete, 1);
            const filePath = this.getFilePath(month, year);
            await workbook.xlsx.writeFile(filePath);
            
            // Update summary sheet after deletion
            await this.updateSummarySheet(month, year);
        }
    }

    async getProjects(month, year) {
        const key = `${year}-${month}`;
        let workbook = this.workbooks[key];

        if (!workbook) {
            return await this.loadOrCreateMonthFile(month, year);
        }

        return await this.readProjectsFromWorkbook(workbook);
    }

    async exportFile(month, year, destinationPath) {
        const sourcePath = this.getFilePath(month, year);
        
        if (!fs.existsSync(sourcePath)) {
            throw new Error('Source file does not exist');
        }

        fs.copyFileSync(sourcePath, destinationPath);
    }

    async createBackup(month, year) {
        const sourcePath = this.getFilePath(month, year);
        
        if (!fs.existsSync(sourcePath)) {
            throw new Error('Source file does not exist');
        }

        const backupPath = this.getBackupPath(month, year);
        const backupDir = path.dirname(backupPath);

        if (!fs.existsSync(backupDir)) {
            fs.mkdirSync(backupDir, { recursive: true });
        }

        fs.copyFileSync(sourcePath, backupPath);
        
        return backupPath;
    }
}

module.exports = ExcelService;