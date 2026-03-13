// Test Script for Excel Service
// Run this with: node test-excel.js

const ExcelService = require('./excelService');
const path = require('path');
const fs = require('fs');

async function testExcelService() {
    console.log('🧪 Testing Excel Service...\n');
    
    // Create test directory
    const testDir = path.join(__dirname, 'test-output');
    if (!fs.existsSync(testDir)) {
        fs.mkdirSync(testDir);
    }
    
    const excelService = new ExcelService(testDir);
    
    try {
        // Test 1: Create new file
        console.log('Test 1: Creating new Excel file...');
        const projects = await excelService.loadOrCreateMonthFile('January', '2024');
        console.log('✓ File created successfully');
        console.log(`  Projects loaded: ${projects.length}`);
        
        // Test 2: Add a project
        console.log('\nTest 2: Adding a project...');
        const testProject = {
            projectName: 'Test Project Alpha',
            numEngineers: 3,
            engineerSalary: 75000,
            ceVisitCharge: 5000,
            visitsPerMonth: 4,
            transportCost: 10000,
            clientPayment: 300000,
            overheadAllocation: 15
        };
        
        await excelService.saveProject('January', '2024', testProject);
        console.log('✓ Project added successfully');
        
        // Test 3: Verify the file exists
        console.log('\nTest 3: Verifying file...');
        const filePath = excelService.getFilePath('January', '2024');
        if (fs.existsSync(filePath)) {
            const stats = fs.statSync(filePath);
            console.log('✓ File exists');
            console.log(`  Path: ${filePath}`);
            console.log(`  Size: ${stats.size} bytes`);
        } else {
            console.log('✗ File does not exist!');
        }
        
        // Test 4: Load and verify projects
        console.log('\nTest 4: Loading projects...');
        const loadedProjects = await excelService.getProjects('January', '2024');
        console.log(`✓ Loaded ${loadedProjects.length} project(s)`);
        
        if (loadedProjects.length > 0) {
            console.log('\nProject Details:');
            loadedProjects.forEach((p, i) => {
                console.log(`\n  Project ${i + 1}: ${p.projectName}`);
                console.log(`    Engineers: ${p.numEngineers}`);
                console.log(`    Salary: ${p.engineerSalary}`);
                console.log(`    Engineer Cost: ${p.engineerCost}`);
                console.log(`    Total Cost: ${p.totalCost}`);
                console.log(`    Client Payment: ${p.clientPayment}`);
                console.log(`    Profit: ${p.profit}`);
            });
        }
        
        // Test 5: Add another project
        console.log('\n\nTest 5: Adding second project...');
        const testProject2 = {
            projectName: 'Test Project Beta',
            numEngineers: 2,
            engineerSalary: 65000,
            ceVisitCharge: 4000,
            visitsPerMonth: 3,
            transportCost: 8000,
            clientPayment: 180000,
            overheadAllocation: 12
        };
        
        await excelService.saveProject('January', '2024', testProject2);
        console.log('✓ Second project added');
        
        // Final verification
        console.log('\n\nFinal Verification:');
        const finalProjects = await excelService.getProjects('January', '2024');
        console.log(`✓ Total projects in file: ${finalProjects.length}`);
        
        // Calculate expected values for first project
        const expectedEngineerCost = testProject.numEngineers * testProject.engineerSalary;
        const expectedCECost = testProject.ceVisitCharge * testProject.visitsPerMonth;
        const expectedDirectCost = expectedEngineerCost + expectedCECost + testProject.transportCost;
        const expectedOverhead = expectedDirectCost * (testProject.overheadAllocation / 100);
        const expectedTotalCost = expectedDirectCost + expectedOverhead;
        const expectedProfit = testProject.clientPayment - expectedTotalCost;
        
        console.log('\n📊 Calculation Verification (Project 1):');
        console.log(`  Expected Engineer Cost: ${expectedEngineerCost} (${testProject.numEngineers} × ${testProject.engineerSalary})`);
        console.log(`  Actual Engineer Cost: ${finalProjects[0].engineerCost}`);
        console.log(`  Match: ${finalProjects[0].engineerCost === expectedEngineerCost ? '✓' : '✗'}`);
        
        console.log(`\n  Expected Total Cost: ${expectedTotalCost.toFixed(2)}`);
        console.log(`  Actual Total Cost: ${finalProjects[0].totalCost}`);
        
        console.log(`\n  Expected Profit: ${expectedProfit.toFixed(2)}`);
        console.log(`  Actual Profit: ${finalProjects[0].profit}`);
        
        console.log('\n\n✅ All tests completed!');
        console.log(`📁 Test files created in: ${testDir}`);
        console.log('   Open the Excel file to verify formulas are working correctly.');
        
    } catch (error) {
        console.error('\n❌ Test failed:', error.message);
        console.error(error.stack);
        process.exit(1);
    }
}

// Run the tests
testExcelService();
