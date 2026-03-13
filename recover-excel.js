// Recovery Script for Corrupted Excel Files
// This script will fix existing Excel files that have wrong calculations

const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const readline = require('readline');

const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

async function fixExcelFile(filePath) {
    console.log(`\n🔧 Fixing file: ${filePath}`);
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const projectsSheet = workbook.getWorksheet('Projects');
    if (!projectsSheet) {
        console.log('❌ No Projects sheet found!');
        return false;
    }
    
    let fixedCount = 0;
    let rowNumber = 2; // Start from row 2 (after header)
    
    projectsSheet.eachRow((row, rowNum) => {
        if (rowNum === 1) return; // Skip header
        
        const projectName = row.getCell(1).value;
        if (!projectName) return; // Skip empty rows
        
        // Get values
        const numEngineers = row.getCell(2).value;
        const engineerSalary = row.getCell(3).value;
        
        // Fix the engineer cost formula (Column I = 9)
        const currentFormula = row.getCell(9).value;
        const correctFormula = { formula: `=B${rowNum}*C${rowNum}` };
        
        // Check if formula is wrong
        if (currentFormula && currentFormula.formula === `=C${rowNum}`) {
            console.log(`  ⚠️  Row ${rowNum} (${projectName}): Wrong formula detected`);
            console.log(`     Before: =C${rowNum} (just salary)`);
            console.log(`     After:  =B${rowNum}*C${rowNum} (engineers × salary)`);
            
            row.getCell(9).value = correctFormula;
            fixedCount++;
        }
    });
    
    if (fixedCount > 0) {
        // Create backup first
        const backupPath = filePath.replace('.xlsx', '_BACKUP_' + Date.now() + '.xlsx');
        fs.copyFileSync(filePath, backupPath);
        console.log(`\n  📦 Backup created: ${backupPath}`);
        
        // Save fixed file
        await workbook.xlsx.writeFile(filePath);
        console.log(`  ✅ Fixed ${fixedCount} row(s) in ${path.basename(filePath)}`);
        return true;
    } else {
        console.log(`  ℹ️  No issues found in ${path.basename(filePath)}`);
        return false;
    }
}

async function findAndFixFiles(directory) {
    console.log(`\n🔍 Searching for Excel files in: ${directory}\n`);
    
    if (!fs.existsSync(directory)) {
        console.log('❌ Directory does not exist!');
        return;
    }
    
    const files = fs.readdirSync(directory);
    const excelFiles = files.filter(f => f.endsWith('.xlsx') && !f.includes('BACKUP'));
    
    if (excelFiles.length === 0) {
        console.log('❌ No Excel files found!');
        return;
    }
    
    console.log(`Found ${excelFiles.length} Excel file(s):\n`);
    excelFiles.forEach((f, i) => {
        console.log(`  ${i + 1}. ${f}`);
    });
    
    console.log('\n───────────────────────────────────────────────────');
    
    let totalFixed = 0;
    for (const file of excelFiles) {
        const filePath = path.join(directory, file);
        try {
            const wasFixed = await fixExcelFile(filePath);
            if (wasFixed) totalFixed++;
        } catch (error) {
            console.log(`\n❌ Error fixing ${file}:`, error.message);
        }
    }
    
    console.log('\n───────────────────────────────────────────────────');
    console.log(`\n✅ Recovery complete!`);
    console.log(`   Files processed: ${excelFiles.length}`);
    console.log(`   Files fixed: ${totalFixed}`);
    console.log(`   Files unchanged: ${excelFiles.length - totalFixed}`);
    
    if (totalFixed > 0) {
        console.log(`\n⚠️  IMPORTANT: Your Excel files have been fixed!`);
        console.log(`   - Original files were backed up with _BACKUP_ in the filename`);
        console.log(`   - Open the files in Excel to see the corrected calculations`);
        console.log(`   - Formulas will now calculate correctly`);
    }
}

function question(query) {
    return new Promise(resolve => rl.question(query, resolve));
}

async function main() {
    console.log('═══════════════════════════════════════════════════');
    console.log('   Excel Recovery Tool - Fix Corrupted Formulas   ');
    console.log('═══════════════════════════════════════════════════\n');
    
    console.log('This tool will fix Excel files with wrong engineer cost formulas.\n');
    
    // Get directory from user
    const defaultDir = process.platform === 'win32' 
        ? 'C:\\Users\\' + require('os').userInfo().username + '\\AppData\\Roaming\\profit-dashboard\\excel-files'
        : process.platform === 'darwin'
        ? require('os').homedir() + '/Library/Application Support/profit-dashboard/excel-files'
        : require('os').homedir() + '/.config/profit-dashboard/excel-files';
    
    console.log(`Default location: ${defaultDir}\n`);
    
    const answer = await question('Enter directory path (or press Enter for default): ');
    const directory = answer.trim() || defaultDir;
    
    await findAndFixFiles(directory);
    
    rl.close();
}

// Run the recovery
main().catch(error => {
    console.error('\n❌ Fatal error:', error);
    process.exit(1);
});
