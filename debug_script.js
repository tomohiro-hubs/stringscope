
const ExcelJS = require('exceljs');

async function debugFile() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('debug_target.xlsx');
    const sheet = workbook.getWorksheet(1);
    
    // User mentioned "L125". 
    // Let's check Row 125, and also search for "PCS 2-2-3" to be sure.
    
    console.log("--- Inspecting Row 125 ---");
    const row125 = sheet.getRow(125);
    console.log(`Row 125 has values: ${row125.hasValues}`);
    console.log(`Col C (3) [Name]: ${JSON.stringify(row125.getCell(3).value)}`);
    console.log(`Col D (4) [Time]: ${JSON.stringify(row125.getCell(4).value)}`);
    console.log(`Col L (12) [PV8]: ${JSON.stringify(row125.getCell(12).value)}`);
    console.log(`Col L Type: ${typeof row125.getCell(12).value}`);
    
    // Also search for PCS 2-2-3 rows to see times
    console.log("\n--- Searching for PCS 2-2-3 ---");
    sheet.eachRow((row, rowNumber) => {
        const name = row.getCell(3).text; // Use text to safely get string
        if (name && name.includes("PCS 2-2-3")) { // Loose match
             const timeVal = row.getCell(4).value;
             const valL = row.getCell(12).value;
             if (Math.abs(valL) < 0.1 || valL === '0') { // Close to 0
                 console.log(`Row ${rowNumber}: Time=${timeVal}, PV8=${valL} (Type: ${typeof valL})`);
             }
        }
    });
}

debugFile().catch(console.error);
