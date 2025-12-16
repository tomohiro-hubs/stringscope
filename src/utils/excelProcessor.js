import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

// Helper to parse "HH:MM" into minutes from midnight
const parseTimeStr = (timeStr) => {
  const [h, m] = timeStr.split(':').map(Number);
  return h * 60 + m;
};

// Helper to get minutes from midnight from a Date object or Date string
const getTimeInMinutes = (cellValue) => {
  if (!cellValue) return null;
  let date;
  if (cellValue instanceof Date) {
    date = cellValue;
  } else {
    // Try parsing string
    date = new Date(cellValue);
    if (isNaN(date.getTime())) return null;
  }
  return date.getHours() * 60 + date.getMinutes();
};

export const processFiles = async (files, timeRange, pcsMaster, onProgress) => {
  if (!files || files.length === 0) throw new Error("ファイルが選択されていません。");
  if (!pcsMaster) throw new Error("PCSマスタが読み込まれていません。");

  // Build Map for fast lookup: "PCS 1-1-1" -> 7
  const pcsMap = new Map();
  pcsMaster.pcsMaster.forEach(p => {
    pcsMap.set(p.pcsKey, p.circuitCount);
  });

  const startMin = parseTimeStr(timeRange.start);
  const endMin = parseTimeStr(timeRange.end);
  const isCrossDay = endMin < startMin; // e.g., 23:00 to 05:00

  const isTimeTarget = (minutes) => {
    if (minutes === null) return false;
    if (isCrossDay) {
      return minutes >= startMin || minutes <= endMin;
    }
    return minutes >= startMin && minutes <= endMin;
  };

  const workbook = new ExcelJS.Workbook();
  const resultSheet = workbook.addWorksheet('Result');
  
  let headerSet = false;
  let currentRowIndex = 1; // 1-based index for writing
  let stats = {
    totalRows: 0,
    targetRows: 0,
    highlightedCells: 0,
    unknownPCS: 0,
    filesProcessed: 0
  };

  const unknownPCSList = new Set();

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    onProgress({ 
        phase: 'reading', 
        current: i + 1, 
        total: files.length, 
        filename: file.name 
    });

    const buffer = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    
    // Assume first sheet
    const sheet = wb.worksheets[0];
    if (!sheet) continue;

    const rowCount = sheet.rowCount;
    // Need at least 5 rows (1-3 meta, 4 header, 5 data)
    if (rowCount < 5) continue;

    // Process Header (only from first file)
    if (!headerSet) {
      // Copy rows 1-4 verbatim
      for (let r = 1; r <= 4; r++) {
        const srcRow = sheet.getRow(r);
        const destRow = resultSheet.getRow(r);
        srcRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const destCell = destRow.getCell(colNumber);
          destCell.value = cell.value;
          destCell.style = cell.style;
        });
        destRow.commit();
      }
      headerSet = true;
      currentRowIndex = 5;
    }

    // Process Data Rows (5+)
    // Iterate manually to avoid memory issues with huge sheets? 
    // ExcelJS `eachRow` is fine for client-side usually if file isn't massive (hundreds of MBs).
    // Prompt implies "too large" for upload context, but user might have reasonably sized logs.
    
    for (let r = 5; r <= rowCount; r++) {
      const srcRow = sheet.getRow(r);
      if (!srcRow.hasValues) continue; // Skip empty rows

      // Copy row content first
      const destRow = resultSheet.getRow(currentRowIndex);
      
      // Copy cells A to max columns
      // Optimally, we only care up to L (12), but let's copy all
      const cellCount = srcRow.cellCount;
      for (let c = 1; c <= Math.max(cellCount, 12); c++) {
        const srcCell = srcRow.getCell(c);
        const destCell = destRow.getCell(c);
        destCell.value = srcCell.value;
        // Copy style (borders, etc) - optional, might be slow. 
        // Let's just copy value for speed, maybe font?
        // Requirements say "Keep format" roughly implied by "Combine".
        // Let's copy style but reset fill later if needed.
        destCell.style = srcCell.style; 
      }

      stats.totalRows++;
      
      // LOGIC STARTS HERE
      // 1. PCS Extract (Column C = 3)
      const pcsCellVal = destRow.getCell(3).text; // Use .text to get string representation
      if (pcsCellVal) {
        // "192.168.1.201 変1 A-1/PCS1-3-4" -> "PCS 1-3-4"
        const parts = pcsCellVal.split('/');
        if (parts.length > 1) {
          let pcsKeyRaw = parts[parts.length - 1].trim(); // "PCS1-3-4"
          // Normalize: insert space after PCS if missing
          // Regex: Replace "PCS" followed by digit with "PCS " + digit
          const pcsKey = pcsKeyRaw.replace(/^PCS(\d)/i, 'PCS $1');
          
          let circuitCount = pcsMap.get(pcsKey);
          
          if (!circuitCount) {
             stats.unknownPCS++;
             unknownPCSList.add(pcsKey);
             circuitCount = 8; // Default to checked everything if unknown, but warn
          }

          // 2. Time Check (Column D = 4)
          const timeVal = destRow.getCell(4).value;
          const timeMin = getTimeInMinutes(timeVal);
          
          if (isTimeTarget(timeMin)) {
             stats.targetRows++;
             // 3. 0A Check (E=5 to L=12)
             // PV1 (5) to PV7 (11) are always checked
             // PV8 (12) is checked only if circuitCount >= 8
             
             const rangeEnd = (circuitCount === 7) ? 11 : 12;

             for (let c = 5; c <= rangeEnd; c++) {
               const cell = destRow.getCell(c);
               const val = cell.value;
               
               // Check for 0 (numeric 0 or string "0")
               // Must strict check for 0, not null/undefined
               let isZero = false;
               if (val === 0) isZero = true;
               if (typeof val === 'string' && parseFloat(val) === 0) isZero = true;

               // Ignore null or "-"
               if (val === null || val === undefined || val === '-' || val === '') isZero = false;

               if (isZero) {
                 // RED HIGHLIGHT
                 // fgColor: { argb: 'FFFF0000' } is pure red. Use a softer pastel red for background?
                 // Requirement: "赤ハイライト" (Red Highlight). Usually means background.
                 cell.fill = {
                   type: 'pattern',
                   pattern: 'solid',
                   fgColor: { argb: 'FFFF9999' } // Light Red
                 };
                 stats.highlightedCells++;
               }
             }
          }
        } else {
             // Malformed PCS name line? Log?
        }
      }

      destRow.commit();
      currentRowIndex++;

      // Yield to event loop every 500 rows to keep UI responsive
      if (r % 500 === 0) {
        await new Promise(resolve => setTimeout(resolve, 0));
        onProgress({ 
            phase: 'processing', 
            current: r, 
            total: rowCount, // estimated
            filename: file.name 
        });
      }
    }
    stats.filesProcessed++;
  }

  // Generate Buffer
  const buffer = await workbook.xlsx.writeBuffer();
  
  return {
    buffer,
    stats,
    unknownPCSList: Array.from(unknownPCSList).slice(0, 10) // Top 10 warnings
  };
};
