
// Helper to parse "HH:MM" into minutes from midnight
const parseTimeStr = (timeStr) => {
  const [h, m] = timeStr.split(':').map(Number);
  return h * 60 + m;
};

// Helper to get minutes from midnight from a Date object or Date string
const getTimeInMinutes = (cellValue) => {
  if (!cellValue) return null;
  
  // Case 1: Date Object (typically from ExcelJS reading a Date cell)
  if (cellValue instanceof Date) {
    // ExcelJS converts Excel serial dates (timezone-agnostic) to JS Dates assuming UTC.
    // e.g. "09:00" in Excel -> "09:00 UTC" in JS Date.
    // Using getUTCHours() correctly extracts the visual time intended in Excel.
    // Using getHours() would add browser timezone offset (e.g. +9h in JST), causing shifts.
    return cellValue.getUTCHours() * 60 + cellValue.getUTCMinutes();
  }

  // Case 2: String (e.g. "2025/12/15 09:00:00")
  if (typeof cellValue === 'string') {
    // Extract "HH:MM" pattern using regex to avoid Date parsing timezone ambiguity
    // Matches "09:00", " 9:00", "T09:00"
    const match = cellValue.match(/(?:^|\s|T)(\d{1,2}):(\d{2})/);
    if (match) {
      return parseInt(match[1], 10) * 60 + parseInt(match[2], 10);
    }
    
    // Fallback: Try Date parse (legacy behavior, but risky)
    const date = new Date(cellValue);
    if (!isNaN(date.getTime())) {
       return date.getHours() * 60 + date.getMinutes();
    }
  }
  
  return null;
};

window.processFiles = async (files, timeRange, pcsMaster, onProgress, options = {}) => {
  if (!files || files.length === 0) throw new Error("ファイルが選択されていません。");
  if (!pcsMaster) throw new Error("PCSマスタが読み込まれていません。");
  
  // Access libraries
  const ExcelJS = window.ExcelJS;
  if (!ExcelJS) throw new Error("ExcelJS library not loaded");

  const outputDebug = options.outputDebug || false; // New option

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
    // if (rowCount < 5) continue; // Allow small files? No, spec says 5+

    // Process Header (only from first file)
    if (!headerSet) {
      // Copy rows 1-4 verbatim
      for (let r = 1; r <= 4; r++) {
        const srcRow = sheet.getRow(r);
        const destRow = resultSheet.getRow(r);
        srcRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const destCell = destRow.getCell(colNumber);
          destCell.value = cell.value;
          // Deep copy style here too just in case
          destCell.style = JSON.parse(JSON.stringify(cell.style));
        });
        
        // Add Debug Header
        if (outputDebug && r === 4) {
            destRow.getCell(13).value = "検証用:回路数";
        }

        destRow.commit();
      }
      headerSet = true;
      currentRowIndex = 5;
    }

    // Process Data Rows (5+)
    for (let r = 5; r <= rowCount; r++) {
      const srcRow = sheet.getRow(r);
      if (!srcRow.hasValues) continue; // Skip empty rows

      // Copy row content first
      const destRow = resultSheet.getRow(currentRowIndex);
      
      const cellCount = srcRow.cellCount;
      for (let c = 1; c <= Math.max(cellCount, 12); c++) {
        const srcCell = srcRow.getCell(c);
        const destCell = destRow.getCell(c);
        destCell.value = srcCell.value;
        // Copy style safely to avoid reference sharing issues
        destCell.style = JSON.parse(JSON.stringify(srcCell.style)); 
      }

      stats.totalRows++;
      
      // LOGIC STARTS HERE
      // 1. PCS Extract (Column C = 3)
      const pcsCellVal = destRow.getCell(3).text; 
      if (pcsCellVal) {
        // "192.168.1.201 変1 A-1/PCS1-3-4" -> "PCS 1-3-4"
        const parts = pcsCellVal.split('/');
        if (parts.length > 1) {
          let pcsKeyRaw = parts[parts.length - 1].trim(); // "PCS1-3-4"
          // Normalize: insert space after PCS if missing
          const pcsKey = pcsKeyRaw.replace(/^PCS(\d)/i, 'PCS $1');
          
          let circuitCount = pcsMap.get(pcsKey);
          
          if (!circuitCount) {
             stats.unknownPCS++;
             unknownPCSList.add(pcsKey);
             circuitCount = 8; // Default to checked everything if unknown
          }

          // Debug Output
          if (outputDebug) {
             destRow.getCell(13).value = circuitCount;
          }

          // 2. Time Check (Column D = 4)
          const timeVal = destRow.getCell(4).value;
          const timeMin = getTimeInMinutes(timeVal);
          
          if (isTimeTarget(timeMin)) {
             stats.targetRows++;
             // 3. 0A Check (E=5 to L=12)
             const rangeEnd = (circuitCount === 7) ? 11 : 12;

             for (let c = 5; c <= rangeEnd; c++) {
               const cell = destRow.getCell(c);
               const val = cell.value;
               
               // Check for 0 (numeric 0 or string "0", but not "0A" or other strings)
               let isZero = false;
               if (val === 0) isZero = true;
               
               // Strict string check: must be "0", "0.0", etc. 
               // parseFloat parses "0A" as 0, which is dangerous if unit is included unexpectedly.
               // However, user data is likely clean. But let's be safe.
               if (typeof val === 'string') {
                   const trimmed = val.trim();
                   // Check if it looks like a number
                   if (!isNaN(Number(trimmed)) && Number(trimmed) === 0 && trimmed !== '') {
                       isZero = true;
                   }
               }

               // Ignore null or "-"
               if (val === null || val === undefined || val === '-' || val === '') isZero = false;

               if (isZero) {
                 // RED HIGHLIGHT
                 cell.fill = {
                   type: 'pattern',
                   pattern: 'solid',
                   fgColor: { argb: 'FFFF9999' } // Light Red
                 };
                 stats.highlightedCells++;
               }
             }
          }
        }
      }

      destRow.commit();
      currentRowIndex++;

      // Yield to event loop every 500 rows
      if (r % 500 === 0) {
        await new Promise(resolve => setTimeout(resolve, 0));
        onProgress({ 
            phase: 'processing', 
            current: r, 
            total: rowCount, 
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
    unknownPCSList: Array.from(unknownPCSList).slice(0, 10)
  };
};
