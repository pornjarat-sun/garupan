function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ระบบบริหารจัดการข้อมูลครุภัณฑ์')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 1. ฟังก์ชันค้นหาข้อมูลครุภัณฑ์
function searchEquipment(keyword, statusFilter) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetNames = ["ครุภัณฑ์ต่ำกว่าเกณฑ์", "ครุภัณฑ์ปกติ"];
    let results = [];

    const rawKw = keyword ? String(keyword).toLowerCase().trim() : '';
    const kwTokens = rawKw ? rawKw.split(/\s+/).filter(t => t.length > 0) : [];

    sheetNames.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;

      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow <= 1) return;

      const fullData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      const headers = fullData[0].map(h => String(h).trim().toUpperCase());

      const colNameIdx   = headers.indexOf("EQUIPMENT_NAME");
      const colBrandIdx  = headers.indexOf("BRAND_NAME");
      const colCodeIdx   = headers.indexOf("FOR_SORT") !== -1 ? headers.indexOf("FOR_SORT") : headers.indexOf("EQUIPMENT_ID");
      const colRoomIdx   = headers.indexOf("ROOM_CURRENT") !== -1 && headers.indexOf("ROOM_CURRENT") > 0 ? headers.indexOf("ROOM_CURRENT") : headers.indexOf("ROOM");
      const colRemarkIdx = headers.indexOf("REMARK");

      const statusIndices = [];
      ["STATUS_1", "STATUS_6", "STATUS_8", "STATUS_2", "STATUS_3", "STATUS_4", "STATUS_5"].forEach(name => {
        const idx = headers.indexOf(name);
        if (idx !== -1) statusIndices.push(idx);
      });
      if (statusIndices.length === 0) {
        for (let c = 46; c <= 52; c++) statusIndices.push(c);
      }

      const dataRows = fullData.slice(1);

      for (let i = 0; i < dataRows.length; i++) {
        const row = dataRows[i];

        let finalStatus = '';
        for (let idx of statusIndices) {
          if (row[idx] && String(row[idx]).trim() !== '') {
            finalStatus = String(row[idx]).trim();
            break;
          }
        }
        if (!finalStatus) finalStatus = 'ใช้การได้';

        if (statusFilter && finalStatus !== statusFilter) continue;

        const name  = colNameIdx !== -1 && row[colNameIdx] ? String(row[colNameIdx]).trim() : '';
        const brand = colBrandIdx !== -1 && row[colBrandIdx] ? String(row[colBrandIdx]).trim() : '';

        if (!name || name === 'รายละเอียดครุภัณฑ์' || name === 'EQUIPMENT_NAME') continue;

        let code = colCodeIdx !== -1 && row[colCodeIdx] ? String(row[colCodeIdx]).trim() : '';
        if (!code) code = row[3] ? String(row[3]).trim() : '-';

        let room = colRoomIdx !== -1 && row[colRoomIdx] ? String(row[colRoomIdx]).trim() : '';
        if (!room) room = row[35] ? String(row[35]).trim() : '-';

        const remark = colRemarkIdx !== -1 && row[colRemarkIdx] ? String(row[colRemarkIdx]).trim() : '-';

        const searchableText = `${name} ${brand} ${code} ${room} ${remark}`.toLowerCase();
        let isMatch = true;

        if (kwTokens.length > 0) {
          const directMatch = searchableText.includes(rawKw);
          const tokenMatch = kwTokens.every(token => searchableText.includes(token));
          let subMatch = false;
          if (rawKw.includes("เครื่องสำรองไฟ")) {
            subMatch = searchableText.includes("สำรองไฟ") || searchableText.includes("เครื่องสำรอง");
          }
          isMatch = directMatch || tokenMatch || subMatch;
        }

        if (isMatch) {
          results.push({
            sheetType: sheetName,
            rowIndex: i + 2, // เก็บบรรทัดใน Google Sheets เพื่อใช้อัปเดต
            code: code,
            name: name,
            brand: brand || '-',
            room: room,
            status: finalStatus,
            remark: remark
          });
        }
      }
    });

    return results;
  } catch (error) {
    Logger.log("Error in searchEquipment: " + error.toString());
    return [];
  }
}

// 2. ฟังก์ชันอัปเดตสถานที่และสถานะครุภัณฑ์
function updateEquipment(sheetName, rowIndex, newRoom, newStatus, newRemark) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'ไม่พบชีตข้อมูล' };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim().toUpperCase());

    // ตำแหน่งคอลัมน์สถานที่ (ROOM_CURRENT)
    let colRoomCurrentIdx = headers.indexOf("ROOM_CURRENT") + 1;
    if (colRoomCurrentIdx === 0) colRoomCurrentIdx = 54; // คอลัมน์ BB (Index 53 + 1)

    // คอลัมน์สถานะ (AU ถึง BA)
    const statusCols = {
      "ใช้การได้": headers.indexOf("STATUS_1") + 1 || 47,   // AU
      "ตัดจำหน่าย": headers.indexOf("STATUS_6") + 1 || 48,  // AV
      "ไม่ใช้งาน": headers.indexOf("STATUS_8") + 1 || 49,   // AW
      "ชำรุด": headers.indexOf("STATUS_2") + 1 || 50,       // AX
      "เสื่อมสภาพ": headers.indexOf("STATUS_3") + 1 || 51,   // AY
      "โอนให้": headers.indexOf("STATUS_4") + 1 || 52,     // AZ
      "สูญหาย": headers.indexOf("STATUS_5") + 1 || 53      // BA
    };

    // ล้างค่าสถานะเดิมทั้งหมดในบรรทัดนั้น (AU ถึง BA)
    for (let statusKey in statusCols) {
      const colIdx = statusCols[statusKey];
      if (colIdx > 0) {
        sheet.getRange(rowIndex, colIdx).setValue("");
      }
    }

    // เขียนค่าสถานะใหม่ลงช่องที่ตรงกัน
    if (statusCols[newStatus]) {
      sheet.getRange(rowIndex, statusCols[newStatus]).setValue(newStatus);
    }

    // อัปเดตสถานที่ใช้งานปัจจุบัน (ROOM_CURRENT)
    if (colRoomCurrentIdx > 0) {
      sheet.getRange(rowIndex, colRoomCurrentIdx).setValue(newRoom);
    }

    // อัปเดตหมายเหตุ (REMARK) ถ้ามี
    const colRemarkIdx = headers.indexOf("REMARK") + 1 || 56;
    if (colRemarkIdx > 0 && newRemark !== undefined) {
      sheet.getRange(rowIndex, colRemarkIdx).setValue(newRemark);
    }

    return { success: true, message: 'บันทึกข้อมูลเรียบร้อยแล้ว' };
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}
