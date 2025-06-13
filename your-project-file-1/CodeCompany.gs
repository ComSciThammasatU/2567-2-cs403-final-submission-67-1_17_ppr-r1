function processLatestFile() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("ยืนยันการตอบรับนักศึกษาสหกิจศึกษา");
  const folderId = "1wQ1ZwGADlj6XRmyl13jPucVMTOo8Sh_a"; // โฟลเดอร์ Google Drive ที่ใช้เก็บไฟล์
  const folder = DriveApp.getFolderById(folderId);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const filePath = sheet.getRange(lastRow, 6).getValue(); // คอลัมน์ F
  const fileName = getCleanFileName(filePath);

  const targetCell = sheet.getRange(lastRow, 9); // คอลัมน์ I
  insertLink(folder, fileName, targetCell);
}

function getCleanFileName(filePath) {
  if (!filePath) return "";
  const parts = filePath.toString().split("_Files_/");
  return parts.length > 1 ? parts[1].trim() : filePath.trim();
}

function insertLink(folder, fileName, targetCell) {
  if (!fileName) {
    targetCell.setValue("");
    return;
  }

  const files = folder.getFilesByName(fileName);
  if (files.hasNext()) {
    const file = files.next();
    const fileUrl = `https://drive.google.com/file/d/${file.getId()}/view`;
    targetCell.setFormula(`=HYPERLINK("${fileUrl}", "${fileName}")`);
  } else {
    targetCell.setValue("ไม่พบไฟล์ในโฟลเดอร์");
  }
}


// 🔹 ตัดชื่อไฟล์ให้สะอาดขึ้น
function getCleanFileName(filePath) {
  const parts = filePath.split("_Files_/");
  return parts.length > 1 ? parts[1] : filePath;
}



function updateInitialStatus() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ลงทะเบียนสถานประกอบการ");
  const lastRow = sheet.getLastRow();

  for (let i = 2; i <= lastRow; i++) {
    const a = sheet.getRange(i, 1).getValue()?.toString().trim();    // คอลัมน์ A
    const q = sheet.getRange(i, 17).getValue()?.toString().trim();   // คอลัมน์ Q
    const r = sheet.getRange(i, 18).getValue()?.toString().trim();   // คอลัมน์ R
    const uCell = sheet.getRange(i, 21);  // คอลัมน์ U
    const tCell = sheet.getRange(i, 20);  // คอลัมน์ T

    const u = uCell.getValue()?.toString().trim();
    const t = tCell.getValue()?.toString().trim();

    // ถ้ายังไม่มีสถานะภาษาอังกฤษ และ A มีข้อมูล → ตั้ง Pending Approval
    if (a && (!u || u === "")) {
      uCell.setValue("Pending Approval");
    }

    // ถ้ามี Q และ R แล้วสถานะยังเป็น Pending Approval → เปลี่ยนเป็น Under Review
    const updatedU = uCell.getValue()?.toString().trim();
    if (a && q && r && updatedU === "Pending Approval") {
      uCell.setValue("Under Review");
    }

    // ถ้า U เป็น Under Review → T ต้องเป็น รอการอนุมัติ (ถ้ายังว่างหรือยังไม่ตรง)
    const refreshedU = uCell.getValue()?.toString().trim();
    if (refreshedU === "Under Review" && t !== "รอการอนุมัติ") {
      tCell.setValue("รอการอนุมัติ");
    }
  }
}


function onEdit(e) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const cell = e.range;

  // 🟢 กรณี: ชีต "นักศึกษาก่อนปฏิบัติสหกิจ"
  if (sheetName === "ลงทะเบียนสถานประกอบการ") {
    const row = cell.getRow();
    const col = cell.getColumn();
    const editedValue = cell.getValue();

    // สำหรับ ชื่อ / Email 
    const codeColumn = 17; // Q
    const nameColumn = 18; // R

    const importSheet = e.source.getSheetByName("รายชื่ออาจารย์");
    const importData = importSheet.getRange(2, 1, importSheet.getLastRow() - 1, 2).getValues(); // A:B

    if (col === codeColumn && editedValue) {
      const found = importData.find(r => r[0] === editedValue);
      sheet.getRange(row, nameColumn).setValue(found ? found[1] : "");
    }

    if (col === nameColumn && editedValue) {
      const found = importData.find(r => r[1] === editedValue);
      sheet.getRange(row, codeColumn).setValue(found ? found[0] : "");
    }
  }
}


