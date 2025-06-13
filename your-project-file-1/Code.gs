function exportSheetToPDF() {
  // ดึงข้อมูลจาก Sheet ปัจจุบัน
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  // ระบุ range ที่จะนำไปใส่ PDF
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var sheetId = sheet.getSheetId();

  // กำหนด URL สำหรับ Export เป็น PDF
  var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheet.getId() + 
            '/export?format=pdf' +
            '&gid=' + sheetId +
            '&size=A4' +         // ขนาดกระดาษ A4
            '&portrait=true' +  // แนวนอน (false) / แนวตั้ง (true)
            '&fitw=true' +       // ปรับให้พอดีกับหน้า
            '&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5' +
            '&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false';

  // ดึง PDF เป็น Blob
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + token
    }
  });
  var blob = response.getBlob().setName(sheet.getName() + '.pdf');

  // **กำหนดโฟลเดอร์ที่ต้องการบันทึกไฟล์**
  var folderId = '1WT6BZ-LkoIWmDrTxVSV8sYrb9jIkGE2Y'; // ใส่ Folder ID ตรงนี้
  var folder = DriveApp.getFolderById(folderId);
  
  // บันทึกไฟล์ PDF ลงในโฟลเดอร์ที่กำหนด
  var pdfFile = folder.createFile(blob);
  Logger.log('PDF สร้างเสร็จแล้ว: ' + pdfFile.getUrl());

  // แสดงข้อความสำเร็จ
  SpreadsheetApp.getUi().alert('สร้างไฟล์ PDF สำเร็จแล้ว!\nไฟล์อยู่ในโฟลเดอร์: ' + pdfFile.getUrl());
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();  // สร้าง UI สำหรับ Google Sheets
  // เพิ่มเมนูใหม่บนแถบเมนู
  ui.createMenu('PDF')  // ชื่อเมนู
    .addItem('Export to PDF', 'exportSheetToPDF')  // เพิ่มรายการที่เรียกใช้งานฟังก์ชัน exportSheetToPDF
    .addToUi();  // เพิ่มเมนูลงใน UI
}

function sortByGPA() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const lastCol = 18; // คอลัมน์ A-R

  // ดึงข้อมูลจาก A2 ถึง R[lastRow]
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // กรองเฉพาะแถวที่มี GPA (H = index 7) เป็นตัวเลข
  const filteredData = data.filter(row => typeof row[7] === "number");

  // เรียงข้อมูลตาม GPA จากมากไปน้อย
  filteredData.sort((a, b) => b[7] - a[7]);

  // เคลียร์ข้อมูลเก่าจาก A2 ถึง R[lastRow]
  sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();

  // เขียนข้อมูลเรียงใหม่กลับลงชีต
  sheet.getRange(2, 1, filteredData.length, lastCol).setValues(filteredData);
}


function sortByCSGrade() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const lastCol = 18; // คอลัมน์ A-R

  // ดึงข้อมูลจาก A2 ถึง R[lastRow]
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // กรองเฉพาะแถวที่มีค่า CS Grade (K = index 10) เป็นตัวเลข
  const filteredData = data.filter(row => typeof row[10] === "number");

  // เรียงข้อมูลตาม CS Grade จากมากไปน้อย
  filteredData.sort((a, b) => b[10] - a[10]);

  // เคลียร์ข้อมูลเดิมจาก A2 ถึง R[lastRow]
  sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();

  // เขียนข้อมูลที่เรียงแล้วกลับลงชีต
  sheet.getRange(2, 1, filteredData.length, lastCol).setValues(filteredData);
}


function selectRows() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('กรุณากรอกจำนวนบรรทัดที่ต้องการเลือก (จากแถวที่ 2):');

  if (response.getSelectedButton() == ui.Button.OK) {
    var numRows = parseInt(response.getResponseText(), 10);
    
    if (!isNaN(numRows) && numRows > 0) {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      var lastRow = sheet.getLastRow();
      var lastCol = 19; // คอลัมน์ A-S (19)

      // ล้างสีพื้นทั้งหมด (แม้แต่แถวว่าง)
      var maxRows = sheet.getMaxRows(); // เพื่อให้แน่ใจว่าล้างครอบคลุม
      sheet.getRange(2, 1, maxRows - 1, lastCol).setBackground('#ffffff');
      sheet.getRange(2, 19, maxRows - 1, 1).setValue(''); // ล้างค่า "ผลการคัดเลือก" ทุกแถว

      // ดึงข้อมูลตั้งแต่แถวที่ 2
      var data = sheet.getRange(2, 1, maxRows - 1, lastCol).getValues();

      var passCount = 0;

      for (var i = 0; i < data.length; i++) {
        var rowIndex = i + 2; // ระบุแถวที่ต้องการปรับเปลี่ยน
        var row = data[i];
        var isRowEmpty = row.every(function(cell) { return cell === "" || cell === null; });

        var bgRange = sheet.getRange(rowIndex, 1, 1, lastCol); // กำหนดช่วงที่ต้องการคลุม
        var resultCell = sheet.getRange(rowIndex, 19); // คอลัมน์ S

        if (passCount < numRows) {
          // คลุมแถวตามจำนวนที่เลือก
          bgRange.setBackground('#FFFF00'); // ตั้งค่าพื้นหลังเป็นสีเหลือง
          resultCell.setValue('PASS');
          passCount++;
        } else if (!isRowEmpty) {
          // แถวที่เหลือจะเป็น "FAIL" ถ้ามีข้อมูล
          resultCell.setValue('FAIL');
        }
      }
    } else {
      ui.alert('กรุณากรอกจำนวนที่เป็นบวก');
    }
  }
}
function onEdit(e) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const cell = e.range;

  // ชีต: นักศึกษาก่อนปฏิบัติสหกิจ
  if (sheetName === "นักศึกษาก่อนปฏิบัติสหกิจ") {
    const row = cell.getRow();
    const col = cell.getColumn();
    const editedValue = cell.getValue();

    const codeColumn = 7; // G
    const nameColumn = 8; // H

    const importSheet = e.source.getSheetByName("importCompany");
    const importData = importSheet.getRange(2, 1, importSheet.getLastRow() - 1, 2).getValues(); // A:B

    if (col === codeColumn && editedValue) {
      const found = importData.find(r => r[0] === editedValue);
      sheet.getRange(row, nameColumn).setValue(found ? found[1] : "");
    }

    if (col === nameColumn && editedValue) {
      const found = importData.find(r => r[1] === editedValue);
      sheet.getRange(row, codeColumn).setValue(found ? found[0] : "");
    }

    // ตำแหน่งงาน I-J
    const posCodeCol = 9;  // I
    const posNameCol = 10; // J

    const positionSheet = e.source.getSheetByName("ชื่อตำแหน่งงาน");
    const positionData = positionSheet.getRange(2, 1, positionSheet.getLastRow() - 1, 2).getValues(); // A:B

    if (col === posCodeCol && editedValue) {
      const found = positionData.find(r => r[0] === editedValue);
      sheet.getRange(row, posNameCol).setValue(found ? found[1] : "");
    }

    if (col === posNameCol && editedValue) {
      const found = positionData.find(r => r[1] === editedValue);
      sheet.getRange(row, posCodeCol).setValue(found ? found[0] : "");
    }
  }

  // เฉพาะชีต "การยืนยันสิทธิ์-ผ่านช่องทางสื่อสาร"
  if (sheetName === "การยืนยันสิทธิ์-ผ่านช่องทางสื่อสาร") {
    const startRow = cell.getRow();
    const startCol = cell.getColumn();
    const numRows = cell.getNumRows();
    const numCols = cell.getNumColumns();

    // เฉพาะเมื่อมีการวางข้อมูลในคอลัมน์ I เป็นต้นไป
    if (startCol >= 9) {
      const editedValues = cell.getValues().flat().map(v => v?.toString().trim()).filter(v => v);
      if (editedValues.length === 0) return;

      const studentData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues(); // A2:G
      const idColumn = 5;     // คอลัมน์ E
      const statusColumn = 7; // คอลัมน์ G
      const studentIdMap = new Map();

      // สร้าง Map สำหรับ lookup เร็วขึ้น
      studentData.forEach((row, idx) => {
        const id = row[idColumn - 1]?.toString().trim();
        if (id) studentIdMap.set(id, idx + 2); // บวก 2 เพราะเริ่มที่แถว 2
      });

      editedValues.forEach(id => {
        const rowNum = studentIdMap.get(id);
        if (rowNum) {
          const currentStatus = sheet.getRange(rowNum, statusColumn).getValue();
          if (currentStatus !== "Accept") {
            sheet.getRange(rowNum, statusColumn).setValue("Accept");
          }
        }
      });
    }
  }

  // ส่วนของ "นักศึกษาก่อนปฏิบัติสหกิจ" ยังเหมือนเดิม (ไม่ต้องแก้)
  // [...เก็บไว้ตามเดิม...]
}




function updateConfirmationStatus() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('การยืนยันสิทธิ์-ผ่านแอป');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const aToF = row.slice(0, 6);
    const confirmStatus = row[6];  // คอลัมน์ G
    const currentStatus = row[7];  // คอลัมน์ H

    const hasData = aToF.some(cell => cell !== '');

    // กรณี 1: ถ้ามีข้อมูลใน A-F และ G ยังว่างหรือไม่ใช่ "ยืนยันสิทธิ์"/"สละสิทธิ์"
    if (hasData && confirmStatus !== 'ยืนยันสิทธิ์' && confirmStatus !== 'สละสิทธิ์') {
      sheet.getRange(i + 1, 7).setValue('รอการยืนยัน');         // G
      sheet.getRange(i + 1, 8).setValue('Pending Acceptance');   // H
    }

    // กรณี 2: ถ้า G เป็น ยืนยันสิทธิ์ → H = Accept
    if (confirmStatus === 'ยืนยันสิทธิ์') {
      sheet.getRange(i + 1, 8).setValue('Accept');
    }

    // กรณี 3: ถ้า G เป็น สละสิทธิ์ → H = Declined
    if (confirmStatus === 'สละสิทธิ์') {
      sheet.getRange(i + 1, 8).setValue('Declined');
    }

    // กรณี 4: ถ้า H เป็น No response → ไม่เปลี่ยน
    // ไม่ต้องเขียนอะไร เพราะไม่แตะต้อง No response
  }
}

function summarizeConfirmationStatus() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ผลการยืนยันสิทธิ์ทั้งหมด');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const g = row[6]; // คอลัมน์ G
    const h = row[7]; // คอลัมน์ H
    const iCol = row[8]; // คอลัมน์ I
    const allStatuses = [g, h, iCol];

    // ข้ามแถวว่าง (A-I ว่างทั้งหมด)
    const isEmpty = row.slice(0, 9).every(cell => cell === '');
    if (isEmpty) continue;

    let finalStatus = '';

    if (allStatuses.some(status => status === 'Declined' || status === 'No response')) {
      finalStatus = 'Declined';
    } else if (allStatuses.includes('Pending Acceptance')) {
      finalStatus = 'Pending Acceptance';
    } else if (allStatuses.every(status => status === 'Accept')) {
      finalStatus = 'Accept';
    } else {
      finalStatus = ''; // กรณีไม่เข้าเงื่อนไขใดเลย
    }

    sheet.getRange(i + 1, 10).setValue(finalStatus); // คอลัมน์ J
  }
}



function processLatestFile() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("ฟอร์มใบสมัคร");
  const folderIdImage = "1hFchAILcVl3OY-LzeGs0Do9sKe-iKb2f";
  const folderIdTranscriptResume = "14Wdny5Z5Zj12dQSbjO9hMcayXa0ikUYD";
  const folderImage = DriveApp.getFolderById(folderIdImage);
  const folderTranscriptResume = DriveApp.getFolderById(folderIdTranscriptResume);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const filePathImage = sheet.getRange(lastRow, 38).getValue(); // AL
  const filePathTranscript = sheet.getRange(lastRow, 39).getValue(); // AM
  const filePathResume = sheet.getRange(lastRow, 40).getValue(); // AN

  const fileNameImage = getCleanFileName(filePathImage);
  const fileNameTranscript = getCleanFileName(filePathTranscript);
  const fileNameResume = getCleanFileName(filePathResume);

  insertLink(folderImage, fileNameImage, sheet.getRange(lastRow, 45)); // AS
  insertLink(folderTranscriptResume, fileNameTranscript, sheet.getRange(lastRow, 46)); // AT
  insertLink(folderTranscriptResume, fileNameResume, sheet.getRange(lastRow, 47)); // AU
}

function getCleanFileName(filePath) {
  if (!filePath) return "";
  const parts = filePath.toString().split("/");
  return parts[parts.length - 1].trim();
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

function onChange(e) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("ฟอร์มใบสมัคร");
  const editedRow = sheet.getLastRow();
  if (editedRow < 2) return;
  processLatestFile();
}

function updatePendingAcceptanceStatus() {
  const sheetName = "การยืนยันสิทธิ์-ผ่านช่องทางสื่อสาร"; // ชื่อชีต
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();

  // เริ่มจากแถวที่ 2 เพราะแถวแรกคือหัวตาราง
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const aToF = row.slice(0, 6); // คอลัมน์ A ถึง F
    const currentStatus = row[6]; // คอลัมน์ G

    const hasData = aToF.every(cell => cell !== ""); // ตรวจว่ามีข้อมูลครบ

    // ถ้าข้อมูล A-F ครบ และสถานะ G ว่างหรือ "Pending Acceptance"
    if (hasData && (!currentStatus || currentStatus === "Pending Acceptance")) {
      sheet.getRange(i + 1, 7).setValue("Pending Acceptance"); // คอลัมน์ G คือคอลัมน์ที่ 7
    }
  }
}

function updatePendingAcceptanceStatusApp() {
  const sheetName = "การยืนยันสิทธิ์-ปฐมนิเทศน์";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const aToF = row.slice(0, 6); // คอลัมน์ A ถึง F
    const currentStatus = row[6]; // คอลัมน์ G

    const hasData = aToF.every(cell => cell !== "");

    if (hasData && (!currentStatus || currentStatus === "Pending Acceptance")) {
      sheet.getRange(i + 1, 7).setValue("Pending Acceptance"); // คอลัมน์ G
    }
  }
}

function checkGroupAndUpdate() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("คัดเลือกนักศึกษา"); // เลือกชีตที่ต้องการ
  var data = sheet.getRange("F2:F" + sheet.getLastRow()).getValues(); // ดึงข้อมูลในคอลัมน์ F
  var output = [];

  // ตรวจสอบข้อมูลในคอลัมน์ F และใส่ผลลัพธ์ลงใน output
  for (var i = 0; i < data.length; i++) {
    if (data[i][0].includes("CIS") || data[i][0].includes("ACS")) {
      output.push(["รังสิต"]);
    } else if (data[i][0].includes("LT")) {
      output.push(["ลำปาง"]);
    } else {
      output.push([""]); // ถ้าไม่มีข้อมูลที่ตรงตามเงื่อนไข ให้เป็นค่าว่าง
    }
  }

  // ใส่ผลลัพธ์ลงในคอลัมน์ G
  sheet.getRange(2, 7, output.length, 1).setValues(output); // ใส่ค่าลงในคอลัมน์ G
}











