function doPost(e) {
  const sheet = SpreadsheetApp.openById("1uioV665lMMYwHAEBLulZMjppYvaeimFoXU09s_xT6d0").getSheetByName("Sarabun");
  const folderId = "1PpMQ2KCy77BZ6tccUiZkoPXVNM7CLP2p?usp=drive_link";
  const year = 2568; // ปีเริ่มต้น
  const now = new Date();
  const from = e.parameter.from;
  const to = e.parameter.to;
  const subject = e.parameter.subject;
  const category = e.parameter.category;

  // หาเลขลำดับล่าสุดของปีนี้
  const data = sheet.getDataRange().getValues();
  let count = 0;
  data.forEach(row => {
    if (String(row[1]).includes("/" + year)) {
      count++;
    }
  });

  const runningNumber = (count + 1) + "/" + year;

  // อัปโหลดไฟล์แนบถ้ามี
  let fileUrl = "";
  if (e.parameter.file) {
    const blob = Utilities.newBlob(Utilities.base64Decode(e.parameter.file), MimeType.PDF, "upload.pdf");
    const file = DriveApp.getFolderById(folderId).createFile(blob);
    fileUrl = file.getUrl();
  }

  // เพิ่มข้อมูลในชีต
  sheet.appendRow([
    now,
    runningNumber,
    from,
    to,
    subject,
    category,
    fileUrl
  ]);

  return ContentService.createTextOutput("บันทึกสำเร็จ เลขสารบรรณ: " + runningNumber);
}
