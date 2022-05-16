const SENT_COLUMN = 4; // คอลัมน์ที่ คำว่า ตอบอีเมล์แล้ว ปรากฏ เริ่มนับจากคอลัมน์แรก คือ 1
const EMAIL_COLUMN = 2;  //คอลัมน์ที่มี email address อยู่ เริ่มนับจากคอลัมน์แรก คือ 1
const EMAIL_SENT_MESSAGE = "ตอบอีเมล์แล้ว"; //เอาไว้แสดงบน google sheet ว่าได้ ส่งอีเมลืไปแล้ว
const SUBJECT = "อีเมล์ส่งมาจาก สมาน วารุกา"; //หัวเรื่องใน อีเมล์
const MESSAGE = "ส่งไฟล์ แผนการสอน "; // ข้อคงวามใน อีเมล์
const FILE_URL = "https://drive.google.com/file/d/0Bw0uPsW6KdXNNXl2bFBiS3o5aTQ/view?usp=sharing&resourcekey=0-3tYV9EOyenjuliB5FDm1FA"

function sendEmail() {
  const sheet = SpreadsheetApp.getActiveSheet(); //ใช้ข้อมูลบน google sheet  ที่เปิดนี้
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues(); // ดึงข้อมูลมาทั้งหมด
  const fileId = getIdFromUrl(FILE_URL);
  const fileFromFileUrl = DriveApp.getFileById(fileId);

  for (let i = 1; i < data.length; i++) { //อ่านข้อมูลใน sheet
    const row = data[i]; // อ่าน แถว
    const emailAddress = row[EMAIL_COLUMN - 1];
    const emailSent = row[SENT_COLUMN - 1]; //กำหนดค่า ตัวแปร  emailSent คือ ข้อมูลใน คอลัมน์ที่ SENT_COLUMN
    Logger.log({ row, emailSent })
    if (emailSent != EMAIL_SENT_MESSAGE) { //ตรวจสอบในคอลัมน์ที่ SENT_COLUMN ว่า ไม่ใช่/ไม่มี ข้อความว่า "ตอบอีเมล์แล้ว"
      MailApp.sendEmail(emailAddress, SUBJECT, MESSAGE, {
        attachments: [fileFromFileUrl],
      }); //ให้ส่ง อีเมล์
      sheet.getRange(i + 1, SENT_COLUMN).setValue(EMAIL_SENT_MESSAGE); //ส่งแล้วก้ไป เขียนว่า "ตอบอีเมล์แล้ว"
      SpreadsheetApp.flush();
    }
  }
}

function getIdFromUrl(url) { return url.match(/[-\w]{25,}/); };
