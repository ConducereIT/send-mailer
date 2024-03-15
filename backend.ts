import nodemailer from "nodemailer";
import XLSX from "xlsx";
import dotenv from "dotenv";

dotenv.config();

async function sendMailviaXlsx(excelName: string, sheetName: string, columnEmailName: string, subject: string, text: string) {
  const transporter = nodemailer.createTransport({
    host: process.env.SEND_MAIL_HOST,
    service: process.env.SEND_MAIL_SERVICE,
    auth: {
      user: process.env.SEND_MAIL_USER,
      pass: process.env.SEND_MAIL_PASS,
    },
  });

  const workbook = XLSX.readFile(excelName);
  const sheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(sheet);

  data.forEach((row: any) => { // eslint-disable-line
    const mailOptions = {
      from: process.env.SEND_MAIL_USER,
      to: row[columnEmailName],
      subject: subject,
      text: text,
    };

    transporter.sendMail(mailOptions, function (error) {
      if (error) {
        console.log("Error at: " + row[columnEmailName] + " " + error);
      } else {
        console.log("Email sent to: " + row[columnEmailName]);
      }
    });
  });
}

await sendMailviaXlsx("data.xlsx", "ETTI", "Mail", "Test", "Test"); //must be .xlsx file format