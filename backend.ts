import nodemailer from "nodemailer";
import XLSX from "xlsx";
import dotenv from "dotenv";

dotenv.config();
function sleep(ms: number) {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}
//eslint-disable-next-line
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

  for (let i = 0; i < data.length; i++) {
    const row :any = data[i]; // eslint-disable-line
    //set delay for sending emails
    await sleep(2000); //obs here could change this time lower, gmail can send only ~350 emails and after you get Too many login attempts for 10min.
    const mailOptions = {
      from: process.env.SEND_MAIL_USER,
      to: row[columnEmailName],
      subject: subject,
      text: text,
    };
    transporter.sendMail(mailOptions, async function (error) {
      if (error) {
        console.log("Error at: " + row[columnEmailName] + " " + error);
        await sleep(2000);
        transporter.sendMail(mailOptions, function (error) {
          if (error) {
            console.log("Error1 at: " + row[columnEmailName] + " " + error);
          } else {
            console.log("Email sent to: " + row[columnEmailName]);
          }
        });
      } else {
        console.log("Email sent to: " + row[columnEmailName]);
      }
    });
  }
}

async function sendMailviaXlsxHTML(excelName: string, sheetName: string, columnEmailName: string, columnsNameUser: string, subject: string) {
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

  for (let i = 0; i < data.length; i++) {
    const row :any = data[i]; // eslint-disable-line
    //set delay for sending emails
    await sleep(2000); //obs here could change this time lower, gmail can send only ~350 emails and after you get Too many login attempts for 10min.
    const mailOptions = {
      from: process.env.SEND_MAIL_USER,
      to: row[columnEmailName],
      subject: subject,
      html: html.replace("{{name}}", row[columnsNameUser]),
    };
    transporter.sendMail(mailOptions, async function (error) {
      if (error) {
        console.log("Error at: " + row[columnEmailName] + " " + error);
        await sleep(2000);
        transporter.sendMail(mailOptions, function (error) {
          if (error) {
            console.log("Error1 at: " + row[columnEmailName] + " " + error);
          } else {
            console.log("Email sent to: " + row[columnEmailName]);
          }
        });
      } else {
        console.log("Email sent to: " + row[columnEmailName]);
      }
    });
  }
}

// Example to use sendMailviaXlsxHTML
const html = `
<!DOCTYPE html>
<html lang="ro">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Simularea Examenului de Admitere</title>
</head>
<body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #fefefe; margin: 0; padding: 0;">
    <div class="container" style="max-width: 700px; margin: auto; background-color: #ffffff; padding: 20px; border-radius: 10px; box-shadow: 0px 0px 20px rgba(0, 0, 0, 0.1);">
        <h1 class="header" style="color: #333333; text-align: center; font-size: 32px; margin-bottom: 40px;">Bună ziua, {{name}}!</h1>
        <p class="text" style="color: #333333; line-height: 1.6;">
            🚀 Anul acesta, <strong style="color: #ff6600;">POLITEHNICA București</strong> are un număr record de peste <strong>2.700</strong> de elevi înscriși la <strong>Simularea Examenului de Admitere!</strong> <br><br>
            📚 Cu peste <strong>58</strong> de săli de concurs, suntem pregătiți să vă susținem în următorul vostru pas către viitorul academic! <br><br>
            ❗ Accesați contul de pe platforma <a href="https://admiteresim.pub.ro/">https://admiteresim.pub.ro/</a> pentru a afla sala în care sunteți repartizați. <br><br>
            🤩 Sute de voluntari au fost mobilizați pentru a face din Simularea Examenului de Admitere o experiență de neuitat pentru fiecare dintre voi! Vă așteptăm cu entuziasm!
            <br><br>
            📍 Căutăm puțin în arhiva noastră și redescoperim acest video care sperăm să fie util pentru voi, filmat în 2018, cu drumul spre <strong>Facultatea de Electronică, Telecomunicații și Tehnologia Informației.</strong>
            <div class="youtube-container" style="text-align: center;margin-top:20px;">
                <a href="https://www.youtube.com/watch?v=Ih1rZiJzNtE&ab_channel=LigaStuden%C8%9BilorElectroni%C8%99tiBucure%C8%99ti" style="display: inline-block; background-color: #ff6b6b; color: #ffffff; padding: 10px 20px; border-radius: 30px; text-decoration: none; font-weight: bold; font-size: 14px;">Vezi pe YouTube</a>
            </div>
        </p>
    </div>
</body>
</html>
`;

const columnsName = "Nume";

await sendMailviaXlsxHTML("data.xlsx", "ETTI", "Mail", columnsName, "Simularea Examenului de Admitere ETTI 2024"); //must be .xlsx file format
