# Send Mailer

## Description

This is a simple mailer that sends an email using the `nodemailer` package.

## Prerequisites
Node.js and npm are required to run this project. You can download them from the following links:

- [Node.js](https://nodejs.org/en/download/)
- [npm](https://www.npmjs.com/get-npm)
- [npx](https://www.npmjs.com/package/npx)

## Installation

Clone the repository and navigate to the project directory in the terminal.

```bash
git clone https://github.com/ConducereIT/send-mailer/
cd send-mailer
```

To install the dependencies, run the following command in the terminal:

```bash
npm install 
```

## Configuration

To use the mailer, you need to configure the following environment variables:

- `SEND_MAIL_HOST`: The host of the mail server
- `SEND_MAIL_USER`: The username of the mail server
- `SEND_MAIL_PASS`: The password of the mail server
- `SEND_MAIL_SERVICE`: The service of the mail server

You can configure the environment variables by creating a `.env` file in the root directory of the project and adding
the following lines:

```env
SEND_MAIL_HOST=smtp.gmail.com
SEND_MAIL_USER=...
SEND_MAIL_PASS=...
SEND_MAIL_SERVICE=Gmail
```

You need to have a .xlsx file with a column for the email addresses. You can configure the file path by changing
the `excelName` variable in the `backend.ts` file.
example to run the mailer:

```typescript
await sendMailviaXlsx("data.xlsx", "ETTI", "Mail", "Test", "Text");
```

where:

- `data.xlsx` is the file path
- `ETTI` is the sheet name
- `Mail` is the column name for the email addresses
- `Test` is the subject
- `Text` is the message

```typescript
await sendMailviaXlsxHTML("data.xlsx", "ETTI", "Mail", "Test", "<body>a</body>");
```

where:

- `data.xlsx` is the file path
- `ETTI` is the sheet name
- `Mail` is the column name for the email addresses
- `Test` is the subject
- `<body>a</body>` is the html message

## Usage

To use the mailer, run the following command in the terminal:

```bash
npx tsx backend.ts
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details
