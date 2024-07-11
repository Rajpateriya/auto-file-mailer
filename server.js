const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const nodemailer = require('nodemailer');

// Load the Excel file
const workbook = XLSX.readFile('data.xlsx');
const sheetName = workbook.SheetNames[0];
const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

// Create a transporter for sending emails
const transporter = nodemailer.createTransport({
  service: 'gmail', // You can use other services
  auth: {
    user: 'your-email@gmail.com',
    pass: 'your-email-password'
  }
});

// Folder where your files are located
const folderPath = './files/';

fs.readdir(folderPath, (err, files) => {
  if (err) {
    return console.log('Unable to scan directory: ' + err);
  }

  files.forEach((file, index) => {
    const filePath = path.join(folderPath, file);

    // Assuming the order in the Excel sheet corresponds to the files in the folder
    const { name, email } = sheet[index];
    const newFileName = `${name}.pdf`;
    const newFilePath = path.join(folderPath, newFileName);

    // Rename the file
    fs.rename(filePath, newFilePath, (err) => {
      if (err) {
        return console.log('Error renaming file: ' + err);
      }

      console.log(`File renamed to ${newFileName}`);

      // Send the email with the renamed file
      const mailOptions = {
        from: 'your-email@gmail.com',
        to: email,
        subject: 'Here is your file',
        text: `Dear ${name},\n\nPlease find the attached file.`,
        attachments: [
          {
            filename: newFileName,
            path: newFilePath
          }
        ]
      };

      transporter.sendMail(mailOptions, (err, info) => {
        if (err) {
          return console.log('Error sending email: ' + err);
        }

        console.log('Email sent: ' + info.response);
      });
    });
  });
});
