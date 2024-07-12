const XLSX = require('xlsx');
const { renameFile, readFilesFromFolder } = require('./helpers/fileHelper');
const { sendEmail } = require('./helpers/emailHelper');

// Load the Excel file
const workbook = XLSX.readFile('./data/data.xlsx');
const sheetName = workbook.SheetNames[0];
const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

// Folder where your files are located
const folderPath = './files/';

readFilesFromFolder(folderPath, (err, files) => {
  if (err) {
    return console.log('Unable to scan directory: ' + err);
  }

  files.forEach((file, index) => {
    const { name, email } = sheet[index];
    const newFileName = `${name}.pdf`;

    renameFile(folderPath, file, newFileName, (err, newFilePath) => {
      if (err) {
        return console.log('Error renaming file: ' + err);
      }

      console.log(`File renamed to ${newFileName}`);

      sendEmail(email, 'Here is your file', `Dear ${name},\n\nPlease find the attached file.`, newFilePath, (err, info) => {
        if (err) {
          return console.log('Error sending email: ' + err);
        }

        console.log('Email sent: ' + info.response);
      });
    });
  });
});
