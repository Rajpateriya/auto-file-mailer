const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// Set the directory containing the files to be renamed
const directory = 'C:/Users/prabh/OneDrive/Desktop/SIMBT/ilovepdf_extracted-pages (20)';

// Load the workbook and select the worksheet
const workbook = xlsx.readFile('C:/Users/prabh/OneDrive/Desktop/Auto Renamer/Renamer/BOOK4.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

// Convert the worksheet to JSON
const rows = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

// Loop through each file in the directory and rename it
fs.readdir(directory, (err, files) => {
    if (err) {
        return console.error('Error reading directory:', err);
    }

    // Sort files based on the number in the filename
    files = files.sort((a, b) => {
        const numA = parseInt(a.split('-')[1].split('.')[0]);
        const numB = parseInt(b.split('-')[1].split('.')[0]);
        return numA - numB;
    });

    // Loop through the files and rename them
    files.forEach((filename, i) => {
        if (i < rows.length) {
            const newName = rows[i][0];
            const newFilename = newName + '.pdf';

            // Rename the file
            fs.rename(
                path.join(directory, filename),
                path.join(directory, newFilename),
                (err) => {
                    if (err) {
                        return console.error('Error renaming file:', err);
                    }

                    // Print a message to indicate that the file has been renamed
                    console.log(`Renamed file ${filename} to ${newFilename}`);
                }
            );
        }
    });

    // Check if there are any remaining rows in the worksheet that were not used
    if (files.length < rows.length) {
        console.log(`WARNING: There are ${rows.length - files.length} unused rows in the worksheet.`);
    }
});
