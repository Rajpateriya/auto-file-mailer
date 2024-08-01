const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');
const imaps = require('imap-simple');
const { simpleParser } = require('mailparser');
const xlsx = require('xlsx');

// Load the email addresses from the Excel file
const workbook = xlsx.readFile("C:/Users/prabh/OneDrive/Desktop/Auto Renamer/mailer/book6.xlsx");
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

// Set the sender's credentials
const senderEmail = "team.simplbyte@gmail.com";
const senderPassword = "cnftskazfsaopney";
const senderName = "Simplbyte";
const attachmentFolder1 = "C:/Users/prabh/OneDrive/Desktop/SIMBT/folder 1";
const attachmentFolder2 = "C:/Users/prabh/OneDrive/Desktop/SIMBT/folder 2";

// Set the Gmail SMTP and IMAP server addresses
const smtpServer = "smtp.gmail.com";
const imapServer = "imap.gmail.com";

// Configure the SMTP transport
const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
        user: senderEmail,
        pass: senderPassword
    }
});

// Configure the IMAP connection
const imapConfig = {
    imap: {
        user: senderEmail,
        password: senderPassword,
        host: imapServer,
        port: 993,
        tls: true,
        authTimeout: 3000
    }
};

// Read rows from the Excel sheet
const rows = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

// Create a draft email for each recipient
async function createDrafts() {
    const connection = await imaps.connect(imapConfig);
    await connection.openBox('INBOX');

    for (let row of rows.slice(1)) {
        const receiverEmail = row[0];
        const attachmentFile1 = path.join(attachmentFolder1, `${row[1]}.pdf`);
        const attachmentFile2 = path.join(attachmentFolder2, `${row[2]}.pdf`);

        // Set the plain text message as a string
        const messageContent = `
        <html>
        <body>
            <p>Dear ${row[1]},</p>
            <p><b>Congratulations on Your Selection for Simplbyte Internship!</b></p>
            <p>We are thrilled to inform you that you have been selected for the Simplbyte ${row[2]} Internship Program! We would like to offer you the Internship Position at Simplbyte.</p>
            <p>Please find your <b>Offer Letter and Task List PDF</b> attached to the end of this email. Also, make sure to join our Insta page by clicking on the provided link. (Mandatory for future updates and opportunities)</p>
            <p><b>Important Details:</b></p>
            <ul>
                <li><b>Start Date:</b> August 2nd</li>
                <li><b>Last Date of Task Submission:</b> September 2nd</li>
                <li><b>Submission Form:</b> <a href="https://forms.gle/93mauBFfZZnepNZ28">https://forms.gle/93mauBFfZZnepNZ28</a> (Please submit your tasks through this form)</li>
            </ul>
            <p><b>Key Requirements:</b></p>
            <ul>
                <li><b>Update your LinkedIn profile and share your achievements, including the Offer Letter and Internship Completion Certificate. Tag Simplbyte (@Simplbyte) and use #simplbyte.</b></li>
                <li><b>No plagiarism in your project/code. It will lead to termination of your internship and a ban from future opportunities at Simplbyte.</b></li>
                <li><b>Share a video explaining your process on LinkedIn for each task submission. Tag Simplbyte and use #simplbyte.</b></li>
                <li><b>For Tech Internships, maintain a separate GitHub repository named SIMBT for all your tasks and share the link in the task submission form.</b></li>
            </ul>
            <p><b>Additional Information:</b></p>
            <ul>
                <li><b>Internship Completion Certificate will be sent to deserving candidates by September 3rd week.</b></li>
            </ul>
            <p><b>Join our social media platforms:</b></p>
            <ul>
                <li><b>LinkedIn:</b> <a href="https://www.linkedin.com/company/simplbyte/">https://www.linkedin.com/company/simplbyte/</a></li>
                <li><b>Facebook:</b> <a href="https://www.facebook.com/people/Simplbyte/100090440826505/">https://www.facebook.com/people/Simplbyte/100090440826505/</a></li>
                <li><b>Instagram:</b> <a href="https://www.instagram.com/simplbyte">https://www.instagram.com/simplbyte</a></li>
            </ul>
            <p><b>Contact Us:</b></p>
            <p>If you have any queries, please feel free to contact us at <b>team.simplbyte@gmail.com</b> or visit our website <b>http://simplbyte.tech/</b>.</p>
            <p>Best Regards,<br>
            <b>Team Simplbyte</b></p>
        </body>
        </html>`;

        // Create the email object
        let mailOptions = {
            from: `${senderName} <${senderEmail}>`,
            to: receiverEmail,
            subject: "Congratulations on Your Selection for Simplbyte Internship!",
            html: messageContent,
            attachments: []
        };

        // Attach the first file if it exists
        if (fs.existsSync(attachmentFile1)) {
            mailOptions.attachments.push({
                filename: path.basename(attachmentFile1),
                path: attachmentFile1
            });
        }

        // Attach the second file if it exists
        if (fs.existsSync(attachmentFile2)) {
            mailOptions.attachments.push({
                filename: path.basename(attachmentFile2),
                path: attachmentFile2
            });
        }

        // Send the email as a draft
        transporter.sendMail(mailOptions, (error, info) => {
            if (error) {
                return console.error('Error sending email:', error);
            }
            console.log(`Draft email created for ${receiverEmail}: ${info.messageId}`);
        });
    }

    // Close the IMAP connection
    connection.end();
}

createDrafts().catch(console.error);
