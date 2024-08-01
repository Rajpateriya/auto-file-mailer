const nodemailer = require('nodemailer');
const imapSimple = require('imap-simple');
const { simpleParser } = require('mailparser');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

// Load the email addresses from the Excel file
const workbook = xlsx.readFile('C:/Users/prabh/OneDrive/Desktop/Auto Renamer/mailer/book7.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const rows = xlsx.utils.sheet_to_json(worksheet);

// Set the sender's credentials
const senderEmail = 'team.simplbyte@gmail.com';
const senderPassword = 'bljfyntxshzkwcsl';
const senderName = 'Simplbyte';
const attachmentFolder = 'C:/Users/prabh/OneDrive/Desktop/SIMBT';

// Create the SMTP transporter
const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
        user: senderEmail,
        pass: senderPassword
    }
});

// IMAP configuration
const imapConfig = {
    imap: {
        user: senderEmail,
        password: senderPassword,
        host: 'imap.gmail.com',
        port: 993,
        tls: true,
        authTimeout: 3000
    }
};

(async () => {
    try {
        // Connect to IMAP
        const connection = await imapSimple.connect(imapConfig);
        await connection.openBox('Drafts');

        for (let row of rows) {
            const receiverEmail = row['Email'];
            const attachmentFile = path.join(attachmentFolder, `${row['Name']}.pdf`);

            // Set the plain text message as a string
            const messageContent = `
Dear ${row['Name']};

Certificate ID: ${row['CertificateID']}

Congratulations! You've successfully completed the Student Internship Program : SIMBT in ${row['Course']}.
On behalf of Simplbyte we are pleased to issue your SIMBT Certificate.

Please find your Certificate in the attachment with this email.
As you have earned the certificate, why not to share it.

Ways to share:

✪Add your Student Internship Program: SIMBT Certificate directly to your LinkedIn profile

Download PDF to print to share

Share on social media or on your LinkedIn Feed and tag Simplbyte
✪Tag: @Simplbyte
✪You can also use hashtags - #simplbyte #simplbyteintern while while posting your Completion Certificate.
✪Follow our insta page to stay updated future paid opportunities with simplbyte:
✪Instagram: https://www.instagram.com/simplbyte/
✪Letter of recommendation will be released to deserving one in few days.

Once again, congratulations on your achievement.

If any Query?
Please feel free to contact us at-

✪Email : team.simplbyte@gmail.com
✪Website : simplbyte.tech
✪Instagram: https://www.instagram.com/simplbyte/

Best Regards,
Team Simplbyte

Disclaimer
--------------------------------------------------------------------------------------------------------------------------------
ᴛʜɪꜱ ᴇ-ᴍᴀɪʟ ᴍᴀʏ ᴄᴏɴᴛᴀɪɴ ᴘʀɪᴠɪʟᴇɢᴇᴅ ᴀɴᴅ ᴄᴏɴꜰɪᴅᴇɴᴛɪᴀʟ ɪɴꜰᴏʀᴍᴀᴛɪᴏɴ ᴡʜɪᴄʜ ɪꜱ ᴛʜᴇ ᴘʀᴏᴘᴇʀᴛʏ ᴏꜰ simplbyte. ɪᴛ ɪꜱ ɪɴᴛᴇɴᴅᴇᴅ ᴏɴʟʏ ꜰᴏʀ ᴛʜᴇ ᴜꜱᴇ ᴏꜰ ᴛʜᴇ ɪɴᴅɪᴠɪᴅᴜᴀʟ ᴏʀ ᴇɴᴛɪᴛʏ ᴛᴏ ᴡʜɪᴄʜ ɪᴛ ɪꜱ ᴀᴅᴅʀᴇꜱꜱᴇᴅ. ɪꜰ ʏᴏᴜ ᴀʀᴇ ɴᴏᴛ ᴛʜᴇ ɪɴᴛᴇɴᴅᴇᴅ ʀᴇᴄɪᴘɪᴇɴᴛ, ʏᴏᴜ ᴀʀᴇ ɴᴏᴛ ᴀᴜᴛʜᴏʀɪᴢᴇᴅ ᴛᴏ ʀᴇᴀᴅ, ʀᴇᴛᴀɪɴ, ᴄᴏᴘʏ, ᴘʀɪɴᴛ, ᴅɪꜱᴛʀɪʙᴜᴛᴇ ᴏʀ ᴜꜱᴇ ᴛʜɪꜱ ᴍᴇꜱꜱᴀɢᴇ. ɪꜰ ʏᴏᴜ ʜᴀᴠᴇ ʀᴇᴄᴇɪᴠᴇᴅ ᴛʜɪꜱ ᴄᴏᴍᴍᴜɴɪᴄᴀᴛɪᴏɴ ɪɴ ᴇʀʀᴏʀ, ᴘʟᴇᴀꜱᴇ ɴᴏᴛɪꜰʏ ᴛʜᴇ ꜱᴇɴᴅᴇʀ ᴀɴᴅ ᴅᴇʟᴇᴛᴇ ᴀʟʟ ᴄᴏᴘɪᴇꜱ ᴏꜰ ᴛʜɪꜱ ᴍᴇꜱꜱᴀɢᴇ. simplbyte, ᴅᴏᴇꜱ ɴᴏᴛ ᴀᴄᴄᴇᴘᴛ ᴀɴʏ ʟɪᴀʙɪʟɪᴛʏ ꜰᴏʀ ᴠɪʀᴜꜱ ɪɴꜰᴇᴄᴛᴇᴅ ᴍᴀɪʟꜱ.
`;

            // Create the email options
            const mailOptions = {
                from: `"${senderName}" <${senderEmail}>`,
                to: receiverEmail,
                subject: '𝗖𝗼𝗻𝗴𝗿𝗮𝘁𝘂𝗹𝗮𝘁𝗶𝗼𝗻𝘀 || Your Certificate Arrived..',
                text: messageContent,
                attachments: [
                    {
                        filename: `${row['Name']}.pdf`,
                        path: attachmentFile
                    }
                ],
                headers: {
                    'X-Priority': '2'
                }
            };

            // Send the email
            const info = await transporter.sendMail(mailOptions);
            console.log(`Draft email created for ${receiverEmail}: ${info.messageId}`);
        }

        // Close the IMAP connection
        connection.end();
    } catch (err) {
        console.error('Error: ', err);
    }
})();
