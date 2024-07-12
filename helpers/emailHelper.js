const transporter = require('../config/emailConfig');

const sendEmail = (to, subject, text, attachmentPath, callback) => {
  const mailOptions = {
    from: 'your-email@gmail.com',
    to,
    subject,
    text,
    attachments: [
      {
        path: attachmentPath
      }
    ]
  };

  transporter.sendMail(mailOptions, (err, info) => {
    if (err) {
      return callback(err);
    }
    callback(null, info);
  });
};

module.exports = {
  sendEmail
};
