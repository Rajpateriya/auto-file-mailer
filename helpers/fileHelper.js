const fs = require('fs');
const path = require('path');

const renameFile = (folderPath, oldName, newName, callback) => {
  const oldFilePath = path.join(folderPath, oldName);
  const newFilePath = path.join(folderPath, newName);

  fs.rename(oldFilePath, newFilePath, (err) => {
    if (err) {
      return callback(err);
    }
    callback(null, newFilePath);
  });
};

const readFilesFromFolder = (folderPath, callback) => {
  fs.readdir(folderPath, (err, files) => {
    if (err) {
      return callback(err);
    }
    callback(null, files);
  });
};

module.exports = {
  renameFile,
  readFilesFromFolder
};
