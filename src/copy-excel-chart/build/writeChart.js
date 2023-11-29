const fs = require('fs');
const archiver = require('archiver');
export function writeCharts(targetExcel, zipFilePath) {
    const sourceFolder = targetExcel.tempDir;
  return new Promise((resolve, reject) => {
    const output = fs.createWriteStream(zipFilePath);
    const archive = archiver('zip', {
      zlib: { level: 9 } // Set compression level
    });

    output.on('close', function () {
      resolve(zipFilePath);
    });

    archive.on('warning', function (err) {
      if (err.code === 'ENOENT') {
        console.warn(err);
      } else {
        reject(err);
      }
    });

    archive.on('error', function (err) {
      reject(err);
    });

    archive.pipe(output);

    archive.directory(sourceFolder, false);

    archive.finalize();
  });
}

//# sourceMappingURL=writeChart.js.map