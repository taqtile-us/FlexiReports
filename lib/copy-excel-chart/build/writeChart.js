"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.writeCharts = void 0;
const fs = require('fs');
const archiver = require('archiver');
function writeCharts(targetExcel, zipFilePath) {
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
            }
            else {
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
exports.writeCharts = writeCharts;
//# sourceMappingURL=writeChart.js.map
