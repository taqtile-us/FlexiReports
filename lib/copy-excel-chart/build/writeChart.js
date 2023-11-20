"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.writeCharts = void 0;
const adm_zip_1 = __importDefault(require("adm-zip"));
function writeCharts(targetExcel, printPath) {
    return new Promise((resolve, reject) => {
        try {
            const targetDir = targetExcel.tempDir;
            const zip = new adm_zip_1.default();
            zip.addLocalFolder(targetDir, '');
            zip.writeZip(printPath);
            resolve(true);
        }
        catch (error) {
            console.log('Write chart file error. targetExcel: ', targetExcel, 'Error: ', error);
            reject(error);
        }
    });
}
exports.writeCharts = writeCharts;
//# sourceMappingURL=writeChart.js.map
