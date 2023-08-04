"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.generateExcelReport = void 0;
const exceljs_1 = __importDefault(require("exceljs"));
const Templater_1 = require("./Templater");
const generateExcelReport = async (dataToFill, filePath) => {
    try {
        const workbook = new exceljs_1.default.Workbook();
        await workbook.xlsx.readFile(filePath);
        const pathToReport = await (0, Templater_1.writeDataToExcel)(dataToFill, filePath);
        return pathToReport;
    }
    catch (e) {
        console.log(e, 'generateExcelReport error');
        return false;
    }
};
exports.generateExcelReport = generateExcelReport;
