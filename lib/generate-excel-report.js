"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.generateExcelReport = void 0;
const exceljs_1 = __importDefault(require("exceljs"));
const Templater_1 = require("./Templater");
const generateExcelReport = (dataToFill, filePath) => __awaiter(void 0, void 0, void 0, function* () {
    try {
        const workbook = new exceljs_1.default.Workbook();
        yield workbook.xlsx.readFile(filePath);
        const pathToReport = yield (0, Templater_1.writeDataToExcel)(dataToFill, filePath);
        return pathToReport;
    }
    catch (e) {
        console.log(e, 'generateExcelReport error');
        return false;
    }
});
exports.generateExcelReport = generateExcelReport;
