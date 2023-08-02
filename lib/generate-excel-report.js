var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
import ExcelJS from 'exceljs';
import { writeDataToExcel } from './Templater';
export const generateExcelReport = (dataToFill, filePath) => __awaiter(void 0, void 0, void 0, function* () {
    try {
        const workbook = new ExcelJS.Workbook();
        yield workbook.xlsx.readFile(filePath);
        const pathToReport = yield writeDataToExcel(dataToFill, filePath);
        return pathToReport;
    }
    catch (e) {
        console.log(e, 'generateExcelReport error');
        return false;
    }
});
