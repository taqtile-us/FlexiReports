var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
import XlsxTemplate from 'xlsx-template';
import path from 'path';
import { readFile, writeFile } from 'fs/promises';
export const writeDataToExcel = (dataToFill, templatePath) => __awaiter(void 0, void 0, void 0, function* () {
    const filePath = 'uploads/report.xlsx';
    const data = yield readFile(path.join(templatePath));
    const template = new XlsxTemplate(data);
    const sheetNumber = 1;
    template.substitute(sheetNumber, dataToFill);
    const binaryFileData = template.generate();
    yield writeFile(filePath, binaryFileData, 'binary');
    return filePath;
});
