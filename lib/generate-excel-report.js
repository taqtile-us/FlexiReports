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
Object.defineProperty(exports, "__esModule", { value: true });
exports.generateExcelReport = void 0;
const Templater_1 = require("./Templater");
const generateExcelReport = (dataToFill, filePath, reportPath, temporaryFolderPath) => __awaiter(void 0, void 0, void 0, function* () {
    try {
        yield (0, Templater_1.writeDataToExcel)(dataToFill, filePath, reportPath, temporaryFolderPath);
        return reportPath;
    }
    catch (e) {
        console.log(e, 'generateExcelReport error');
        return false;
    }
});
exports.generateExcelReport = generateExcelReport;
