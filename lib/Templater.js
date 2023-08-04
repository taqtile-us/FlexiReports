"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.writeDataToExcel = void 0;
const XlsxTemplate = require('xlsx-template-ex');
const path_1 = __importDefault(require("path"));
const promises_1 = require("fs/promises");
const writeDataToExcel = async (dataToFill, templatePath) => {
    const filePath = 'uploads/report.xlsx';
    const buffer = await XlsxTemplate.xlsxBuildByTemplate(dataToFill, path_1.default.join(templatePath));
    await (0, promises_1.writeFile)(filePath, buffer, 'binary');
    return filePath;
};
exports.writeDataToExcel = writeDataToExcel;
