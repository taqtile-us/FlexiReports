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
const generate_excel_report_1 = require("./generate-excel-report");
const Formuler_1 = require("./Formuler");
const generate = (dataToFill, filePath) => __awaiter(void 0, void 0, void 0, function* () {
    const res = yield (0, generate_excel_report_1.generateExcelReport)(dataToFill, filePath);
    console.log(res, 'res');
    if (res) {
        const coords = yield (0, Formuler_1.getCoordsOfFormulaCell)(res);
        console.log(coords, 'coords');
        yield (0, Formuler_1.generateReportWithFormula)(res, coords);
    }
    return res;
});
exports.default = generate;
