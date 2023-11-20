"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const copy_excel_chart_1 = __importDefault(require("copy-excel-chart"));
const readCharts = copy_excel_chart_1.default.readCharts;
const copyChart = copy_excel_chart_1.default.copyChart;
const writeCharts = copy_excel_chart_1.default.writeCharts;
console.log(writeCharts);
