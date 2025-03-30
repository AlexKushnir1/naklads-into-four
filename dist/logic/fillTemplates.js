"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
exports.fillTemplatesFromParsed = fillTemplatesFromParsed;
const xlsx = __importStar(require("xlsx"));
const writeToTemplate_1 = require("./writeToTemplate");
const AllTemplates_1 = require("../interfaces/AllTemplates");
function fillTemplatesFromParsed(data) {
    // Читаємо шаблони один раз
    const zagalna = readWorkbook("data/zagalna.xlsx");
    const khmilnyk = readWorkbook("data/khmilnyk.xlsx");
    const koziatyn = readWorkbook("data/koziatyn.xlsx");
    const kalynivka = readWorkbook("data/kalynivka.xlsx");
    // Вставляємо дані
    data.forEach((item) => {
        (0, writeToTemplate_1.writeToTemplate)(zagalna, item.parsedNakladna.zagalna, item.operation, item.cell);
        (0, writeToTemplate_1.writeToTemplate)(khmilnyk, item.parsedNakladna.khmilnyk, item.operation, item.cell);
        (0, writeToTemplate_1.writeToTemplate)(koziatyn, item.parsedNakladna.koziatyn, item.operation, item.cell);
        (0, writeToTemplate_1.writeToTemplate)(kalynivka, item.parsedNakladna.kalynivka, item.operation, item.cell);
    });
    // Повертаємо оновлені шаблони
    return new AllTemplates_1.AllTemplates(zagalna, khmilnyk, koziatyn, kalynivka);
}
function readWorkbook(filePath) {
    const workbook = xlsx.readFile(filePath);
    if (!workbook) {
        throw new Error(`Не вдалося прочитати файл: ${filePath}`);
    }
    return workbook;
}
