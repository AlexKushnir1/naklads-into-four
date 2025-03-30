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
exports.writeToTemplate = writeToTemplate;
const xlsx = __importStar(require("xlsx"));
/**
 * Вставляє кількості з мапи у відповідний аркуш книги Excel, починаючи з рядка 4 у вибраній колонці
 * @param workbook Відкритий Excel-файл (xlsx.readFile(...))
 * @param dataMap Мапа: код => кількість
 * @param operation "приход" або "відход" — обирає аркуш
 * @param columnLetter Наприклад "C", "D" — назва колонки (рядок завжди з 4)
 */
function writeToTemplate(workbook, dataMap, operation, columnLetter) {
    const sheetName = operation === "приход" ? "НАДІЙШЛО" : "ВИБУЛО";
    const worksheet = workbook.Sheets[sheetName];
    console.log("➡ Вставка в аркуш:", sheetName);
    console.log("🧾 Колонка:", columnLetter);
    console.log("🗂️ Кількість елементів:", dataMap.size);
    if (!worksheet) {
        throw new Error(`Аркуш "${sheetName}" не знайдено в Excel-файлі.`);
    }
    let row = 4; // завжди починаємо з 4-го рядка
    for (const [, quantity] of dataMap) {
        const cell = `${columnLetter}${row}`;
        console.log("➡ Вставка у клітинку:", cell);
        console.log("➡ Значення:", quantity);
        xlsx.utils.sheet_add_aoa(worksheet, [[quantity]], { origin: cell });
        row++;
    }
    //   Оновлюємо межі аркуша (!ref), щоб зміни відображались
    const range = xlsx.utils.decode_range(worksheet["!ref"] || `${columnLetter}4:${columnLetter}${row}`);
    range.e.r = Math.max(range.e.r, row);
    worksheet["!ref"] = xlsx.utils.encode_range(range);
}
