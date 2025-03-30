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
exports.getSheetName = getSheetName;
exports.getDataRangeFromSheet = getDataRangeFromSheet;
exports.stringifyTableAsTSV = stringifyTableAsTSV;
const xlsx = __importStar(require("xlsx"));
function getSheetName(operation) {
    return (operation === "приход" ? "НАДІЙШЛО" : "ВИБУЛО");
}
/**
 * Повертає всі значення з діапазону C4:AO363 (включно з порожніми клітинками)
 * @param workbook xlsx.WorkBook
 * @param sheetName назва аркуша
 * @returns Масив масивів з рядками і комірками (порожні клітинки = "")
 */
function getDataRangeFromSheet(workbook, sheetName) {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) {
        throw new Error(`Аркуш "${sheetName}" не знайдено у файлі`);
    }
    const startCol = xlsx.utils.decode_col("C"); // 2
    const endCol = xlsx.utils.decode_col("AO"); // 40
    const startRow = 4;
    const endRow = 363;
    const result = [];
    for (let row = startRow; row <= endRow; row++) {
        const currentRow = [];
        for (let col = startCol; col <= endCol; col++) {
            const cellAddress = xlsx.utils.encode_cell({ c: col, r: row - 1 }); // -1 бо 0-індексація
            const cell = worksheet[cellAddress];
            currentRow.push(cell ? String(cell.v) : "");
        }
        result.push(currentRow);
    }
    return result;
}
function stringifyTableAsTSV(table) {
    return table.map(row => row.join("\t")).join("\n");
}
