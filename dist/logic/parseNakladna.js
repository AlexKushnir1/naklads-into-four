"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.SheetName = void 0;
exports.parseNakladna = parseNakladna;
const ParsedNakladna_1 = require("../interfaces/ParsedNakladna");
// Enum для назв аркушів
var SheetName;
(function (SheetName) {
    SheetName["ZAGALNA"] = "\u0417\u0430\u0433\u0430\u043B\u044C\u043D\u0430";
    SheetName["KHMILNYK"] = "\u0425\u043C\u0456\u043B\u044C\u043D\u0438\u043A 1";
    SheetName["KOZIATYN"] = "\u041A\u043E\u0437\u044F\u0442\u0438\u043D 1";
    SheetName["KALYNIVKA"] = "\u041A\u0430\u043B\u0438\u043D\u0456\u0432\u043A\u0430 1";
})(SheetName || (exports.SheetName = SheetName = {}));
const sheetConfigs = [
    {
        key: "zagalna",
        sheetName: SheetName.ZAGALNA,
        startRow: 4,
        codeCol: "C",
        qtyCol: "D",
        maxRows: 363, // ✅ Загальна має 363 рядки
    },
    {
        key: "khmilnyk",
        sheetName: SheetName.KHMILNYK,
        startRow: 19,
        codeCol: "C",
        qtyCol: "F",
        maxRows: 378,
    },
    {
        key: "koziatyn",
        sheetName: SheetName.KOZIATYN,
        startRow: 19,
        codeCol: "C",
        qtyCol: "F",
        maxRows: 378,
    },
    {
        key: "kalynivka",
        sheetName: SheetName.KALYNIVKA,
        startRow: 19,
        codeCol: "C",
        qtyCol: "F",
        maxRows: 378,
    },
];
// ✅ Тепер без параметра maxRows
function parseSheetToMap(worksheet, config) {
    const map = new Map();
    if (!worksheet) {
        console.error(`❌ Sheet "${config.sheetName}" not found.`);
        throw new Error(`Sheet "${config.sheetName}" not found.`);
    }
    for (let row = config.startRow; row <= config.maxRows; row++) {
        const codeCell = worksheet[`${config.codeCol}${row}`];
        const qtyCell = worksheet[`${config.qtyCol}${row}`];
        const codeRaw = codeCell?.v;
        const qtyRaw = qtyCell?.v;
        const code = typeof codeRaw === "undefined" ? "" : String(codeRaw).trim();
        if (code === "")
            continue; // ✅ Пропускаємо порожні рядки
        if (code.length !== 5) {
            throw new Error(`Invalid code length in row ${row}: ${code}`);
        }
        let qty = 0;
        try {
            if (qtyRaw === undefined ||
                qtyRaw === null ||
                qtyRaw === "-" ||
                qtyRaw === "") {
                qty = 0;
            }
            else if (typeof qtyRaw === "string") {
                qty = parseFloat(qtyRaw.replace(",", "."));
            }
            else {
                qty = Number(qtyRaw);
            }
            if (isNaN(qty))
                qty = 0;
        }
        catch (e) {
            console.error(`❌ Error parsing qty in row ${row}: ${e}`);
            continue;
        }
        map.set(code, qty);
    }
    return map;
}
// Основна функція
function parseNakladna(workbook) {
    try {
        if (!workbook) {
            throw new Error("Workbook is undefined or null.");
        }
    }
    catch (e) {
        console.error(`❌ Failed to process workbook: ${e}`);
        throw e;
    }
    const result = {};
    for (const config of sheetConfigs) {
        const sheet = workbook.Sheets[config.sheetName];
        result[config.key] = parseSheetToMap(sheet, config);
    }
    return new ParsedNakladna_1.ParsedNakladna(result.zagalna, result.khmilnyk, result.koziatyn, result.kalynivka);
}
