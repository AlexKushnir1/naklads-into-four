import * as xlsx from "xlsx";
import { ParsedNakladna } from "../interfaces/ParsedNakladna";

// Enum для назв аркушів
export enum SheetName {
    ZAGALNA = "Загальна",
    KHMILNYK = "Хмільник 1",
    KOZIATYN = "Козятин 1",
    KALYNIVKA = "Калинівка 1",
}

// Додаємо maxRows в конфіг
type SheetConfig = {
    key: keyof ParsedNakladna;
    sheetName: SheetName;
    startRow: number;
    codeCol: string;
    qtyCol: string;
    maxRows: number;
};

const sheetConfigs: SheetConfig[] = [
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
function parseSheetToMap(
    worksheet: xlsx.WorkSheet | undefined,
    config: SheetConfig
): Map<string, number> {
    const map = new Map<string, number>();

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

        if (code === "") continue; // ✅ Пропускаємо порожні рядки

        if (code.length !== 5) {
            throw new Error(`Invalid code length in row ${row}: ${code}`);
        }

        let qty = 0;
        try {
            if (
                qtyRaw === undefined ||
                qtyRaw === null ||
                qtyRaw === "-" ||
                qtyRaw === ""
            ) {
                qty = 0;
            } else if (typeof qtyRaw === "string") {
                qty = parseFloat(qtyRaw.replace(",", "."));
            } else {
                qty = Number(qtyRaw);
            }

            if (isNaN(qty)) qty = 0;
        } catch (e) {
            console.error(`❌ Error parsing qty in row ${row}: ${e}`);
            continue;
        }

        map.set(code, qty);
    }

    return map;
}

// Основна функція
export function parseNakladna(workbook: xlsx.WorkBook): ParsedNakladna {
    try {
        if (!workbook) {
            throw new Error("Workbook is undefined or null.");
        }
    } catch (e) {
        console.error(`❌ Failed to process workbook: ${e}`);
        throw e;
    }

    const result: any = {};

    for (const config of sheetConfigs) {
        const sheet = workbook.Sheets[config.sheetName];
        result[config.key] = parseSheetToMap(sheet, config);
    }

    return new ParsedNakladna(
        result.zagalna,
        result.khmilnyk,
        result.koziatyn,
        result.kalynivka
    );
}
