import * as xlsx from "xlsx";

export type OperationType = "приход" | "відход";

export type SheetName<T extends OperationType> =
    T extends "приход" ? "НАДІЙШЛО" :
    T extends "відход" ? "ВИБУЛО" :
    never;


export function getSheetName<T extends OperationType>(operation: T): SheetName<T> {
    return (operation === "приход" ? "НАДІЙШЛО" : "ВИБУЛО") as SheetName<T>;
}

/**
 * Повертає всі значення з діапазону C4:AO363 (включно з порожніми клітинками)
 * @param workbook xlsx.WorkBook
 * @param sheetName назва аркуша
 * @returns Масив масивів з рядками і комірками (порожні клітинки = "")
 */
export function getDataRangeFromSheet(workbook: xlsx.WorkBook, sheetName: string): string[][] {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) {
        throw new Error(`Аркуш "${sheetName}" не знайдено у файлі`);
    }

    const startCol = xlsx.utils.decode_col("C"); // 2
    const endCol = xlsx.utils.decode_col("AO");  // 40
    const startRow = 4;
    const endRow = 363;

    const result: string[][] = [];

    for (let row = startRow; row <= endRow; row++) {
        const currentRow: string[] = [];
        for (let col = startCol; col <= endCol; col++) {
            const cellAddress = xlsx.utils.encode_cell({ c: col, r: row - 1 }); // -1 бо 0-індексація
            const cell = worksheet[cellAddress];
            currentRow.push(cell ? String(cell.v) : "");
        }
        result.push(currentRow);
    }

    return result;
}

export function stringifyTableAsTSV(table: string[][]): string {
    return table.map(row => row.join("\t")).join("\n");
}
