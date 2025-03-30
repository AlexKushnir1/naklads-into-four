import * as path from "path";
import * as xlsx from "xlsx";
import { getDataRangeFromSheet, stringifyTableAsTSV } from "../units";

describe("getDataRangeFromSheet", () => {
    const filePath = path.resolve(__dirname, "../../data/zagalna.xlsx"); // або інший файл
    const sheetName = "НАДІЙШЛО"; // Переконайся, що такий аркуш є

    let workbook: xlsx.WorkBook;

    beforeAll(() => {
        workbook = xlsx.readFile(filePath);
    });

    it("повертає правильний масив з діапазону C4:AO363", () => {
        const data = getDataRangeFromSheet(workbook, sheetName);

        // Перевіримо кількість рядків (363 - 4 + 1)
        expect(data.length).toBe(360);

        // Перевіримо кількість колонок (AO = 41, C = 3 → 41 - 3 + 1 = 39)
        expect(data[0].length).toBe(39);

        // Перевіримо, що значення є текстами або порожніми
        data.forEach(row => {
            row.forEach(cell => {
                expect(typeof cell === "string" || cell === undefined).toBeTruthy();
            });
        });
        
    });

    it("викидає помилку, якщо аркуш не існує", () => {
        expect(() => {
            getDataRangeFromSheet(workbook, "НемаТакогоАркуша");
        }).toThrow("Аркуш \"НемаТакогоАркуша\" не знайдено у файлі");
    });
});