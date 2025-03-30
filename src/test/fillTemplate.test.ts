import { fillTemplatesFromParsed } from "../logic/fillTemplates";
import { NakladnaWithMeta } from "../interfaces/NakladnaWithMeta";
import * as xlsx from "xlsx";

// Мокуємо writeToTemplate і readWorkbook (опціонально, якщо хочеш ізольовано)
jest.mock("../logic/writeToTemplate", () => ({
    writeToTemplate: jest.fn()
}));

// Припускаємо, що тестові шаблони є у /data
// Якщо хочеш змокати і readWorkbook — теж можу показати як

describe("fillTemplatesFromParsed", () => {
    const testData: NakladnaWithMeta[] = [
        {
            operation: "приход",
            cell: "B5",
            parsedNakladna: {
                zagalna: new Map([
                    ["11111", 10.5],
                    ["22222", 5.2],
                ]),
                khmilnyk: new Map([
                    ["11111", 3.3]
                ]),
                koziatyn: new Map([
                    ["22222", 2.2]
                ]),
                kalynivka: new Map([
                    ["11111", 1.1],
                    ["22222", 4.4]
                ])
            }
        }
    ];

    it("повертає об'єкт AllTemplates з workbook-ами", () => {
        const result = fillTemplatesFromParsed(testData);

        // Перевіряємо, що кожен workbook є валідним об'єктом Excel
        expect(result).toHaveProperty("zagalna");
        expect(result).toHaveProperty("khmilnyk");
        expect(result).toHaveProperty("koziatyn");
        expect(result).toHaveProperty("kalynivka");

        // Мінімальні перевірки на структуру workbook
        expect(result.zagalna.SheetNames.length).toBeGreaterThan(0);
        expect(result.khmilnyk.SheetNames.length).toBeGreaterThan(0);
        expect(result.koziatyn.SheetNames.length).toBeGreaterThan(0);
        expect(result.kalynivka.SheetNames.length).toBeGreaterThan(0);
    });
});