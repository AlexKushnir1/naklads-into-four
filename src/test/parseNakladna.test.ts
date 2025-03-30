import { parseNakladna } from "../logic/parseNakladna";
import * as path from "path";
import * as xlsx from "xlsx";

describe("parseNakladna", () => {
    const filePath = path.join(__dirname, "../../data/накладна столова 01.12.24.xlsm");
    const workbook = xlsx.readFile(filePath);

    console.log("Workbook:", workbook);
    const parsed = parseNakladna(workbook);

    test("Загальна мапа має містити відомі значення", () => {
        const qty = parsed.zagalna.get("11001");
        expect(qty).toBeDefined();
        expect(typeof qty).toBe("number");
    });

    test("рядки без кількості отримують 0", () => {
        const qty = parsed.khmilnyk.get("11002"); // припустимо, такий код є без кількості
        expect(qty).toBe(0);
    });

    test("рядки без коду ігноруються", () => {
        expect(parsed.zagalna.has("")).toBe(false);
    });
});