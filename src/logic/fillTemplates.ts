import * as xlsx from "xlsx";
import { writeToTemplate } from "./writeToTemplate";
import { NakladnaWithMeta } from "../interfaces/NakladnaWithMeta";
import { AllTemplates } from "../interfaces/AllTemplates";

export function fillTemplatesFromParsed(data: NakladnaWithMeta[]): AllTemplates {
    // Читаємо шаблони один раз
    const zagalna = readWorkbook("data/zagalna.xlsx");
    const khmilnyk = readWorkbook("data/khmilnyk.xlsx");
    const koziatyn = readWorkbook("data/koziatyn.xlsx");
    const kalynivka = readWorkbook("data/kalynivka.xlsx");

    // Вставляємо дані
    data.forEach((item) => {
        writeToTemplate(zagalna, item.parsedNakladna.zagalna, item.operation, item.cell);
        writeToTemplate(khmilnyk, item.parsedNakladna.khmilnyk, item.operation, item.cell);
        writeToTemplate(koziatyn, item.parsedNakladna.koziatyn, item.operation, item.cell);
        writeToTemplate(kalynivka, item.parsedNakladna.kalynivka, item.operation, item.cell);
    });

    // Повертаємо оновлені шаблони
    return new AllTemplates(zagalna, khmilnyk, koziatyn, kalynivka);
}

function readWorkbook(filePath: string): xlsx.WorkBook {
    const workbook = xlsx.readFile(filePath);
    if (!workbook) {
        throw new Error(`Не вдалося прочитати файл: ${filePath}`);
    }
    return workbook;
}
