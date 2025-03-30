import * as path from "path";
import * as xlsx from "xlsx";
import { writeToTemplate } from "../logic/writeToTemplate";

describe("writeToTemplate (in-memory)", () => {
  const templatePath = path.join(
    __dirname,
    "../../data/zagalna.xlsx"
  );

  test("вставляє значення у аркуш 'НАДІЙШЛО'", () => {
    const dataMap = new Map<string, number>([
      ["12345", 12.5],
      ["67890", 7.3],
    ]);

    const workbook = xlsx.readFile(templatePath);
    writeToTemplate(workbook, dataMap, "приход", "F");

    const sheet = workbook.Sheets["НАДІЙШЛО"];
    expect(sheet["F4"]?.v).toBe(12.5);
    expect(sheet["F5"]?.v).toBe(7.3);
  });

  test("вставляє значення у аркуш 'ВИБУЛО'", () => {
    const dataMap = new Map<string, number>([
      ["54321", 1.1],
      ["99999", 2.2],
    ]);

    const workbook = xlsx.readFile(templatePath);
    writeToTemplate(workbook, dataMap, "відход", "G");

    const sheet = workbook.Sheets["ВИБУЛО"];
    expect(sheet["G4"]?.v).toBe(1.1);
    expect(sheet["G5"]?.v).toBe(2.2);
  });
});