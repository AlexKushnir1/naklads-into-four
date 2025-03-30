import * as xlsx from "xlsx";

type OperationType = "приход" | "відход";

/**
 * Вставляє кількості з мапи у відповідний аркуш книги Excel, починаючи з рядка 4 у вибраній колонці
 * @param workbook Відкритий Excel-файл (xlsx.readFile(...))
 * @param dataMap Мапа: код => кількість
 * @param operation "приход" або "відход" — обирає аркуш
 * @param columnLetter Наприклад "C", "D" — назва колонки (рядок завжди з 4)
 */
export function writeToTemplate(
  workbook: xlsx.WorkBook,
  dataMap: Map<string, number>,
  operation: OperationType,
  columnLetter: string
): void {
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
    xlsx.utils.sheet_add_aoa(worksheet, [[quantity]], {origin: cell});
    row++;
  }

//   Оновлюємо межі аркуша (!ref), щоб зміни відображались
  const range = xlsx.utils.decode_range(worksheet["!ref"] || `${columnLetter}4:${columnLetter}${row}`);
  range.e.r = Math.max(range.e.r, row);
  worksheet["!ref"] = xlsx.utils.encode_range(range);
  
}