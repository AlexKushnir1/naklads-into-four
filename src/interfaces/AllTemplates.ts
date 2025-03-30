import * as xlsx from "xlsx";

export class AllTemplates {
    zagalna: xlsx.WorkBook;
    khmilnyk: xlsx.WorkBook;
    koziatyn: xlsx.WorkBook;
    kalynivka: xlsx.WorkBook;

    constructor(zagalna: xlsx.WorkBook, khmilnyk: xlsx.WorkBook, koziatyn: xlsx.WorkBook, kalynivka: xlsx.WorkBook) {
        this.zagalna = zagalna;
        this.khmilnyk = khmilnyk;
        this.koziatyn = koziatyn;
        this.kalynivka = kalynivka;
    }
}