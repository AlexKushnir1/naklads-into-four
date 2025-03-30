"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.NakladnaWithMeta = void 0;
class NakladnaWithMeta {
    parsedNakladna;
    operation;
    cell;
    constructor(parsedNakladna, operation, cell) {
        this.parsedNakladna = parsedNakladna;
        this.operation = operation;
        this.cell = cell;
    }
}
exports.NakladnaWithMeta = NakladnaWithMeta;
