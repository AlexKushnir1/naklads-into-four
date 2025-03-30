import { ParsedNakladna } from "./ParsedNakladna";
import { OperationType } from "../units";

export class NakladnaWithMeta {
  parsedNakladna: ParsedNakladna;
  operation: OperationType;
  cell: string;

constructor(
    parsedNakladna: ParsedNakladna,
    operation: OperationType,
    cell: string
) {
    this.parsedNakladna = parsedNakladna;
    this.operation = operation;
    this.cell = cell;
}
}