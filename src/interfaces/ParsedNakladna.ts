export class ParsedNakladna {
    zagalna: Map<string, number>;
    khmilnyk: Map<string, number>;
    koziatyn: Map<string, number>;
    kalynivka: Map<string, number>;
  
    constructor(
      zagalna: Map<string, number>,
      khmilnyk: Map<string, number>,
      koziatyn: Map<string, number>,
      kalynivka: Map<string, number>
    ) {
      this.zagalna = zagalna;
      this.khmilnyk = khmilnyk;
      this.koziatyn = koziatyn;
      this.kalynivka = kalynivka;
    }
  }
  