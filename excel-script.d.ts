declare namespace ExcelScript {
  interface Workbook {
    getApplication(): Application;
    getWorksheet(name: string): Worksheet;
  }

  interface Application {
    setCalculationMode(mode: CalculationMode): void;
    calculate(type: CalculationType): void;
  }

  interface Worksheet {
    getRange(address: string): Range;
  }

  interface Range {
    insert(shift: InsertShiftDirection): void;
    copyFrom(sourceRange: Range, copyType: RangeCopyType, skipBlanks: boolean, transpose: boolean): void;
    clear(applyTo: ClearApplyTo): void;
    delete(shift: DeleteShiftDirection): void;
    getValue(): unknown;
    setValue(value: unknown): void;
  }

  enum CalculationMode {
    automatic,
    manual
  }

  enum CalculationType {
    full
  }

  enum InsertShiftDirection {
    down
  }

  enum DeleteShiftDirection {
    up
  }

  enum RangeCopyType {
    all,
    values
  }

  enum ClearApplyTo {
    contents
  }
}
