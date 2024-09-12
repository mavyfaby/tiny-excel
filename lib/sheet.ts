import { EventEmitter } from "node:events";

/**
 * Sheet class
 * @author Maverick Fabroa (mavyfaby)
 */
class Sheet extends EventEmitter {
  private _index: number;
  private _data: Record<string, any>;
  private _stringTable: Record<string, any>;

  /**
   * Sheet constructor
   * @param data XML data in object form
   * @param stringTable XML data in object form
   */
  constructor(index: number, data: Record<string, any>, stringTable: Record<string, any>) {
    super();
    this._index = index;
    this._data = data;
    this._stringTable = stringTable;
  }

  /**
   * Get cell value
   * @param cell ex. A1, B2, C3, D4, etc.
   */
  getCell(cell_name: string): string | null {
    // Check if cell name is valid
    if (!cell_name.match(/^[A-Z]+\d+$/)) {
      throw new Error("Invalid cell name");
    }
    
    // Get worksheet
    const worksheet = this._data['worksheet'];
    // Get sheet data rows
    const sheetData = worksheet['sheetData']['row'];
    // Get shared strings map
    const stringTable = this._stringTable.sst.si.map((si: any) => si.t);

    // For each row in the sheet data
    for (const row of sheetData) {
      // Get cells
      const cells = Array.isArray(row.c) ? row.c : [row.c];

      // For each cell in the row
      for (const _cell of cells) {
        // Get cell reference
        if (_cell['@_r'] === cell_name) { // '@_r' is the cell reference (e.g., A1, B5)
          return stringTable[_cell.v] || null;
        }
      }
    }

    return null;
  }

  /**
   * Set cell value
   * @param cell_name ex. A1, B2, C3, D4, etc.
   * @param value Value to set
   * @param type Type of value (string or formula)
   */
  setCell(cell_name: string, value: string | number, type: "string" | "formula" = "string"): void {
    // Check if cell name is valid
    if (!cell_name.match(/^[A-Z]+\d+$/)) {
      throw new Error("Invalid cell name");
    }

    // Get worksheet
    const worksheet = this._data['worksheet'];
    // Get sheet data rows
    const sheetData = worksheet['sheetData']['row'];

    // Find target row
    let target = sheetData.find((row: any) => {
      const cells = Array.isArray(row.c) ? row.c : [row.c];
      return cells.some((cell: any) => cell['@_r'] === cell_name);
    });

    // If target cell is NOT found
    if (!target) {
      // Create new row
      const row = parseInt(cell_name.match(/\d+/g)![0]);
      // Set target to new row
      target = { "@_r": row, c: [] };
      // Add new row to sheet data
      sheetData.push(target);
    }

    // Get target cell
    let targetCell = target.c.find((cell: any) => cell['@_r'] === cell_name);

    // If target cell is NOT found
    if (!targetCell) {
      // Create new cell
      targetCell = { "@_r": cell_name };
      // Add new cell to target row
      target.c.push(targetCell);
    }

    // Get current cell value
    const currentValue = this.getCell(cell_name) || "";

    // If value is a formula
    if (type === "formula") {
      targetCell.f = value;
      delete targetCell.v;
      delete targetCell["@_t"];
    }

    // If value is a number
    else if (typeof value === "number") {
      targetCell.v = value.toString();
      delete targetCell["@_t"];
    }
    
    // If value is a string
    else {
      const stringIndex = this._getSharedString(currentValue, value);
      targetCell.v = stringIndex;
      targetCell["@_t"] = "s";
    }

    // Emit change event
    this.emit("change", {
      data: this._data,
      stringTable: this._stringTable,
      index: this._index
    });
  }

  // Function to add a string to sharedStrings.xml if it doesn't exist
  private _getSharedString(old_value: string, value: string): number {
    // Get shared strings map
    const stringTable = this._stringTable.sst.si.map((si: any) => si.t);
    // Find index of value in shared
    const index = stringTable.indexOf(old_value);

    // If value is found
    if (index !== -1) {
      // Update string value
      this._stringTable.sst.si[index].t = value;
      // Return index of existing string
      return index;
    } 

    // Otherwise, add new string to sharedStrings.xml
    stringTable.push(value);
    // Return index of new string
    this._stringTable.sst.si.push({ t: value });
    // Return index of new string
    return stringTable.length - 1;
  }
}

export default Sheet;