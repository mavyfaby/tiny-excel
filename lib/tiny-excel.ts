import { XMLBuilder, XMLParser } from "fast-xml-parser";
import JSZip from "jszip";
import Bun from "bun";

import Sheet from "./sheet";

type LoadOptions = {
  excludeSheetIndexes?: number[];
}

/**
 * TinyExcel class
 * @author Maverick Fabroa (mavyfaby)
 */
class TinyExcel {
  /**
   * Path of the Excel file
   */
  private _path: string;

  /**
   * Loaded file
   */
  private _file: JSZip;

  /**
   * Copy of the data
   */
  private _sheets = new Map<number, Record<string, any>>();

  /**
   * Shared strings
   */
  private _sharedStrings: Record<string, any>;

  /**
   * TinyExcel constructor
   * @param path Path of the Excel file
   * @author Maverick Fabroa (mavyfaby)
   */
  constructor(path: string) {
    this._path = path;
  }

  /**
   * Load the Excel file
   */
  async load(options?: LoadOptions): Promise<void> {
    return new Promise(async (resolve, reject) => {
      try {
        // Lazily load path
        const file = Bun.file(this._path);
        
        // Check if file exists
        if (!(await file.exists())) {
          return reject(`File ${this._path} does not exist.`);
        }

        // If file is not an Excel file
        if (!this._path.endsWith(".xlsx")) {
          return reject(`File ${this._path} is not an Excel file.`);
        }

        // Get buffer file
        const data = await file.arrayBuffer();
        // Create JSZip instance
        const instance = new JSZip();

        // Load zip data
        this._file = await instance.loadAsync(data);

        // Load all sheets
        const sheet_keys = Object.keys(this._file.files).filter((key) => key.startsWith("xl/worksheets/sheet"));

        // Parse sheet XML data
        const xml = new XMLParser({
          ignoreAttributes: false,
          attributeNamePrefix: "@_",
        });

        // Get shared strings XML file
        const sharedStrings = this._file.file("xl/sharedStrings.xml");

        // If shared strings is not found
        if (!sharedStrings) {
          throw new Error("Shared strings not found.");
        }

        // For each sheet
        for (let i = 0; i < sheet_keys.length; i++) {
          // Skip if sheet index is excluded
          if (options?.excludeSheetIndexes?.includes(i)) {
            continue;
          }

          // Get sheet
          const sheetFile = this._file.file(sheet_keys[i]);

          // If sheet is not found
          if (!sheetFile) {
            throw new Error(`Sheet with index ${i} not found.`);
          }

          // Load sheet
          const data = await sheetFile.async("string");

          // Parse sheet XML data
          const parsed = xml.parse(data);

          // Add to sheets
          this._sheets.set(i, parsed);
        }

        // Parse shared strings XML data
        this._sharedStrings = xml.parse(await sharedStrings.async("string"));
        
        // Resolve
        resolve();
      }

      catch (e) {
        reject(e);
      }
    });
  }

  /**
   * Get sheet by index
   */
  getSheet(index: number): Promise<Sheet> {
    return new Promise(async (resolve, reject) => {
      try {
        // Check if sheet exists
        if (!this._sheets.has(index)) {
          throw new Error(`Sheet with index ${index} not found.`);
        }

        // Create new sheet instance
        const sheet = new Sheet(index, this._sheets.get(index)!, this._sharedStrings);
        // Resolve
        resolve(sheet);
      }

      catch (e) {
        reject(e);
      }
    });
  }

  /**
   * Save the sheet
   */
  async save(): Promise<File> {
    return new Promise(async (resolve, reject) => {
      try {
        const buffer = await this.saveBuffer();

        // Save buffer to file
        const file = new File([buffer], "adasdasdasd.xlsx");

        console.log(file);

        // Resolve
        resolve(file);
      }

      catch (e) {
        reject(e);
      }
    });
  }

  /**
   * Save excel file as buffer
   */
  async saveBuffer(): Promise<Buffer> {
    return new Promise(async (resolve, reject) => {
      try {
        // Create XML builder instance
        const builder = new XMLBuilder({
          ignoreAttributes: false,
          attributeNamePrefix: "@_",
        });

        // For every sheet saved
        for (const [index, sheet] of this._sheets) {
          // Convert data to XML
          const dataXML = builder.build(sheet);
          // Save data to zip
          this._file.file(`xl/worksheets/sheet${index + 1}.xml`, dataXML);
        }

        // Convert string table to XML
        const stringTableXML = builder.build(this._sharedStrings);
        // Save string table to zip
        this._file.file("xl/sharedStrings.xml", stringTableXML);

        // Generate buffer
        const buffer = await this._file.generateAsync({ compression: "DEFLATE", type: "nodebuffer" });

        return resolve(Buffer.from(buffer));
      }

      catch (e) {
        reject(e);
      }
    });
  }
}

export default TinyExcel;
