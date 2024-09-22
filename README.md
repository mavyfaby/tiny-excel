# Tiny Excel
A promise-based, fast and simple .xlsx file editor using Bun APIs.

## Features

- **Promise-based**
  - Asynchronous methods for streamlined file operations.
- **Lightweight**
  - Built with minimal dependencies, ensuring fast execution.
- **High Performance**
  - Leveraging Bun.js for optimal speed and efficiency.
- **Replace values**
  - Replace a value in the sheet with another value.

## Installation

```bash
bun install tiny-excel
```

## Usage

```ts
import TinyExcel from 'tiny-excel';

// Create a new instance of TinyExcel
const excel = new TinyExcel("/path/to/file.xlsx");

// Lazy load the file
await excel.load();

// Get sheet by index
const sheet = excel.getSheet(0);

// Get value from a cell
const value = sheet.getCell("A1");

// Replace a value in the sheet
sheet.setCell("A1", "Hello, World!");

// Save the file (returns a File object)
const file = await excel.save("file.xlsx");
```

When saving the file, there are two options we can use:

1. **Save file to disk**

```ts
await Bun.write(file);
```

2. **Return the file when using native Bun server**

```ts
Bun.serve({
  port: 3000,
  fetch(request) {
    ...
    
    return new Response(file);
  }
})
```

## API

### TinyExcel class
Represents an Excel file, allowing operations such as loading, retrieving sheets, and saving.

`constructor(path: string)`
- Parameters:
  - `path` *(string)*: The path to the .xlsx file.

`load(options?: LoadOptions): Promise<void>`
Loads the Excel file into memory.

- Parameters:
  - `options` *(LoadOptions)*: Optional parameter to exclude certain sheets by index.

`getSheet(index: number): Sheet`
Retrieves a sheet by index.

- Parameters:
  - `index` *(number)*: The sheet index.

`save(name?: string): Promise<File>`
Saves the current Excel data as a file.

- Parameters:
  - `name` *(string)*: Optional file name.

`saveBuffer(): Promise<Buffer>`
Returns the Excel data as a Buffer.

### Sheet class
Represents a sheet in an Excel file, allowing operations such as getting and setting cell values.

`getCell(cell_name: string): string | null`
Gets the value of a cell.

- Parameters:
  - `cell_name` *(string)*: The cell name (e.g., A1, B2, C3, D4, etc.).

`setCell(cell_name: string, value: string | number, type?: "string" | "formula"): void`
Sets the value of a cell.

- Parameters:
  - `cell_name` *(string)*: The cell name (e.g., A1, B2, C3, D4, etc.).
  - `value` *(string | number)*: The value to set.
  - `type` *(string)*: The type of value (**string** or **formula**).

## TODO

- [ ] Add opt-in support for Node.js runtime.
- [ ] Add benchmarks for performance testing.

## Contributing
Feel free to submit issues or pull requests. For major changes, please open an issue first to discuss what you would like to change.

## License

[MIT](LICENSE) © [Maverick Fabroa](https://mavyfaby.me)
