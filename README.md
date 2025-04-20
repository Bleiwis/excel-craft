### excel-craft

## An Open-Source Excel Macros Library

A lightweight Node.js library for reading, manipulating, and creating Excel files (`.xlsx` and `.xlsm`) while preserving macros in `.xlsm` files. Built with minimal dependencies and optimized for performance.

This project was initially created to address specific challenges encountered with other popular libraries like `exceljs` and `xlsx-populate`. While these libraries are excellent and widely used, certain use cases involving macros and large file handling required a more tailored solution. This library aims to complement existing tools by providing a focused approach to Excel file manipulation.

---

## Features

- **Read and write Excel files**: Supports `.xlsx` and `.xlsm` formats.
- **Preserve macros**: Ensures macros in `.xlsm` files are not lost during processing.
- **Minimal dependencies**: Uses mostly Node.js native modules.
- **Stream-based processing**: Handles large files efficiently.
- **Framework-agnostic**: Works seamlessly with Node.js, NestJS, and other frameworks.

---

## Installation

This library is currently not published on npm. You can clone this repository and use it directly in your projects.

```bash
git clone https://github.com/your-username/excel-craft.git
cd excel-craft
npm install
```

---

## Usage

### Basic Example

```typescript
import { ExcelWorkbook } from './src';

async function main() {
    const workbook = new ExcelWorkbook('input.xlsx');
    
    // Read an existing Excel file
    await workbook.readWorkbook();
    
    // Update a cell value
    workbook.updateCell('Sheet1', 'A1', 'Hello, World!');
    
    // Save the modified workbook
    await workbook.writeWorkbook('output.xlsx');
}

main();
```

### Working with Macros

The library automatically detects and preserves macros in `.xlsm` files:

```typescript
const workbook = new ExcelWorkbook('input.xlsm');
await workbook.readWorkbook();

// Macros are automatically preserved when saving
await workbook.writeWorkbook('output.xlsm');
```

---

## API Reference

### `ExcelWorkbook`

#### Methods

- **`readWorkbook(filePath: string): Promise<void>`**
  - Reads an Excel file and loads its contents.
  - **Parameters**:
    - `filePath`: Path to the input Excel file.
  - **Example**:

    ```typescript
    await workbook.readWorkbook('input.xlsx');
    ```

- **`writeWorkbook(filePath: string): Promise<void>`**
  - Writes the workbook to a file.
  - **Parameters**:
    - `filePath`: Path to save the output Excel file.
  - **Example**:

    ```typescript
    await workbook.writeWorkbook('output.xlsx');
    ```

- **`updateCell(sheetName: string, cellRef: string, value: string): void`**
  - Updates a cell value in the specified sheet.
  - **Parameters**:
    - `sheetName`: Name of the sheet to update.
    - `cellRef`: Cell reference (e.g., `A1`).
    - `value`: New value for the cell.
  - **Example**:

    ```typescript
    workbook.updateCell('Sheet1', 'A1', 'New Value');
    ```

- **`getSheetNames(): string[]`**
  - Returns a list of all sheet names in the workbook.
  - **Example**:

    ```typescript
    const sheetNames = workbook.getSheetNames();
    console.log(sheetNames);
    ```

- **`getSheetCount(): number`**
  - Returns the total number of sheets in the workbook.
  - **Example**:

    ```typescript
    const sheetCount = workbook.getSheetCount();
    console.log(sheetCount);
    ```

---

## Understanding Excel Files: Under the Hood

Before diving into using this library, it's helpful to understand how Excel files work internally. Excel files, especially `.xlsx` and `.xlsm`, are essentially ZIP archives containing various XML files that define the workbook's structure, data, and formatting.

### Key Components of an Excel File

1. **[Content_Types].xml**
   - Defines the content types of the files within the archive.
   - Specifies the relationships between different parts of the workbook.

2. **_rels/.rels**
   - Contains relationships between the main parts of the workbook.
   - For example, it links the workbook to its sheets and shared strings.

3. **xl/workbook.xml**
   - The main file that defines the workbook structure.
   - Lists all the sheets in the workbook and their properties.

4. **xl/worksheets/sheetX.xml**
   - Contains the data for each sheet in the workbook.
   - Each cell's value, type, and optional formula are stored here.

5. **xl/sharedStrings.xml**
   - Stores all the text values used in the workbook.
   - Text values are referenced by index in the sheet files to save space.

6. **xl/styles.xml**
   - Defines the styles applied to cells, such as fonts, colors, and borders.

7. **xl/vbaProject.bin** (for `.xlsm` files)
   - Contains the macros (VBA code) embedded in the workbook.

### How This Library Works

This library interacts directly with these internal components to read, update, and write Excel files. For example:
- When you update a cell, the library modifies the corresponding `sheetX.xml` file.
- When you save a workbook, it reassembles the ZIP archive with the updated XML files.

### Why This Knowledge is Important

Understanding the internal structure of Excel files can help you:
- Debug issues when working with Excel files.
- Extend the library to support additional features.
- Optimize performance for large files.

If you're new to this, don't worry! The library abstracts most of these details, but having a basic understanding can be very helpful.

---

## Input and Output Templates

### Input Template

- Place your input Excel file in the `test/templates` directory.
- Example file path: `test/templates/input_template.xlsm`.

### Output Template

- The output file will be saved in the `test/output` directory.
- Example file path: `test/output/modified_template_output.xlsm`.

---

## Known Issues

- **Corrupted Output Files**:
  - When generating the output file, Excel reports the file as corrupted. However, Excel can repair the file automatically upon opening.
  - **Status**: This is a known bug and is being investigated.

---

## Contributing

We welcome contributions from the community! Here's how you can help:

1. Fork the repository and create a new branch for your feature or bug fix.
2. Write clear and concise code with comments where necessary.
3. Add or update tests to cover your changes.
4. Submit a pull request with a detailed description of your changes.

### Guidelines

- Follow the existing code style and structure.
- Ensure all tests pass before submitting a pull request.
- Document any new features or changes in the `README.md`.

---

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.
