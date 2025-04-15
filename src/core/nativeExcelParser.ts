import { DOMParser, XMLSerializer } from 'xmldom';
import * as fs from 'fs';
import * as fsp from 'fs/promises';
import path from 'path';
import { ExcelLogger } from './logger';

/**
 * Interface representing a cell in an Excel worksheet
 */
interface ExcelCell {
    /** The value of the cell */
    value: string;
    /** The type of the cell (n: number, s: string, b: boolean) */
    type: string;
    /** Optional formula in the cell */
    formula?: string;
    /** Optional style of the cell */
    style?: string;
}

/**
 * Interface representing a sheet in an Excel workbook
 */
interface ExcelSheet {
    /** Name of the sheet */
    name: string;
    /** Map of cell references to cell data */
    data: Map<string, ExcelCell>;
    /** Raw XML content of the sheet */
    xmlContent: string;
}

/**
 * Class for parsing and manipulating Excel files using native methods
 */
export class NativeExcelParser {
    private logger: ExcelLogger;
    private sheets: Map<string, ExcelSheet>;
    private sharedStrings: Map<string, number>;
    private domParser: DOMParser;
    private xmlSerializer: XMLSerializer;

    /**
     * Creates a new instance of NativeExcelParser
     * @param logger - Logger instance for tracking operations
     */
    constructor(logger: ExcelLogger) {
        this.logger = logger;
        this.sheets = new Map();
        this.sharedStrings = new Map();
        this.domParser = new DOMParser();
        this.xmlSerializer = new XMLSerializer();
    }

    /**
     * Parses an Excel sheet from its XML content
     * @param xmlContent - The XML content of the sheet
     * @param sheetName - Name of the sheet to parse
     * @throws Error if the sheet data cannot be parsed
     */
    public async parseSheet(xmlContent: string, sheetName: string): Promise<void> {
        try {
            const doc = this.domParser.parseFromString(xmlContent, 'text/xml');
            const sheetData = doc.getElementsByTagName('sheetData')[0];
            
            if (!sheetData) {
                throw new Error('No sheetData found in XML');
            }

            const sheet: ExcelSheet = {
                name: sheetName,
                data: new Map(),
                xmlContent
            };

            const rows = sheetData.getElementsByTagName('row');
            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                const cells = row.getElementsByTagName('c');
                
                for (let j = 0; j < cells.length; j++) {
                    const cell = cells[j];
                    const cellRef = cell.getAttribute('r');
                    if (!cellRef) continue;

                    const valueElement = cell.getElementsByTagName('v')[0];
                    const formulaElement = cell.getElementsByTagName('f')[0];
                    
                    const cellData: ExcelCell = {
                        value: valueElement ? valueElement.textContent || '' : '',
                        type: cell.getAttribute('t') || 'n',
                        formula: formulaElement ? formulaElement.textContent || '' : undefined
                    };

                    sheet.data.set(cellRef, cellData);
                }
            }

            this.sheets.set(sheetName, sheet);
            this.logger.log(`Successfully parsed sheet: ${sheetName}`);
        } catch (error) {
            this.logger.logError(error as Error, 'parseSheet');
            throw error;
        }
    }

    /**
     * Updates the value of a specific cell in a sheet
     * @param sheetName - Name of the sheet containing the cell
     * @param cellRef - Reference of the cell (e.g., 'A1')
     * @param value - New value to set in the cell
     * @throws Error if the sheet is not found or the cell cannot be updated
     */
    public updateCell(sheetName: string, cellRef: string, value: string): void {
        try {
            const sheet = this.sheets.get(sheetName);
            if (!sheet) {
                throw new Error(`Sheet '${sheetName}' not found`);
            }

            const doc = this.domParser.parseFromString(sheet.xmlContent, 'text/xml');
            const sheetData = doc.getElementsByTagName('sheetData')[0];
            
            if (!sheetData) {
                throw new Error('No sheetData found in XML');
            }

            const cell = this.findOrCreateCell(doc, sheetData, cellRef);
            this.updateCellContent(doc, cell, value);

            // Update the sheet's XML content
            sheet.xmlContent = this.xmlSerializer.serializeToString(doc);
            
            // Update the in-memory data
            const cellData = sheet.data.get(cellRef) || { value: '', type: 's' };
            cellData.value = value;
            sheet.data.set(cellRef, cellData);

            console.log(`Cell ${cellRef} updated successfully in sheet ${sheetName}`);
        } catch (error) {
            this.logger.logError(error as Error, 'updateCell');
            throw error;
        }
    }

    /**
     * Finds an existing cell or creates a new one in the sheet
     * @param doc - The XML document
     * @param sheetData - The sheetData element
     * @param cellRef - Reference of the cell to find or create
     * @returns The found or created cell element
     */
    private findOrCreateCell(doc: Document, sheetData: Element, cellRef: string): Element {
        const [column, row] = this.parseCellRef(cellRef);
        const rowNum = parseInt(row);
        
        // Find or create the row
        let targetRow: Element | null = null;
        const rows = sheetData.getElementsByTagName('row');
        
        for (let i = 0; i < rows.length; i++) {
            const currentRow = rows[i];
            const currentRowNum = parseInt(currentRow.getAttribute('r') || '0');
            if (currentRowNum === rowNum) {
                targetRow = currentRow;
                break;
            }
        }

        if (!targetRow) {
            targetRow = doc.createElement('row');
            targetRow.setAttribute('r', row);
            sheetData.appendChild(targetRow);
        }

        // Find or create the cell
        const cells = targetRow.getElementsByTagName('c');
        for (let i = 0; i < cells.length; i++) {
            const cell = cells[i];
            if (cell.getAttribute('r') === cellRef) {
                return cell;
            }
        }

        // Create new cell if not found
        const newCell = doc.createElement('c');
        newCell.setAttribute('r', cellRef);
        targetRow.appendChild(newCell);
        return newCell;
    }

    /**
     * Updates the content of a cell with a new value
     * @param doc - The XML document
     * @param cell - The cell element to update
     * @param value - The new value to set
     */
    private updateCellContent(doc: Document, cell: Element, value: string): void {
        // Clear existing content
        while (cell.firstChild) {
            cell.removeChild(cell.firstChild);
        }

        // Set cell attributes
        cell.setAttribute('t', 's'); // Set type to string
        cell.setAttribute('s', '1'); // Set style to 1 (general)

        // Create and add value element
        const valueElement = doc.createElement('v');
        valueElement.textContent = value;
        cell.appendChild(valueElement);
    }

    /**
     * Parses a cell reference into column and row components
     * @param cellRef - The cell reference (e.g., 'A1')
     * @returns Tuple containing [column, row]
     * @throws Error if the cell reference is invalid
     */
    private parseCellRef(cellRef: string): [string, string] {
        // Validate cell reference format (e.g., A1, B2, AA1, etc.)
        const match = cellRef.match(/^([A-Z]+)(\d+)$/);
        if (!match) {
            throw new Error(`Invalid cell reference: ${cellRef}`);
        }
        
        const [, column, row] = match;
        
        // Validate column is not empty
        if (!column) {
            throw new Error(`Invalid cell reference: ${cellRef} - Column is required`);
        }
        
        // Validate row is a positive number
        const rowNum = parseInt(row);
        if (isNaN(rowNum) || rowNum <= 0) {
            throw new Error(`Invalid cell reference: ${cellRef} - Row must be a positive number`);
        }
        
        return [column, row];
    }

    /**
     * Gets the names of all sheets in the workbook
     * @returns Array of sheet names
     */
    public getSheetNames(): string[] {
        return Array.from(this.sheets.keys());
    }

    /**
     * Gets the total number of sheets in the workbook
     * @returns Number of sheets
     */
    public getSheetCount(): number {
        return this.sheets.size;
    }

    /**
     * Gets the XML content of a specific sheet
     * @param sheetName - Name of the sheet
     * @returns The XML content of the sheet or undefined if not found
     */
    public getSheetXml(sheetName: string): string | undefined {
        return this.sheets.get(sheetName)?.xmlContent;
    }
} 