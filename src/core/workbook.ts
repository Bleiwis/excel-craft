import { ZipHandler } from './zipHandler';
import { XmlParser } from './xmlParser';
import { NativeExcelParser } from './nativeExcelParser';
import * as fs from 'fs';
import * as fsp from 'fs/promises';
import { DOMParser, XMLSerializer } from 'xmldom';
import AdmZip = require('adm-zip');
import path from 'path';
import os from 'os';
import archiver from 'archiver';
import ExcelJS from 'exceljs';
import { ExcelLogger } from './logger';

interface Cell {
    value: string;
    type: string;
    formula?: string;
    style?: string;
}

interface Sheet {
    name: string;
    data: Map<string, Cell>;
}

interface Workbook {
    sheets: Sheet[];
    sharedStrings: string[];
    macros?: {
        hasMacros: boolean;
        vbaProject?: Buffer;
    };
}

export class ExcelWorkbook {
    private filePath: string;
    private sheets: Map<string, Sheet>;
    private sheetIdToName: Map<string, string>;
    private sharedStrings: Map<string, number>;
    private cells: Map<string, string>;
    private originalSheets: Map<string, string>;
    private nativeParser: NativeExcelParser;
    private logger: ExcelLogger;
    private zipHandler: ZipHandler;
    private domParser: DOMParser;
    private sheetInfo: Map<string, { id: string, name: string }>;

    constructor(filePath: string) {
        this.filePath = filePath;
        this.sheets = new Map();
        this.sheetIdToName = new Map();
        this.sharedStrings = new Map();
        this.cells = new Map();
        this.originalSheets = new Map();
        this.logger = new ExcelLogger(path.join(os.tmpdir(), 'excel-macros-lib', 'logs'));
        this.nativeParser = new NativeExcelParser(this.logger);
        this.zipHandler = new ZipHandler(this.logger);
        this.domParser = new DOMParser();
        this.sheetInfo = new Map();
    }

    public async readWorkbook(): Promise<void> {
        try {
            this.logger.log('Starting workbook read', { filePath: this.filePath });
            
            // Extract the workbook contents
            const tempDir = await this.zipHandler.extractWorkbook(this.filePath);
            
            // Read workbook.xml to get sheet names
            const workbookXmlPath = path.join(tempDir, 'xl', 'workbook.xml');
            const workbookContent = await fsp.readFile(workbookXmlPath, 'utf8');
            console.log('Workbook XML content:', workbookContent);
            
            const workbookDoc = this.domParser.parseFromString(workbookContent, 'text/xml');
            const sheets = Array.from(workbookDoc.getElementsByTagName('sheet'));
            
            // Store sheet information
            for (const sheet of sheets) {
                const sheetName = sheet.getAttribute('name');
                const sheetId = sheet.getAttribute('sheetId');
                if (sheetName && sheetId) {
                    this.sheetInfo.set(sheetName, { id: sheetId, name: sheetName });
                }
            }
            
            // Log all sheet names found
            console.log('All sheets found in workbook:');
            for (const [name, info] of this.sheetInfo) {
                console.log(`- Sheet: ${name}, ID: ${info.id}`);
            }
            
            // Read and parse each sheet
            const sheetFiles = await fsp.readdir(path.join(tempDir, 'xl', 'worksheets'));
            console.log('Sheet files found:', sheetFiles);
            
            for (const sheetFile of sheetFiles) {
                if (sheetFile.endsWith('.xml')) {
                    const sheetContent = await fsp.readFile(
                        path.join(tempDir, 'xl', 'worksheets', sheetFile),
                        'utf8'
                    );
                    const sheetDoc = this.domParser.parseFromString(sheetContent, 'text/xml');
                    const sheetName = this.getSheetNameFromWorkbook(sheets, sheetFile);
                    
                    if (sheetName) {
                        console.log('Processing sheet:', { sheetName, sheetFile });
                        await this.nativeParser.parseSheet(sheetContent, sheetName);
                        this.originalSheets.set(sheetName, sheetContent);
                    }
                }
            }

            // Clean up temporary directory
            await fsp.rm(tempDir, { recursive: true, force: true });
            
            this.logger.log('Workbook loaded successfully');
        } catch (error) {
            console.error('Error details:', error);
            this.logger.logError(error as Error, 'readWorkbook');
            throw error;
        }
    }

    private getSheetNameFromWorkbook(sheets: Element[], sheetFile: string): string | null {
        const match = sheetFile.match(/^sheet(\d+)\.xml$/);
        if (!match) return null;
        
        const sheetId = match[1];
        console.log('Processing sheet file:', sheetFile);
        console.log('Looking for sheetId:', sheetId);
        console.log('Available sheets:', sheets.map(s => ({
            name: s.getAttribute('name'),
            id: s.getAttribute('sheetId')
        })));
        
        for (const sheet of sheets) {
            if (sheet.getAttribute('sheetId') === sheetId) {
                const sheetName = sheet.getAttribute('name');
                console.log('Found matching sheet:', sheetName);
                return sheetName;
            }
        }
        return null;
    }

    public updateCell(sheetName: string, cellRef: string, value: string): void {
        try {
            console.log(`ExcelWorkbook: Updating cell ${cellRef} in sheet ${sheetName} with value: ${value}`);
            
            // Verify the sheet exists
            if (!this.sheetInfo.has(sheetName)) {
                throw new Error(`Sheet '${sheetName}' not found in workbook`);
            }
            
            // Update the cell using the native parser
            this.nativeParser.updateCell(sheetName, cellRef, value);
            
            // Verify the update was successful
            const sheetXml = this.nativeParser.getSheetXml(sheetName);
            if (!sheetXml) {
                throw new Error(`Failed to get updated XML for sheet ${sheetName}`);
            }
            
            console.log(`Successfully updated cell ${cellRef} in sheet ${sheetName}`);
        } catch (error) {
            this.logger.logError(error as Error, 'updateCell');
            throw error;
        }
    }

    public async writeWorkbook(outputPath: string): Promise<void> {
        try {
            this.logger.log('Starting workbook write', { outputPath });
            
            // Create a temporary directory
            const tempDir = path.join(os.tmpdir(), 'excel-macros-lib', `temp-${Date.now()}`);
            await fsp.mkdir(tempDir, { recursive: true });
            
            console.log('Extracting original workbook to temporary directory...');
            // Extract the original workbook to the temporary directory
            const originalZip = new AdmZip(this.filePath);
            originalZip.extractAllTo(tempDir, true);
            
            // Get list of all files in the temporary directory
            const getAllFiles = async (dir: string): Promise<string[]> => {
                const entries = await fsp.readdir(dir, { withFileTypes: true });
                const files = await Promise.all(entries.map(entry => {
                    const res = path.resolve(dir, entry.name);
                    return entry.isDirectory() ? getAllFiles(res) : res;
                }));
                return files.flat();
            };
            
            const allFiles = await getAllFiles(tempDir);
            console.log(`Total files in temporary directory: ${allFiles.length}`);
            
            // Modify the sheets that need to be updated
            const sheetNames = this.nativeParser.getSheetNames();
            console.log('Updating sheets:');
            for (const sheetName of sheetNames) {
                const sheetXml = this.nativeParser.getSheetXml(sheetName);
                if (sheetXml) {
                    const sheetNumber = this.getSheetNumberFromName(sheetName);
                    const sheetPath = path.join(tempDir, 'xl', 'worksheets', `sheet${sheetNumber}.xml`);
                    console.log(`- Updating sheet: ${sheetPath}`);
                    await fsp.writeFile(sheetPath, sheetXml);
                }
            }
            
            // Create new ZIP with all files from the temporary directory
            console.log('Creating new ZIP file...');
            const zip = new AdmZip();
            
            // Add all files to the ZIP, maintaining the directory structure
            for (const file of allFiles) {
                const relativePath = path.relative(tempDir, file);
                console.log(`- Adding to ZIP: ${relativePath}`);
                zip.addLocalFile(file, path.dirname(relativePath));
            }
            
            // Write the ZIP file
            console.log('Writing final ZIP file...');
            zip.writeZip(outputPath);
            
            // Clean up temporary directory
            console.log('Cleaning up temporary files...');
            await fsp.rm(tempDir, { recursive: true, force: true });
            
            // Verify the output file size
            const inputStats = fs.statSync(this.filePath);
            const outputStats = fs.statSync(outputPath);
            console.log(`Input file size: ${inputStats.size} bytes`);
            console.log(`Output file size: ${outputStats.size} bytes`);
            
            if (outputStats.size < inputStats.size * 0.9) {
                console.warn('Warning: Output file is significantly smaller than input file. Some data might have been lost.');
            } else {
                console.log('File sizes match within acceptable range.');
            }
            
            this.logger.log('Workbook written successfully');
        } catch (error) {
            this.logger.logError(error as Error, 'writeWorkbook');
            throw error;
        }
    }

    private getSheetNumberFromName(sheetName: string): number {
        const sheetInfo = this.sheetInfo.get(sheetName);
        if (!sheetInfo) {
            throw new Error(`Sheet '${sheetName}' not found in workbook`);
        }
        return parseInt(sheetInfo.id);
    }

    public getSheetNames(): string[] {
        return this.nativeParser.getSheetNames();
    }

    public getSheetCount(): number {
        return this.nativeParser.getSheetCount();
    }
} 