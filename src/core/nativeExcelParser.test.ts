import { NativeExcelParser } from './nativeExcelParser';
import { ExcelLogger } from './logger';
import * as fs from 'fs';
import * as path from 'path';
import { DOMParser } from 'xmldom';

describe('NativeExcelParser', () => {
    let parser: NativeExcelParser;
    let logger: ExcelLogger;
    const testSheetName = 'Sheet1';
    const testXmlContent = `
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
                <row r="1">
                    <c r="A1" t="s">
                        <v>Test Value</v>
                    </c>
                    <c r="B1" t="n">
                        <v>42</v>
                    </c>
                </row>
            </sheetData>
        </worksheet>
    `;

    beforeEach(() => {
        logger = new ExcelLogger(path.join(__dirname, '..', '..', 'test', 'logs'));
        parser = new NativeExcelParser(logger);
    });

    describe('parseSheet', () => {
        it('should parse sheet XML content correctly', async () => {
            await parser.parseSheet(testXmlContent, testSheetName);
            
            const sheetNames = parser.getSheetNames();
            expect(sheetNames).toContain(testSheetName);
            expect(parser.getSheetCount()).toBe(1);
        });

        it('should extract cell values correctly', async () => {
            await parser.parseSheet(testXmlContent, testSheetName);
            
            const sheetXml = parser.getSheetXml(testSheetName);
            expect(sheetXml).toBeDefined();
            
            const doc = new DOMParser().parseFromString(sheetXml!, 'text/xml');
            const cells = doc.getElementsByTagName('c');
            
            expect(cells.length).toBe(2);
            expect(cells[0].getAttribute('r')).toBe('A1');
            expect(cells[1].getAttribute('r')).toBe('B1');
        });
    });

    describe('updateCell', () => {
        it('should update cell value correctly', async () => {
            await parser.parseSheet(testXmlContent, testSheetName);
            
            parser.updateCell(testSheetName, 'A1', 'New Value');
            
            const sheetXml = parser.getSheetXml(testSheetName);
            const doc = new DOMParser().parseFromString(sheetXml!, 'text/xml');
            const cell = doc.getElementsByTagName('c')[0];
            const value = cell.getElementsByTagName('v')[0].textContent;
            
            expect(value).toBe('New Value');
        });

        it('should create new cell if it doesn\'t exist', async () => {
            await parser.parseSheet(testXmlContent, testSheetName);
            
            parser.updateCell(testSheetName, 'C1', 'New Cell');
            
            const sheetXml = parser.getSheetXml(testSheetName);
            const doc = new DOMParser().parseFromString(sheetXml!, 'text/xml');
            const cells = doc.getElementsByTagName('c');
            
            expect(cells.length).toBe(3);
            expect(cells[2].getAttribute('r')).toBe('C1');
            expect(cells[2].getElementsByTagName('v')[0].textContent).toBe('New Cell');
        });
    });

    describe('error handling', () => {
        it('should throw error for invalid sheet name', () => {
            expect(() => {
                parser.updateCell('NonExistentSheet', 'A1', 'Value');
            }).toThrow('Sheet \'NonExistentSheet\' not found');
        });

        it('should throw error for invalid cell reference', async () => {
            // First parse a sheet so it exists
            await parser.parseSheet(testXmlContent, testSheetName);
            
            // Then try to update with invalid cell reference
            expect(() => {
                parser.updateCell(testSheetName, 'InvalidRef', 'Value');
            }).toThrow('Invalid cell reference: InvalidRef');
        });
    });
}); 