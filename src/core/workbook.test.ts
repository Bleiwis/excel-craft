import { ExcelWorkbook } from './workbook';
import * as fs from 'fs';
import * as path from 'path';
import * as fsp from 'fs/promises';
import AdmZip from 'adm-zip';

describe('ExcelWorkbook', () => {
    const testDir = path.join(__dirname, '..', '..', 'test');
    const templatesDir = path.join(testDir, 'templates');
    const outputDir = path.join(testDir, 'output');
    let workbook: ExcelWorkbook;

    beforeAll(async () => {
        // Create test directories if they don't exist
        await fsp.mkdir(templatesDir, { recursive: true });
        await fsp.mkdir(outputDir, { recursive: true });
    });

    beforeEach(() => {
        // Create a test Excel file
        const testFilePath = path.join(templatesDir, 'test.xlsx');
        const zip = new AdmZip();
        
        // Add minimal required files for a valid Excel file
        zip.addFile('_rels/.rels', Buffer.from(`
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
            </Relationships>
        `));
        
        zip.addFile('xl/workbook.xml', Buffer.from(`
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheets>
                    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
                </sheets>
            </workbook>
        `));
        
        zip.addFile('xl/worksheets/sheet1.xml', Buffer.from(`
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <sheetData>
                    <row r="1">
                        <c r="A1" t="s">
                            <v>0</v>
                        </c>
                        <c r="B1" t="s">
                            <v>1</v>
                        </c>
                    </row>
                </sheetData>
            </worksheet>
        `));
        
        zip.addFile('xl/sharedStrings.xml', Buffer.from(`
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
                <si><t>Hello</t></si>
                <si><t>World</t></si>
            </sst>
        `));
        
        zip.writeZip(testFilePath);
        workbook = new ExcelWorkbook(testFilePath);
    });

    afterEach(async () => {
        // Clean up test files
        const files = await fsp.readdir(outputDir);
        for (const file of files) {
            await fsp.unlink(path.join(outputDir, file));
        }
    });

    describe('readWorkbook', () => {
        it('should read workbook successfully', async () => {
            await workbook.readWorkbook();
            expect(workbook.getSheetNames()).toEqual(['Sheet1']);
            expect(workbook.getSheetCount()).toBe(1);
        });
    });

    describe('updateCell', () => {
        it('should update cell value', async () => {
            await workbook.readWorkbook();
            workbook.updateCell('Sheet1', 'A1', 'New Value');
            
            const outputPath = path.join(outputDir, 'output.xlsx');
            await workbook.writeWorkbook(outputPath);
            
            // Verify the updated value
            const outputZip = new AdmZip(outputPath);
            const sheetContent = outputZip.getEntry('xl/worksheets/sheet1.xml')?.getData().toString();
            expect(sheetContent).toContain('<v>New Value</v>');
        });
    });

    describe('writeWorkbook', () => {
        it('should write workbook with macros', async () => {
            // Create a test file with macros
            const macroFilePath = path.join(templatesDir, 'test_macro.xlsm');
            const macroZip = new AdmZip();
            
            // Add the same files as before plus vbaProject.bin
            macroZip.addFile('_rels/.rels', Buffer.from(`
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                    <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="xl/vbaProject.bin"/>
                </Relationships>
            `));
            
            macroZip.addFile('xl/workbook.xml', Buffer.from(`
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                    <sheets>
                        <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
                    </sheets>
                </workbook>
            `));
            
            macroZip.addFile('xl/worksheets/sheet1.xml', Buffer.from(`
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                    <sheetData>
                        <row r="1">
                            <c r="A1" t="s">
                                <v>0</v>
                            </c>
                        </row>
                    </sheetData>
                </worksheet>
            `));
            
            macroZip.addFile('xl/vbaProject.bin', Buffer.from('dummy vba project content'));
            macroZip.writeZip(macroFilePath);

            // Test reading and writing the macro file
            const macroWorkbook = new ExcelWorkbook(macroFilePath);
            await macroWorkbook.readWorkbook();
            
            const outputPath = path.join(outputDir, 'output_macro.xlsm');
            await macroWorkbook.writeWorkbook(outputPath);

            // Verify the macro was preserved
            const outputZip = new AdmZip(outputPath);
            expect(outputZip.getEntry('xl/vbaProject.bin')).toBeDefined();
        });
    });
}); 