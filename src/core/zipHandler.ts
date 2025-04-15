import * as fsp from 'fs/promises';
import * as path from 'path';
import * as os from 'os';
import AdmZip = require('adm-zip');
import { ExcelLogger } from './logger';

export class ZipHandler {
    private logger: ExcelLogger;

    constructor(logger: ExcelLogger) {
        this.logger = logger;
    }

    public async extractWorkbook(filePath: string): Promise<string> {
        try {
            // Create a temporary directory
            const tempDir = path.join(os.tmpdir(), 'excel-macros-lib', Date.now().toString());
            await fsp.mkdir(tempDir, { recursive: true });
            this.logger.log('Created temporary directory', { tempDir });

            // Extract the workbook
            const zip = new AdmZip(filePath);
            zip.extractAllTo(tempDir, true);
            this.logger.log('Extracted workbook to temporary directory');

            return tempDir;
        } catch (error) {
            this.logger.logError(error as Error, 'extractWorkbook');
            throw error;
        }
    }

    public async createWorkbook(sourceDir: string, outputPath: string): Promise<void> {
        try {
            // Create output directory if it doesn't exist
            const outputDir = path.dirname(outputPath);
            await fsp.mkdir(outputDir, { recursive: true });

            // Create a new ZIP file
            const zip = new AdmZip();

            // Add all files from the source directory
            const files = await this.getAllFiles(sourceDir);
            for (const file of files) {
                const relativePath = path.relative(sourceDir, file);
                const content = await fsp.readFile(file);
                zip.addFile(relativePath, content);
            }

            // Ensure required files are present
            const requiredFiles = {
                '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                        <Default Extension="xml" ContentType="application/xml"/>
                        <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                        <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                        <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
                        <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
                    </Types>`,
                'xl/styles.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                        <fonts count="1">
                            <font>
                                <sz val="11"/>
                                <color theme="1"/>
                                <name val="Calibri"/>
                                <family val="2"/>
                            </font>
                        </fonts>
                        <fills count="1">
                            <fill>
                                <patternFill patternType="none"/>
                            </fill>
                        </fills>
                        <borders count="1">
                            <border>
                                <left/>
                                <right/>
                                <top/>
                                <bottom/>
                                <diagonal/>
                            </border>
                        </borders>
                        <cellStyleXfs count="1">
                            <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
                        </cellStyleXfs>
                        <cellXfs count="1">
                            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
                        </cellXfs>
                    </styleSheet>`
            };

            for (const [fileName, content] of Object.entries(requiredFiles)) {
                if (!files.some(file => path.relative(sourceDir, file) === fileName)) {
                    zip.addFile(fileName, Buffer.from(content));
                }
            }

            // Write the ZIP file
            zip.writeZip(outputPath);
            this.logger.log('Created workbook', { outputPath });
        } catch (error) {
            this.logger.logError(error as Error, 'createWorkbook');
            throw error;
        }
    }

    private async getAllFiles(dir: string): Promise<string[]> {
        const files: string[] = [];
        const entries = await fsp.readdir(dir, { withFileTypes: true });

        for (const entry of entries) {
            const fullPath = path.join(dir, entry.name);
            if (entry.isDirectory()) {
                files.push(...(await this.getAllFiles(fullPath)));
            } else {
                files.push(fullPath);
            }
        }

        return files;
    }
}