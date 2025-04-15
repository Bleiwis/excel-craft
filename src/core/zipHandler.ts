import * as fs from 'fs';
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