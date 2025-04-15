import * as fs from 'fs';
import * as path from 'path';
import { DOMParser, XMLSerializer } from 'xmldom';

// Node type constants
const ELEMENT_NODE = 1;
const TEXT_NODE = 3;

// Custom interface for XML elements
interface XmlElement {
    nodeType: number;
    nodeName: string;
    textContent: string | null;
    attributes?: NamedNodeMap;
    childNodes?: NodeList;
}

export class ExcelLogger {
    private logFile: string;
    private logStream: fs.WriteStream | null = null;
    private xmlSerializer: XMLSerializer;
    private cellChanges: Map<string, { before: any, after: any }> = new Map();

    constructor(logDir: string) {
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        this.logFile = path.join(logDir, `excel-debug-${timestamp}.log`);
        this.xmlSerializer = new XMLSerializer();
        this.initializeLogFile();
    }

    private initializeLogFile() {
        try {
            // Create directory if it doesn't exist
            const logDir = path.dirname(this.logFile);
            if (!fs.existsSync(logDir)) {
                fs.mkdirSync(logDir, { recursive: true });
            }

            // Create or clear the log file
            fs.writeFileSync(this.logFile, '');
            this.logStream = fs.createWriteStream(this.logFile, { flags: 'a' });
            
            this.log('=== Excel Debug Log Started ===');
            this.log(`Timestamp: ${new Date().toISOString()}`);
        } catch (error) {
            console.error('Error initializing log file:', error);
        }
    }

    public log(message: string, data?: any) {
        const timestamp = new Date().toISOString();
        const logEntry = {
            timestamp,
            message,
            data
        };
        console.log(JSON.stringify(logEntry, null, 2));
        this.logStream?.write(JSON.stringify(logEntry) + '\n');
    }

    public logCellChange(sheetName: string, cellRef: string, oldValue: any, newValue: any, attributes: any) {
        const timestamp = new Date().toISOString();
        const cellKey = `${sheetName}!${cellRef}`;
        
        // Store the change for detailed analysis
        this.cellChanges.set(cellKey, {
            before: {
                value: oldValue,
                attributes: attributes
            },
            after: {
                value: newValue,
                attributes: attributes
            }
        });

        const changeEntry = {
            timestamp,
            type: 'cell_change',
            sheet: sheetName,
            cell: cellRef,
            oldValue,
            newValue,
            attributes
        };
        console.log(JSON.stringify(changeEntry, null, 2));
        this.logStream?.write(JSON.stringify(changeEntry) + '\n');
    }

    public logSheetModification(sheetName: string, changes: any, xmlBefore?: Document, xmlAfter?: Document) {
        this.log(`Sheet Modification - ${sheetName}`, {
            changes,
            xmlBefore: xmlBefore ? this.xmlSerializer.serializeToString(xmlBefore) : null,
            xmlAfter: xmlAfter ? this.xmlSerializer.serializeToString(xmlAfter) : null
        });
    }

    public logFileOperation(operation: string, filePath: string, details?: any) {
        this.log(`File Operation - ${operation}: ${filePath}`, details);
    }

    public logError(error: Error, context?: string) {
        const timestamp = new Date().toISOString();
        const errorEntry = {
            timestamp,
            type: 'error',
            context,
            message: error.message,
            stack: error.stack
        };
        console.error(JSON.stringify(errorEntry, null, 2));
        this.logStream?.write(JSON.stringify(errorEntry) + '\n');
    }

    public logXmlStructure(node: Element, message: string, detailed: boolean = false): void {
        try {
            const serializer = new XMLSerializer();
            const xml = serializer.serializeToString(node);
            this.log(message, {
                xml,
                detailed,
                nodeType: node.nodeType,
                nodeName: node.nodeName,
                attributes: detailed ? Array.from(node.attributes).map(attr => ({ name: attr.name, value: attr.value })) : undefined
            });
        } catch (error) {
            this.logError(error as Error, 'logXmlStructure');
        }
    }

    public logCompressionDetails(originalSize: number, compressedSize: number, method: number) {
        this.log('Compression Details', {
            originalSize,
            compressedSize,
            compressionRatio: (compressedSize / originalSize * 100).toFixed(2) + '%',
            method
        });
    }

    public logRepairedRecords(records: any[]): void {
        const timestamp = new Date().toISOString();
        const repairEntry = {
            timestamp,
            type: 'repair_records',
            records: records.map(record => ({
                type: record.type,
                location: record.location,
                details: record.details,
                affectedCells: this.getAffectedCells(record)
            }))
        };
        console.log(JSON.stringify(repairEntry, null, 2));
        this.logStream?.write(JSON.stringify(repairEntry) + '\n');
    }

    private getElementStructure(element: XmlElement): any {
        // Handle text nodes
        if (element.nodeType === TEXT_NODE) {
            return {
                type: 'text',
                value: element.textContent?.trim()
            };
        }

        // Handle element nodes
        if (element.nodeType === ELEMENT_NODE) {
            const structure: any = {
                type: 'element',
                tagName: element.nodeName,
                attributes: {},
                children: []
            };

            // Get all attributes
            if (element.attributes && typeof element.attributes.length === 'number') {
                for (let i = 0; i < element.attributes.length; i++) {
                    const attr = element.attributes[i];
                    if (attr && attr.name && attr.value) {
                        structure.attributes[attr.name] = attr.value;
                    }
                }
            }

            // Get all child nodes
            const childNodes = element.childNodes ? Array.from(element.childNodes) : [];
            for (const child of childNodes) {
                if (child) {
                    structure.children.push(this.getElementStructure(child as unknown as XmlElement));
                }
            }

            return structure;
        }

        // Handle other node types
        return {
            type: 'unknown',
            nodeType: element.nodeType,
            nodeName: element.nodeName
        };
    }

    private getAffectedCells(record: any): any[] {
        const affectedCells: any[] = [];
        
        // Check if the record affects specific cells
        if (record.location && record.location.includes('sheet')) {
            const sheetName = this.getSheetNameFromLocation(record.location);
            if (sheetName) {
                // Find all cell changes for this sheet
                for (const [cellKey, change] of this.cellChanges) {
                    if (cellKey.startsWith(sheetName)) {
                        affectedCells.push({
                            cell: cellKey,
                            before: change.before,
                            after: change.after
                        });
                    }
                }
            }
        }

        return affectedCells;
    }

    private getSheetNameFromLocation(location: string): string | null {
        const match = location.match(/sheet(\d+)\.xml/);
        if (match) {
            return `Sheet${match[1]}`;
        }
        return null;
    }

    public close() {
        if (this.logStream) {
            this.logStream.end();
            this.logStream = null;
        }
    }
} 