import { parseString, Builder } from 'xml2js';
import { promisify } from 'util';
import { ExcelLogger } from './logger';

const parseStringAsync = promisify(parseString);

export interface XmlNode {
    tag: string;
    attributes: Record<string, string>;
    children: XmlNode[];
    text?: string;
}

export class XmlParser {
    private logger: ExcelLogger;

    constructor(logger: ExcelLogger) {
        this.logger = logger;
    }

    /**
     * Parses XML content into a tree structure
     * @param xml XML string to parse
     * @returns Promise with the root node of the XML tree
     */
    public async parseXml(xml: string): Promise<any> {
        try {
            const result = await parseStringAsync(xml, {
                explicitArray: false,
                mergeAttrs: true,
                explicitRoot: true,
                preserveChildrenOrder: true,
                strict: true,
                trim: true,
                normalize: true,
                normalizeTags: false,
                explicitChildren: false,
                emptyTag: null,
                ignoreAttrs: false,
                explicitCharkey: false,
                attrkey: '$',
                charkey: '_',
                includeWhiteChars: false,
                async: true
            });
            return result;
        } catch (error) {
            this.logger.logError(error as Error, 'parseXml');
            throw error;
        }
    }

    public buildXml(obj: any): string {
        try {
            const builder = new Builder({
                renderOpts: { pretty: false },
                xmldec: { version: '1.0', encoding: 'UTF-8', standalone: true },
                doctype: { sysID: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main' },
                headless: false,
                allowSurrogateChars: false,
                cdata: false
            });
            return builder.buildObject(obj);
        } catch (error) {
            this.logger.logError(error as Error, 'buildXml');
            throw error;
        }
    }

    /**
     * Finds nodes by tag name (with or without namespace)
     * @param node Root node to search from
     * @param tag Tag name to search for (with or without namespace)
     * @returns Array of matching nodes
     */
    static findNodesByTag(node: XmlNode, tag: string): XmlNode[] {
        const nodes: XmlNode[] = [];
        
        // Check both with and without namespace
        const checkTag = (nodeTag: string) => {
            return nodeTag === tag || 
                   nodeTag === `x_${tag}` || 
                   nodeTag === `ns_${tag}` ||
                   nodeTag.endsWith(`_${tag}`);
        };

        if (checkTag(node.tag)) {
            nodes.push(node);
        }

        node.children.forEach(child => {
            nodes.push(...this.findNodesByTag(child, tag));
        });

        return nodes;
    }

    /**
     * Gets the text content of a node
     * @param node Node to get text from
     * @returns Text content or empty string
     */
    static getNodeText(node: XmlNode): string {
        return node.text || '';
    }

    /**
     * Gets an attribute value from a node
     * @param node Node to get attribute from
     * @param name Attribute name (with or without namespace)
     * @returns Attribute value or undefined
     */
    static getAttribute(node: XmlNode, name: string): string | undefined {
        // Check both with and without namespace
        return node.attributes[name] || 
               node.attributes[`x_${name}`] || 
               node.attributes[`ns_${name}`] ||
               node.attributes[`r_${name}`];
    }
}