import { ExcelWorkbook } from '../core/workbook';
import * as fs from 'fs';
import * as path from 'path';

async function main() {
    try {
        const templatePath = path.join(__dirname, '../../test/templates/input.xlsm');
        
        // Verificar que el archivo existe
        if (!fs.existsSync(templatePath)) {
            throw new Error(`No se encontró el archivo de plantilla en: ${templatePath}`);
        }
        
        console.log('Tamaño del archivo:', fs.statSync(templatePath).size, 'bytes');
        
        const workbook = new ExcelWorkbook(templatePath);
        await workbook.readWorkbook();
        
        // Listar todas las hojas disponibles
        const sheetNames = workbook.getSheetNames();
    //    console.log('Hojas disponibles:', sheetNames);

        // Obtener el número total de hojas
        const sheetCount = workbook.getSheetCount();
       // console.log('Número total de hojas:', sheetCount);
        
        // Actualizar celda usando el nombre de la hoja
        workbook.updateCell('Hoja1', 'A14', 'Nuevo Valor');

        workbook.updateCell('Hoja1', 'B14', 'B  Valor');
        workbook.updateCell('Hoja1', 'C14', 'C Value');
        
        // Asegurarse de que el directorio de salida existe
        const outputDir = path.join(__dirname, '../../test/output');
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }
        
        // Usar un nombre diferente para el archivo de salida
        const outputPath = path.join(outputDir, 'modified_template_output.xlsm');
        await workbook.writeWorkbook(outputPath);
        
        console.log('Plantilla modificada guardada en:', outputPath);
        console.log('Tamaño del archivo de salida:', fs.statSync(outputPath).size, 'bytes');
    } catch (error) {
        console.error('Error al procesar la plantilla:', error);
        if (error instanceof Error) {
            console.error('Stack trace:', error.stack);
        }
    }
}

main(); 