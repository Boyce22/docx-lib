const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');
const { PDFDocument } = require('pdf-lib');

const logger = require('bunyan').createLogger({ name: 'ConvertService' });

class ConvertService {
    static sofficePath = `"C:\\Program Files\\LibreOffice\\program\\soffice.exe"`;

    /**
     * Converte múltiplos arquivos .docx para .pdf
     * @param {string[]} inputPaths - Caminhos dos arquivos .docx
     * @param {string} outputDir - Diretório onde os PDFs serão salvos
     * @returns {Promise<string[]>} Caminhos dos PDFs convertidos
     */
    static async convertDocxCollectionToPdf(inputPaths, outputDir) {
        logger.info({ count: inputPaths.length }, 'Iniciando conversão de arquivos DOCX para PDF');

        const convertedFiles = [];

        for (const inputPath of inputPaths) {
            try {
                logger.info({ inputPath }, 'Convertendo arquivo DOCX para PDF');
                const pdfPath = await this._convertDocxToPdf(inputPath, outputDir);
                logger.info({ pdfPath }, 'Conversão concluída com sucesso');
                convertedFiles.push(pdfPath);
            } catch (err) {
                logger.error({ err, inputPath }, 'Erro ao converter DOCX para PDF');
            }
        }

        logger.info({ totalConvertidos: convertedFiles.length }, 'Conversão finalizada');
        return convertedFiles;
    }

    /**
     * Converte um único arquivo .docx para .pdf
     * @private
     */
    static _convertDocxToPdf(inputPath, outputDir) {
        return new Promise((resolve, reject) => {
            const command = `${this.sofficePath} --headless --convert-to pdf --outdir "${outputDir}" "${inputPath}"`;

            exec(command, (error, _, stderr) => {
                if (error) {
                    logger.error({ error, stderr }, 'Erro no LibreOffice durante conversão');
                    return reject(new Error(`Erro ao converter DOCX para PDF: ${stderr}`));
                }

                const fileName = path.basename(inputPath, '.docx') + '.pdf';
                resolve(path.join(outputDir, fileName));
            });
        });
    }

    /**
     * Combina múltiplos arquivos PDF em um único buffer PDF
     * @param {Array<{ pathPdf: string }>} modelos - Lista de caminhos para arquivos PDF
     * @returns {Promise<Buffer>}
     */
    static async combinePdfModels(modelos) {
        logger.info({ total: modelos.length }, 'Iniciando combinação de modelos PDF');

        try {
            const pdfCombinado = await PDFDocument.create();

            for (const { pathPdf } of modelos) {
                logger.info({ pathPdf }, 'Adicionando PDF ao combinado');
                const pdfBytes = await fs.promises.readFile(pathPdf);
                const pdf = await PDFDocument.load(pdfBytes);

                const copiedPages = await pdfCombinado.copyPages(pdf, pdf.getPageIndices());
                copiedPages.forEach((page) => pdfCombinado.addPage(page));
            }

            logger.info('Combinação de PDFs concluída com sucesso');
            return await pdfCombinado.save();
        } catch (error) {
            logger.error({ error }, 'Falha ao combinar modelos PDF');
            throw new Error(`Falha ao combinar modelos PDF: ${error.message}`);
        }
    }
}

module.exports = ConvertService;
