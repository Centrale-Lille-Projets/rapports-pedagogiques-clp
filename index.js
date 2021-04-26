const excelToJson = require('convert-excel-to-json');
const fs = require('fs-extra');
const moment = require('moment');
const createReport = require('docx-templates').createReport;
const docxConverter = require('custom-soffice-to-pdf');

const result = excelToJson({
    sourceFile: 'rapports.xlsx'
});

console.log('Liste des pages du Excel:');
keys = Object.keys(result);
for (i in keys) {
    console.log('-' + keys[i]);
}
console.log('\n');

const letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

async function generate() {
    fs.removeSync('result/');
    fs.ensureDirSync('result/');
    // Première page 
    const firstPageTemplate = fs.readFileSync('template-RP.docx');
    const firstPage = result.RP;
    for (i in firstPage) {
        if (i != 0) {

            const real = firstPage[i];

            for (i in letters) {
                const letter = letters[i]
                if (typeof real[letter] == 'undefined') {
                    real[letter] = '';
                }
            }

            real.formattedDate = moment(new Date(real.A)).format("DD/MM/YYYY");

            const filename = `result/${real.B} ${real.C}`;
            const buffer = await createReport({
                template: firstPageTemplate,
                data: real,
                cmdDelimiter: ['{', '}']
            });


            var isWin = process.platform === "win32";
            var customCommand;
            if (isWin) {
                customCommand = '"C:\\Program Files\\LibreOffice\\program\\soffice"';
            } else {
                customCommand = 'soffice';
            }

            const pdfBuffer = await docxConverter(buffer, customCommand)
            fs.writeFileSync(filename + '.pdf', pdfBuffer);

            // ou remplacer les deux dernières lignes par
            // fs.writeFileSync(filename + '.docx', buffer);
        }
    }

}
generate();