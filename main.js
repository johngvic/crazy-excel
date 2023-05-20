const xlsx = require('xlsx');
const fs = require('fs');
const uuid = require('uuid');

const invoke = async () => {
    const encounteredFile = fs.readdirSync('.').find((file) => file.endsWith('.xlsx'));

    if (!encounteredFile) {
        console.log('Could not find XLSX file');
        return;
    }

    console.log(`Found ${encounteredFile}. Start processing...`);

    const inputWorkbook = new xlsx.readFile(encounteredFile, { dense: true });
    const sheets = inputWorkbook.SheetNames;

    for (const sheet of sheets) {
        console.log(`\nProcessing tab sheet "${sheet}"`);
        const outputWorkbook = xlsx.utils.book_new();
        const workbookId = uuid.v4()
        const rows = xlsx.utils.sheet_to_json(inputWorkbook.Sheets[sheet], { header: 1 });
        const rawObjs = [];

        rows.forEach((row, rowIndex) => {
            const auxObj = {};

            row.forEach((element, elementIndex) => {
                element != null ? auxObj[elementIndex] = element : auxObj[elementIndex] = '';
    
                if (rowIndex > 1 && auxObj[elementIndex] == '' && elementIndex > 0) {
                    auxObj[elementIndex] = '-'
                }
            })
    
            if (rowIndex == 1) {
                auxObj[65] = '';
                auxObj[66] = 'Ano';
                auxObj[67] = 'Venda/Devolução';
                auxObj[68] = 'Classificação';
                auxObj[69] = 'Valor Isenção';
                auxObj[70] = 'Valor Redução';
            } else if (rowIndex > 1) {
                const sellOrDevolution = auxObj[56] ? auxObj[56].startsWith('Venda') ? 'Venda' : 'Devolução' : '-';
                const exemptionOrReduction = auxObj[57] ? auxObj[57] == '040' ? 'Isenção' : 'Redução' : '-';
                let exemptionValue = 0;
                let reductionValue = 0;
    
                if (exemptionOrReduction != '-') {
                    if (exemptionOrReduction == 'Isenção') {
                        const exemption = auxObj[54] * 0.18;
                        exemptionValue = exemption.toFixed(2);
                    } else {
                        const reduction = auxObj[64] * (auxObj[59] / 100);
                        reductionValue = reduction.toFixed(2);
                    }
                }

                auxObj[66] = auxObj[22] ? auxObj[22].split('/')[2] : '-'
                auxObj[67] = sellOrDevolution
                auxObj[68] = exemptionOrReduction
                auxObj[69] = exemptionValue
                auxObj[70] = reductionValue
            }
    
            rawObjs.push(auxObj)
        });

        const worksheet = xlsx.utils.json_to_sheet(rawObjs);
        xlsx.utils.book_append_sheet(outputWorkbook, worksheet, sheet);
        xlsx.writeFile(outputWorkbook, `${workbookId}.xlsx`);
    }
}

invoke();