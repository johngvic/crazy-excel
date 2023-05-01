const { getXlsxStream } = require('xlstream')
const xlsx = require('xlsx');
const fs = require('fs');

(async () => {
    const outputWorkbook = xlsx.utils.book_new();
    const encounteredFile = fs.readdirSync('.').find((file) => file.endsWith('.xlsx'));

    if (!encounteredFile) {
        console.log('Could not find XLSX file');
        return;
    }

    console.log(`Found ${encounteredFile}. Start processing...`);

    for (let i = 0; i < 10; i++) {
        console.log('Reading tab ' + i)
        
        try {
            const rows = [];
            const stream = await getXlsxStream({
                filePath: `./${encounteredFile}`,
                sheet: i,
            });

            stream
                .on('data', (row) => rows.push(row.raw.arr))
                .on('end', () => {
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
                            auxObj[66] = '';
                            auxObj[67] = 'Ano';
                            auxObj[68] = 'Venda/Devolução';
                            auxObj[69] = 'Classificação';
                            auxObj[70] = 'Valor Isenção';
                            auxObj[71] = 'Valor Redução';
                        } else if (rowIndex > 1) {
                            const sellOrDevolution = auxObj[57] ? auxObj[57].startsWith('Venda') ? 'Venda' : 'Devolução' : '-';
                            const exemptionOrReduction = auxObj[58] ? auxObj[58] == '040' ? 'Isenção' : 'Redução' : '-';
                            let exemptionValue = 0;
                            let reductionValue = 0;
        
                            if (exemptionOrReduction != '-') {
                                if (exemptionOrReduction == 'Isenção') {
                                    const exemption = auxObj[55] * 0.18;
                                    exemptionValue = exemption.toFixed(2);
                                } else {
                                    const reduction = auxObj[65] * (auxObj[60] / 100);
                                    reductionValue = reduction.toFixed(2);
                                }
                            }
        
                            auxObj[67] = auxObj[23] ? auxObj[23].split('/')[2] : '-'
                            auxObj[68] = sellOrDevolution
                            auxObj[69] = exemptionOrReduction
                            auxObj[70] = exemptionValue
                            auxObj[71] = reductionValue
                        }
        
                        rawObjs.push(auxObj)
                    });
        
                    const worksheet = xlsx.utils.json_to_sheet(rawObjs);
                    xlsx.utils.book_append_sheet(outputWorkbook, worksheet, `Sheet ${i}`);
                    xlsx.writeFile(outputWorkbook, "OutputSheetDef.xlsx");
                });
        } catch {
            break;
        }
    }
})();