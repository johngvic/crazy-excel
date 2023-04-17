const xlsx = require('xlsx');
const fs = require('fs');

const invoke = async () => {
    const encounteredFile = fs.readdirSync('.').find((file) => file.endsWith('.xlsx'));

    if (!encounteredFile) {
        console.log('Could not find XLSX file');
        return;
    }

    console.log('Start processing...');

    const inputWorkbook = new xlsx.readFile(encounteredFile, { dense: true });
    const outputWorkbook = xlsx.utils.book_new();
    const sheets = inputWorkbook.SheetNames;

    for (const sheet of sheets) {
        console.log(`Processing tab sheet "${sheet}"`);
        const rows = xlsx.utils.sheet_to_json(inputWorkbook.Sheets[sheet], { header: 1 });
        const rawObjs = [];

        rows.forEach((row, rowIndex) => {
            const auxObj = {}
    
            row.forEach((element, elementIndex) => {
                element != null ? auxObj[elementIndex] = element : auxObj[elementIndex] = ''
    
                if (rowIndex > 1 && auxObj[elementIndex] == '' && elementIndex > 0) {
                    auxObj[elementIndex] = '-'
                }
            })
    
            if (rowIndex == 1) {
                auxObj[67] = ''
                auxObj[68] = 'Ano'
                auxObj[69] = 'Venda/Devolução'
                auxObj[70] = 'Classificação'
                auxObj[71] = 'Valor Isenção'
                auxObj[72] = 'Valor Redução'
            } else if (rowIndex > 1) {
                const sellOrDevolution = auxObj[56].startsWith('Venda') ? 'Venda' : 'Devolução'
                const exemptionOrReduction = auxObj[57] == '040' ? 'Isenção' : 'Redução'
                let exemptionValue = '';
                let reductionValue = '';
    
                if (exemptionOrReduction == 'Isenção') {
                    const exemption = auxObj[54] * 0.18
                    exemptionValue = exemption.toFixed(2)
                } else {
                    const reduction = auxObj[64] * (auxObj[59] / 100)
                    reductionValue = reduction.toFixed(2)
                }
    
                auxObj[68] = auxObj[2].split('/')[2]
                auxObj[69] = sellOrDevolution
                auxObj[70] = exemptionOrReduction
                auxObj[71] = exemptionValue
                auxObj[72] = reductionValue
            }
    
            rawObjs.push(auxObj)
        });

        const worksheet = xlsx.utils.json_to_sheet(rawObjs);
        xlsx.utils.book_append_sheet(outputWorkbook, worksheet, sheet);
    }

    xlsx.writeFile(outputWorkbook, "OutputSheet.xlsx");
}

invoke();