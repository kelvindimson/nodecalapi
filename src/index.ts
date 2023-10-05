import express from 'express';
import { CellErrorValue, CellFormulaValue, CellHyperlinkValue, CellRichTextValue, CellSharedFormulaValue, Workbook, Worksheet } from 'exceljs';
import path from 'path';

const FormulaParser = require('hot-formula-parser').Parser;
const parser = new FormulaParser();

const app = express();
const PORT = 8000;

app.use(express.json());

function getCellResult(worksheet: Worksheet, cellLabel: string) {
    if (worksheet.getCell(cellLabel).formula) {
        return parser.parse(worksheet.getCell(cellLabel).formula).result;
    } else {
        return worksheet.getCell(cellLabel).value;
    }
}

app.post('/write-excel', async (req, res) => {
    // destructure revenue of type number and expense of type number from req.body
    const { revenue, expense } = req.body as { revenue: number; expense: number; };

    const workbook = new Workbook();
    await workbook.xlsx.readFile(path.join(__dirname, 'file', 'excelfile.xlsx'));

    const worksheet = workbook.getWorksheet(1);

    parser.on('callCellValue', function(cellCoord: { label: string | number; }, done: (arg0: any) => void) {
        if (worksheet.getCell(cellCoord.label).formula) {
            done(parser.parse(worksheet.getCell(cellCoord.label).formula).result);
        } else {
            done(worksheet.getCell(cellCoord.label).value);
        }
    });

    // ... (other parser events, like callRangeValue)

    worksheet.getCell('A1').value = `${revenue}`;
    worksheet.getCell('A2').value = `${expense}`;

    await workbook.xlsx.writeFile(path.join(__dirname, 'file', 'excelfile.xlsx'));

    const revenuefromexcel = worksheet.getCell('A1').value;
    const expensefromexcel = worksheet.getCell('A2').value;

    // Use the getCellResult function to obtain the computed result for cell C1
    const profitfromexcel = getCellResult(worksheet, 'C1');

    res.json({
        revenue: `${revenuefromexcel}`,
        expense: `${expensefromexcel}`,
        profit: profitfromexcel
    });
});

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});