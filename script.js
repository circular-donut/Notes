const ExcelJS = require('exceljs');
const fs = require('fs');
let date;

const workSheetColumns = [
    { header: 'Date', key: 'date', width: 10 },
    { header: 'Tags', key: 'tags', width: 10 },
    { header: 'Notes', key: 'notes', width: 10}
  ];

const workbook = new ExcelJS.Workbook();
let worksheet;
workbook.creator = 'Me';
workbook.lastModifiedBy = 'Me';
workbook.created = new Date(1985, 8, 30);
workbook.modified = new Date();
workbook.lastPrinted = new Date(2016, 9, 27);

const main = async (cliArgs) => {
    date = cliArgs[2] + '';
    const filePath = cliArgs[3] + '';

    worksheet = workbook.addWorksheet('generated notes');
    initColumns();
    const text = fs.readFileSync(filePath, 'utf-8');

    const textLinesArr = text.split('\n');
    
    handleTextLines(textLinesArr)
    await workbook.xlsx.writeFile(`./auto_notes`);
}

const initColumns = () => {
    worksheet.columns = workSheetColumns
}

const handleTextLines = (textLinesArr) => {
    let cellXIndex = 1;
    let cellYIndex = 2;
    let currentNoteString = ''
    textLinesArr.forEach(
        textLine => {
            const splitAsterisk = textLine.split('*')
            const currentIndent = splitAsterisk.length >= 2 ? splitAsterisk[0].length : 0

            if(currentIndent === 0){
                writeCell(currentNoteString, cellXIndex, cellYIndex);

                cellYIndex++
                cellXIndex = 1;
                currentNoteString = textLine
                writeCell(currentNoteString, cellXIndex, cellYIndex);
            } else {
                if(cellXIndex === 1){
                    cellXIndex = 2;
                    currentNoteString = textLine + '\n'
                } else{
                    currentNoteString = (currentNoteString + textLine + '\n')
                }
            }
            
        }
    )
    writeCell(currentNoteString, cellXIndex, cellYIndex);
}

const writeCell = (content, x, y) => {
    const row = worksheet.getRow(y);
    let xOffset = 1

    if(x === 1){
        const dateCell = row.getCell(1);
        dateCell.value = date;
    }
    
    const cell = row.getCell(x + xOffset);
    cell.value = content
}

main(process.argv)