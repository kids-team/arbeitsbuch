function main(workbook: ExcelScript.Workbook) {
  const sheets = workbook.getWorksheets();
  
  const currentYear: string = workbook.getWorksheet('Übersicht').getCell(28, 15).getValue().toString();
  const position = ExcelScript.WorksheetPositionType.end
  const calcSheet = workbook.getWorksheet('Rechenblatt');


  if(sheets.length > 3) {
    removeOldYear(workbook)
  }

  
  for (let i=2; i < 54; i++) {
    const template = workbook.getWorksheet('KW 1');
    const newSheet = template.copy(position);
    newSheet.setName(`KW ${i}`)
    newSheet.getCell(0, 15).setValue(i)
    newSheet.getCell(25, 21).setFormula(`='KW ${i-1}'!V32`)
    workbook.getWorksheet('Rechenblatt').getCell(i-1, 1).setFormula(`='KW ${i}'!V35`)
  }

  let loopDate = new Date(`${currentYear}-01-01`)
  let isLeapYear = [2024, 2028, 2032, 2036, 2040, 2044, 2048, 2052].includes(parseInt(currentYear))
  let counter = 1

  do {
    calcSheet.getCell(counter, 3).setValue(loopDate.toLocaleDateString())
    const currentWeek = getWeek(loopDate);
    const currentDay = loopDate.getDay()
    const columns = ['', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP']
    for (let i=1; i<10; i++) {
      calcSheet.getCell(counter, i + 3).setFormula(`='KW ${currentWeek}'!${columns[i]}${currentDay + 5}`)
    }

    if(!isLeapYear && loopDate.getMonth() === 1 && loopDate.getDate() === 28 ) {
      counter +=2
    } else {
      counter++
    }
    
    loopDate.setDate(loopDate.getDate() + 1);
  } while (loopDate.getFullYear().toString() == currentYear)

}

const getWeek = function (today: Date) {
  let firstOfYear = new Date(today.getFullYear(), 0, 1);
  return Math.ceil((((today.getTime() - firstOfYear.getTime()) / (24 * 60 * 60 * 1000))) / 7);
};


const removeOldYear = (workbook: ExcelScript.Workbook) => {
  for (let i = 2; i < 54 + 1; i++) {
    workbook.getWorksheet(`KW ${i}`)?.delete()
  }
}
