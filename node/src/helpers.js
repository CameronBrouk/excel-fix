import * as xlsx from 'xlsx'
import * as R from 'ramda'

export const getSheetData = (excelFileName, sheetName, params = {}) => {
  const workbook = xlsx.readFile(excelFileName, params)
  const { SheetNames, Sheets } = workbook
  console.log(SheetNames)
  let sheetData = xlsx.utils.sheet_to_json(Sheets[SheetNames[sheetName]], {
    header: 1,
    defval: '',
    blankrows: true,
  })
  const [headers, ...data] = sheetData
  return { sheetData, headers, data }
}

export const createNewExcelFile = (name, newSheet) => {
  const newWorkbook = xlsx.utils.book_new()
  xlsx.utils.book_append_sheet(
    newWorkbook,
    xlsx.utils.json_to_sheet(newSheet),
    '1',
  )
  xlsx.writeFile(newWorkbook, name)
}
