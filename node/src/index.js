import * as Excel from 'exceljs'
import * as xlsx from 'xlsx'
import * as R from 'ramda'
import { getSheetData, createNewExcelFile } from './helpers'

const {
  sheetData: brokenSheetData,
  headers: brokenSheetHeaders,
} = getSheetData('../KDP_messedup.xlsx', 2)
const {
  sheetData: originalSheetData,
  headers: originalSheetHeaders,
} = getSheetData('../KDP_original.xlsx', 0)
// to limit rows, add { sheetRows: number }

const brokenPropertyIndex = R.indexOf('md5', brokenSheetHeaders)
const originalPropertyIndex = R.indexOf('md5', originalSheetHeaders)

const getMatchingRow = brokenDataRow => {
  return originalSheetData.reduce((matchingData, originalDataRow, i) => {
    if (i === 0) return matchingData
    return R.equals(
      originalDataRow[originalPropertyIndex],
      brokenDataRow[brokenPropertyIndex],
    )
      ? originalDataRow
      : matchingData
  }, {})
}

const newSheet = brokenSheetData.map((data, i) => {
  if (i === 0) return data // first el in array are headers
  let fixedData = data

  const matchingRow = getMatchingRow(data)
  // if oid doesn't match match
  if (!R.equals(data[0], matchingRow[0])) {
    fixedData[0] = matchingRow[0]
  }
  return fixedData
})

createNewExcelFile('KDP_Fixed.xlsx', newSheet)

console.log('====================================')
