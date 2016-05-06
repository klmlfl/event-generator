'use strict'
const XLSX = require('xlsx')

function createWorksheets () {
  // set up workbook objects
  var wb = {}
  wb.Sheets = {}
  wb.Props = {}
  wb.SSF = {}
  wb.SheetNames = []

  // create worksheet
  var ws = {}

  // add worksheet to workbook
  wb.SheetNames.push('Person', 'Course', 'Course Section', 'Enrollment')
  wb.Sheets['Person'] = ws

  // write file
  XLSX.writeFile(wb, 'mock_files/output/test_worksheets.xlsx')
}

createWorksheets()


