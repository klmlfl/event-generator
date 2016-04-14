var XLSX = require('xlsx')
var workbook = XLSX.readFile('test.xlsx')
var first_worksheet = workbook.SheetNames[0]
var data_worksheet = workbook.Sheets[first_worksheet]
var headers = {}
var data = []
var studentIds = new Set()
var events = []
var uuidG = 0

function generateSystemCreateUser () {
  for (let z in data_worksheet) {
    if (z[0] === '!') continue

    // parse out the column, row, and value
    var col = z.substring(0, 1)
    var row = parseInt(z.substring(1))
    var value = data_worksheet[z].v

    if (!data[row]) data[row] = {}
    data[row][headers[col]] = value

    // store header names
    if (row === 1) {
      headers[col] = value
      continue
    }
  }

  // console.log(data)

  for (let d in data) {
    var student_id = data[d].syStudentID
    studentIds.add(student_id)
  }

  studentIds.forEach(function (value) {
    let event = {
      uuid: uuidG,
      time: '',
      type: 'system.create.user',
      source: 'lou',
      objVal: {
        person_id: value,
        username: 'UN' + value,
        email: 'srahman+' + value + '@learningobjects.com',
        given_name: 'FN' + value,
        middle_name: null,
        family_name: 'LN' + value
      },
      objValOld: {}

    }
    events.push(event)
    uuidG++
  })

  console.log(events)
}

generateSystemCreateUser()
