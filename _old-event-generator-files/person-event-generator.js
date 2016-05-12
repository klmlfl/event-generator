'use strict'
const XLSX = require('xlsx')
var fs = require('fs')
var Chance = require('chance')
var chance = new Chance()

function generateSystemCreateUser () {
  var workbook = XLSX.readFile('mock_files/test.xlsx')
  var firstWorksheet = workbook.SheetNames[0]
  var dataWorksheet = workbook.Sheets[firstWorksheet]
  var headers = {}
  var data = []
  var studentIds = new Set()
  var persons = []
  var events = []
  for (let z in dataWorksheet) {
    if (z[0] === '!') continue

    // parse out the column, row, and value
    var col = z.substring(0, 1)
    var row = parseInt(z.substring(1))
    var value = dataWorksheet[z].v

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
    // Skip the first row; header row
    if (d > 1) {
      var studentId = data[d].syStudentID
      if (!studentIds.has(studentId)) {
        studentIds.add(studentId)

        var person = {
          'id': studentId,
          'username': null,
          'email': null,
          'givenName': data[d].StudentName.split(',')[1].trim(),
          'middleName': data[d].StudentName.split(',')[2] || '',
          'familyName': data[d].StudentName.split(',')[0]
        }

        persons.push(person)
      }
    }
  }

  persons.forEach(function (value) {
    let event = {
      uuid: chance.guid(),
      // time: moment().format('MM/DD/YYYY')
      time: new Date().toJSON(),
      type: 'system.create.person',
      source: 'lou',
      subj: {
        type: 'system',
        key: {
          system_id: 'lou'
        }
      },
      action: {
        type: 'create',
        time: new Date().toJSON()
      },
      obj: {
        type: 'student',
        key: {
          person_id: value.id
        },
        val: {
          username: value.username,
          email: value.email,
          given_name: value.givenName,
          middle_name: value.middleName,
          family_name: value.familyName
        }
      }
    }
    if (event.obj.key.person_id % 2 === 0) {
      event.obj.val.middle_name = chance.first()
    }
    events.push(event)
  })

  // Pretty JSON format
  fs.writeFile('mock_files/output/person-events.json', JSON.stringify(events, null, 4), function (err) {
    // fs.writeFile("mock_files/events.json", JSON.stringify(events), function(err) { 
    if (err) {
      return console.log(err)
    }
    console.log('Person events file saved!')
  })
}

generateSystemCreateUser()
