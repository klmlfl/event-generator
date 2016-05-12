'use strict'
const XLSX = require('xlsx')
var fs = require('fs')
var Chance = require('chance')
var chance = new Chance()

function generatePerson () {
  var workbook = XLSX.readFile('mock_files/test_worksheets_manual.xlsx')
  var personWorksheet = workbook.SheetNames[workbook.SheetNames.indexOf('Person')]
  var dataWorksheet = workbook.Sheets[personWorksheet]
  var headers = {}
  var data = []
  var personIds = new Set()
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
      var personId = data[d].person_id
      if (!personIds.has(personId)) {
        personIds.add(personId)

        var person = {
          'person_id': personId,
          'username': data[d].username || '',
          'email': data[d].email || '',
          'given_name': data[d].given_name || '',
          'middle_name': data[d].middle_name || '',
          'family_name': data[d].family_name || ''
        }

        persons.push(person)
      }
    }
  }

  persons.forEach(function (value) {
    let event = {
      uuid: chance.guid(),
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
        type: 'person',
        key: {
          person_id: value.person_id
        },
        val: {
          username: value.username,
          email: value.email,
          given_name: value.given_name,
          middle_name: value.middle_name,
          family_name: value.family_name
        }
      }
    }
    events.push(event)
  })

  // Pretty JSON format
  fs.writeFile('mock_files/output/persons.json', JSON.stringify(events, null, 4), function (err) {
    // fs.writeFile("mock_files/events.json", JSON.stringify(events), function(err) {
    if (err) {
      return console.log(err)
    }
    console.log('Events file saved!')
  })
}

generatePerson()
