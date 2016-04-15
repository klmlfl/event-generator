'use strict'
const XLSX = require('xlsx')
var fs = require('fs')
var Chance = require ('chance')
var chance = new Chance();

function generateSystemCreateUser () {
  var workbook = XLSX.readFile('test.xlsx')
  var first_worksheet = workbook.SheetNames[0]
  var data_worksheet = workbook.Sheets[first_worksheet]
  var headers = {}
  var data = []
  var studentIds = new Set()
  var events = []
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
      uuid: chance.guid(),
      time: chance.date({year: 2012}),
      type: 'system.create.user',
      source: 'lou',
      objVal: {
        person_id: value,
        username: chance.word() + chance.natural({min:0, max:999999}),
        email: chance.email(),
        given_name: chance.first(),
        middle_name: null,
        family_name: chance.last()
      },
      objValOld: {}

    }
    events.push(event)
  })

  //console.log(events)
  fs.writeFile("temp/events.json", JSON.stringify(events), function(err) {
    if(err) {
      return console.log(err);
    }
    console.log("Events file saved!");
  })
}

generateSystemCreateUser()
