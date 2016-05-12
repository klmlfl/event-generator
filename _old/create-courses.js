'use strict'
const XLSX = require('xlsx')
var fs = require('fs')
var Chance = require('chance')
var chance = new Chance()

function generateCourse () {
  var workbook = XLSX.readFile('mock_files/testerys23.xlsx')
  var courseWorksheet = workbook.SheetNames[workbook.SheetNames.indexOf('Course')]
  var dataWorksheet = workbook.Sheets[courseWorksheet]
  var headers = {}
  var data = []
  var courseIds = new Set()
  var courses = []
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
      var courseId = data[d].course_id
      if (!courseIds.has(courseId)) {
        courseIds.add(courseId)

        var course = {
          'course_id': courseId,
          'course_code': data[d].course_code || '',
          'course_name': data[d].course_name || ''
        }

        courses.push(course)
      }
    }
  }

  courses.forEach(function (value) {
    let event = {
      uuid: chance.guid(),
      // time: moment().format('MM/DD/YYYY')
      time: new Date().toJSON(),
      type: 'system.create.course',
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
        type: 'course',
        key: {
          course_id: value.course_id
        },
        val: {
          course_code: value.course_code,
          course_name: value.course_name
        }
      }
    }
    events.push(event)
  })

  // Pretty JSON format
  fs.writeFile('mock_files/output/courses.json', JSON.stringify(events, null, 4), function (err) {
    // fs.writeFile("mock_files/events.json", JSON.stringify(events), function(err) {
    if (err) {
      return console.log(err)
    }
    console.log('Events file saved!')
  })
}

generateCourse()
