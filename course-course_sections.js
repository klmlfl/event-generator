'use strict'
const XLSX = require('xlsx')
var moment = require('moment')
var SSF = require ('ssf')
var fs = require('fs')
var Chance = require('chance')
var chance = new Chance()

function generateCourse () {
  var workbook = XLSX.readFile('mock_files/test_worksheets_manual.xlsx', {binary: true, cellDates: false, cellStyles: true})
  var firstWorksheet = workbook.SheetNames[workbook.SheetNames.indexOf('Course_Section')]
  var dataWorksheet = workbook.Sheets[firstWorksheet]
  var headers = {}
  var data = []
  var courseIds = new Set()
  var courseSections = []
  var events = []
  // var timeFrom = "01-01-1900"
  for (let z in dataWorksheet) {
    if (z[0] === '!') continue

    // parse out the column, row, and value
    var col = z.substring(0, 1)
    var row = parseInt(z.substring(1))
    var dateCols = ['start_date', 'end_date', 'last_day_to_withdraw']
    var value

    if (row > 1 && dateCols.indexOf(headers[col]) > -1) {
      value = SSF.parse_date_code(dataWorksheet[z].v)
    } else {
      value = dataWorksheet[z].v
    }

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
    // Skip the header row
    if (d > 1) {
      var courseSectionId = data[d].course_section_id
      if (!courseIds.has(courseSectionId)) {
        courseIds.add(courseSectionId)

        var courseSection = {
          'course_section_id': courseSectionId,
          'course_section_code': data[d].course_section_code || '',
          'term_code': data[d].term_code || '',
          //        'start_date': moment(timeFrom, 'DD-MM-YYYY').add(data[d].start_date, 'days') || '',
          //        'end_date': moment(timeFrom, 'DD-MM-YYYY').add(data[d].end_date, 'days') || '',
          'last_day_to_withdraw': data[d].last_day_to_withdraw || ''
        }

        courseSections.push(courseSection)
      }
    }
  }

  courseSections.forEach(function (value) {
    let event = {
      uuid: chance.guid(),
      // time: moment().format('MM/DD/YYYY')
      time: new Date().toJSON(),
      type: 'system.create.courseSection',
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
        type: 'course section',
        key: {
          course_section_id: value.course_section_id
        },
        val: {
          course_section_code: value.course_section_code,
          term_code: value.term_code,
          start_date: value.start_date,
          end_date: value.end_date,
          last_day_to_withdraw: value.last_day_to_withdraw

        }
      }
    }
    events.push(event)
  })

  // Pretty JSON format
  fs.writeFile('mock_files/output/course_sections.json', JSON.stringify(events, null, 4), function (err) {
    // fs.writeFile("mock_files/events.json", JSON.stringify(events), function(err) {
    if (err) {
      return console.log(err)
    }
    console.log('Events file saved!')
  })
}

generateCourse()
