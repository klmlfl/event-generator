'use strict'
const XLSX = require('xlsx')
var SSF = require('ssf')
var fs = require('fs')
var Chance = require('chance')
var chance = new Chance()

function generateCourseSection () {
  var workbook = XLSX.readFile('mock_files/DEAN_import_format.xlsx', { binary: true, cellDates: false, cellStyles: true })
  var courseSectionWorksheet = workbook.SheetNames[workbook.SheetNames.indexOf('Course_Section')]
  var dataWorksheet = workbook.Sheets[courseSectionWorksheet]
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
      value = XLSX.SSF.parse_date_code(dataWorksheet[z].v)
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
          'course_id': data[d].course_id || '',
          'delivery_method': data[d].delivery_method || '',
          'term_code': data[d].term_code || '',
          'start_date': data[d].start_date || '',
          'end_date': data[d].end_date || '',
          'last_day_to_withdraw': data[d].last_day_to_withdraw || '',
          'instructor_id': data[d].instructor_id || '',
          'campus_name': data[d].campus_name || ''
        }

        courseSections.push(courseSection)

        var printStartDate = courseSection.start_date.m + '/' + courseSection.start_date.d + "/" + courseSection.start_date.y
        var printEndDate = courseSection.end_date.m + '/' + courseSection.end_date.d + "/" + courseSection.end_date.y
        var printLastDayToWithdraw = courseSection.last_day_to_withdraw.m + '/' + courseSection.last_day_to_withdraw.d + "/" + courseSection.last_day_to_withdraw.y
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
          course_id: value.course_id,
          delivery_method: value.delivery_method,
          term_code: value.term_code,
          start_date: printStartDate,
          end_date: printEndDate,
          last_day_to_withdraw: printLastDayToWithdraw,
          instructor_id: value.instructor_id,
          campus_name: value.campus_name

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

generateCourseSection()
