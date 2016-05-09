'use strict'
const XLSX = require('xlsx')
var SSF = require('ssf')
var fs = require('fs')
var Chance = require('chance')
var chance = new Chance()

function generateEnrollment () {
  var workbook = XLSX.readFile('mock_files/DEAN_import_format.xlsx', {binary: true, cellDates: false, cellStyles: true})
  var firstWorksheet = workbook.SheetNames[workbook.SheetNames.indexOf('Student_Enrollment')]
  var dataWorksheet = workbook.Sheets[firstWorksheet]
  var headers = {}
  var data = []
  var enrollmentIds = new Set()
  var studentEnrollments = []
  var events = []
  // var timeFrom = "01-01-1900"
  for (let z in dataWorksheet) {
    if (z[0] === '!') continue

    // parse out the column, row, and value
    var col = z.substring(0, 1)
    var row = parseInt(z.substring(1))
    var dateCols = ['enrollment_date', 'enrollment_status_change_date']
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
      var enrollmentId = data[d].enrollment_id
      if (!enrollmentIds.has(enrollmentId)) {
        enrollmentIds.add(enrollmentId)

        var studentEnrollment = {
          'enrollment_id': enrollmentId,
          'course_section_id': data[d].course_section_id || '',
          'person_id': data[d].person_id || '',
          'enrollment_date': data[d].enrollment_date || '',
          'completion_flag': data[d].completion_flag || '',
          'completion_success_flag': data[d].completion_success_flag || '',
          'withdrawal_flag': data[d].withdrawal_flag || '',
          'drop_flag': data[d].drop_flag || '',
          'enrollment_status_change_date': data[d].enrollment_status_change_date || '',
          'course_grade_number': data[d].course_grade_number || '',
          'course_grade_letter': data[d].course_grade_letter || ''
        }

        studentEnrollments.push(studentEnrollment)

        if (studentEnrollment.enrollment_date != null) {
          var printEnrollmentDate = studentEnrollment.enrollment_date.m + '/' + studentEnrollment.enrollment_date.d + '/' + studentEnrollment.enrollment_date.y
        }
        else
          printEnrollmentDate = ''

        if (studentEnrollment.enrollment_status_change_date.m != null) {
          var printEnrollmentStatusChangeDate = studentEnrollment.enrollment_status_change_date.m + '/' + studentEnrollment.enrollment_status_change_date.d + '/' + studentEnrollment.enrollment_status_change_date.y
        }
        else
          printEnrollmentStatusChangeDate = ''
      }
    }
  }

  studentEnrollments.forEach(function (value) {
    let event = {
      uuid: chance.guid(),
      // time: moment().format('MM/DD/YYYY')
      time: new Date().toJSON(),
      type: 'system.create.studentEnrollment',
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
        type: 'student enrollment',
        key: {
          enrollment_id: value.enrollment_id
        },
        val: {
          course_section_id: value.course_section_id,
          person_id: value.person_id,
          enrollment_date: printEnrollmentDate,
          completion_flag: value.completion_flag,
          completion_success_flag: value.completion_success_flag,
          withdrawal_flag: value.withdrawal_flag,
          drop_flag: value.drop_flag,
          enrollment_status_change_date: printEnrollmentStatusChangeDate,
          course_grade_number: value.course_grade_number,
          course_grade_letter: value.course_grade_letter

        }
      }
    }
    events.push(event)
  })

  // Pretty JSON format
  fs.writeFile('mock_files/output/student_enrollments.json', JSON.stringify(events, null, 4), function (err) {
    // fs.writeFile("mock_files/events.json", JSON.stringify(events), function(err) {
    if (err) {
      return console.log(err)
    }
    console.log('Events file saved!')
  })
}

generateEnrollment()
