'use strict'
const XLSX = require('xlsx')
const moment = require('moment')
const fs = require('fs')
const Chance = require('chance')
const chance = new Chance()

var workbook = XLSX.readFile('mock_files/input/test_formatted.xlsx')

function createPersons () {
  let personCreationEvents = []
  let roa = XLSX.utils.sheet_to_row_object_array(
      workbook.Sheets[workbook.SheetNames[workbook.SheetNames.indexOf('Person')]])

  for (let row of roa) {
    personCreationEvents.push(
      {
        'uuid': chance.guid(),
        'time': moment().format(),
        'type': 'system.create.person',
        'source': 'lou',
        'subj': {
          'type': 'system',
          'key': {
            'system_id': 'lou'
          }
        },
        'action': {
          'type': 'create',
          'time': moment().format()
        },
        'obj': {
          'type': 'person',
          'key': {
            'person_id': row.person_id || ''
          },
          'val': {
            'username': row.username || '',
            'email': row.email || '',
            'given_name': row.given_name || '',
            'middle_name': row.middle_name || '',
            'family_name': row.family_name || ''
          }
        }
      }
    )
  }

  function writePersons () {
    fs.writeFile('mock_files/output/person-events.json', JSON.stringify(personCreationEvents, null, 4), function (err) {
      if (err) {
        return console.log(err)
      }
      console.log('Person events file saved!')
    })
  }

  writePersons()
}

function createCourses () {
    let courseCreationEvents = []
    let roa = XLSX.utils.sheet_to_row_object_array(
        workbook.Sheets[workbook.SheetNames[workbook.SheetNames.indexOf('Course')]])

  for (let row of roa) {
      courseCreationEvents.push(
        {
          'uuid': chance.guid(),
          'time': moment().format(),
          'type': 'system.create.course',
          'source': 'lou',
          'subj': {
            'type': 'system',
            'key': {
              'system_id': 'lou'
            }
          },
          'action': {
            'type': 'create',
            'time': moment().format()
          },
          'obj': {
            'type': 'course',
            'key': {
              'course_id': row.course_id || ''
            },
            'val': {
              'course_code': row.course_code || '',
              'course_name': row.course_name || ''

            }
          }
        }
      )
    }


  function writeCourses () {
    fs.writeFile('mock_files/output/course-events.json', JSON.stringify(courseCreationEvents, null, 4), function (err) {
      if (err) {
        return console.log(err)
      }
      console.log('Course events file saved!')
    })
  }

  writeCourses()
}

function createCourseSections () {
    let courseSectionCreationEvents = []
    let roa = XLSX.utils.sheet_to_row_object_array(
        workbook.Sheets[workbook.SheetNames[workbook.SheetNames.indexOf('Course_Section')]])

  for (let row of roa) {
      courseSectionCreationEvents.push(
        {
          'uuid': chance.guid(),
          'time': moment().format(),
          'type': 'system.create.course_section',
          'source': 'lou',
          'subj': {
            'type': 'system',
            'key': {
              'system_id': 'lou'
            }
          },
          'action': {
            'type': 'create',
            'time': moment().format()
          },
          'obj': {
            'type': 'course_section',
            'key': {
              'course_section_id': row.course_section_id || ''
            },
            'val': {
              'course_section_code': row.course_section_code || '',
              'course_id': row.course_id || '',
              'delivery_method': row.delivery_method || '',
              'term_code': row.term_code || '',
              'start_date': moment(row.start_date, 'MM-DD-YY').format() || '',
              'end_date': moment(row.end_date, 'MM-DD-YY').format() || '',
              'last_date_to_withdraw': moment(row.last_date_to_withdraw, 'MM-DD-YY').format() || '',
              'instructor_id': row.instructor_id || '',
              'campus_name': row.campus_name || ''

            }
          }
        }
      )
    }

  function writeCourseSections () {
    fs.writeFile('mock_files/output/course_section-events.json', JSON.stringify(courseSectionCreationEvents, null, 4), function (err) {
      if (err) {
        return console.log(err)
      }
      console.log('Course Section events file saved!')
    })
  }

  writeCourseSections()
}

function createStudentEnrollments () {
    let studentEnrollmentEvent = []
    let roa = XLSX.utils.sheet_to_row_object_array(
        workbook.Sheets[workbook.SheetNames[workbook.SheetNames.indexOf('Student_Enrollment')]])

  for (let row of roa) {
      studentEnrollmentEvent.push(
        {
          'uuid': chance.guid(),
          'time': moment().format(),
          'type': 'system.enroll.student',
          'source': 'lou',
          'subj': {
            'type': 'system',
            'key': {
              'system_id': 'lou'
            }
          },
          'action': {
            'type': 'enroll',
            'time': moment().format()
          },
          'obj': {
            'type': 'student_enrollment',
            'key': {
              'enrollment_id': row.enrollment_id || ''
            },
            'val': {
              'course_section_id': row.course_section_id || '',
              'person_id': row.person_id || '',
              'enrollment_date': moment(row.enrollment_date, 'MM-DD-YY').format() || '',
              'completion_flag': row.completion_flag || '',
              'completion_success_flag': row.completion_success_flag || '',
              'withdrawal_flag': row.withdrawal_flag || '',
              'drop_flag': row.drop_flag || '',
              'enrollment_status_change_date': moment(row.enrollment_status_change_date, 'MM-DD-YY').format() || '',
              'course_grade_numeric': row.course_grade_numeric || '',
              'course_grade_letter': row.course_grade_letter || ''

            }
          }
        }
      )
    }
  function writeStudentEnrollments () {
    fs.writeFile('mock_files/output/student_enrollment-events.json', JSON.stringify(studentEnrollmentEvent, null, 4), function (err) {
      if (err) {
        return console.log(err)
      }
      console.log('Student Enrollment events file saved!')
    })
  }
  writeStudentEnrollments()
}


createPersons()
createCourses()
createCourseSections()
createStudentEnrollments()
