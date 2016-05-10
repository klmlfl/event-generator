'use strict'
const XLSX = require('xlsx')

var workbook = XLSX.readFile('mock_files/test.xlsx')
var firstWorksheet = workbook.SheetNames[0]
var dataWorksheet = workbook.Sheets[firstWorksheet]
var headers = {}

var data = []
var persons = []
var courses = []
var courseSections = []
var studentEnrollments = []

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
  wb.SheetNames.push('Person', 'Course', 'Course_Section', 'Student_Enrollment')
  wb.Sheets['Person'] = ws

  // write file
  XLSX.writeFile(wb, 'mock_files/output/test_worksheetsXXX.xlsx')
}

function extractData () {
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
}


function createPersons () {
  var personIds = new Set()
  for (let d in data) {
    if (d > 1) {
      var personId = data[d].syStudentID
      if (!personIds.has(personId)) {
        personIds.add(personId)

        var person = {
          'id': personId,
          'username': data[d].StudentId || '',
          'email': data[d].StudentEmail || '',
          'givenName': data[d].StudentName.split(',')[1].trim() || '',
          'middleName': data[d].StudentName.split(',')[2] || '',
          'familyName': data[d].StudentName.split(',')[0]
        }
        persons.push(person)
      }
    }
  }
}

function createCourses () {
  var courseIds = new Set()
  for (let d in data) {
    if (d > 1) {
      var courseId = data[d].code
      if (!courseIds.has(courseId)) {
        courseIds.add(courseId)

        var course = {
          'course_id': courseId,
          'course_code': data[d].code || '',
          'course_name': data[d].CourseDescrip || '',
        }
        courses.push(course)
      }
    }
  }
}

function createCourseSections () {
  var courseSectionIds = new Set()
  for (let d in data) {
    if (d > 1) {
      var courseSectionId = data[d].AdClassSchedID
      if (!courseSectionIds.has(courseSectionId)) {
        courseSectionIds.add(courseSectionId)

        var courseSection = {
          'course_section_id': courseSectionId,
          'course_section_code': data[d].Section || '',
          'course_id': data[d].AdClassSchedID || '',
          'delivery_method': data[d].DeliveryMethod || '',
          'term_code': data[d].TermCode || '',
          'start_date': data[d].ClassSchedStart || '',
          'end_date': data[d].ClassSchedEnd || '',
          'last_day_to_withdraw': '',
          'instructor_id': 10001,
          'campus_name': data[d].CampusName

        }
        courseSections.push(courseSection)
      }
    }
  }
}

function createStudentEnrollments () {
  var studentEnrollmentIds = new Set()
  for (let d in data) {
    if (d > 1) {
      var studentEnrollmentId = data[d].AdEnrollSchedID
      if (!studentEnrollmentIds.has(studentEnrollmentId)) {
        studentEnrollmentIds.add(studentEnrollmentId)

        var studentEnrollment = {
          'enrollment_id': studentEnrollmentId,
          'course_section_id': data[d].course_section_id || '',
          'person_id': data[d].person_id || '',
          'enrollment_date': data[d].enrollment_date || '',
          'completion_flag': data[d].completion_flag || '',
          'completion_success_flag': data[d].completion_success_flag || '',
          'withdrawal_flag': data[d].withdrawal_flag || '',
          'drop_flag': data[d].drop_flag || '',
          'enrollment_status_change_date': data[d].enrollment_status_change_date || '',
          'course_grade_number': data[d].course_grade_number || '',
          'course_grade_letter': data[d].adGradeLetterCode || ''
        }
        studentEnrollments.push(studentEnrollment)
      }
    }
  }
}

createWorksheets()
extractData()
createPersons()
createCourses()
createCourseSections()
createStudentEnrollments()

console.log(persons)
