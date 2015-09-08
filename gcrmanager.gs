/*
*  Google Classroom Manager
*  Developer: Kanakorn Horsiritham
*    Computer Center, Prince of Songkla University
*    Hat Yai, Songkhla, THAILAND
*  Create Date: 2015-09-08
*  Website: http://sysadmin.psu.ac.th/author/kanakorn-h/
*  Base on original demo code of Classroom API
*  https://developers.google.com/classroom/quickstart/apps-script
*/

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('GCR Manager')
      .addItem('List All Courses', 'listAllCourses')
      .addItem('List by Selected Teachers', 'listCourseByTeacher')
      .addItem('List by Selected Students', 'listCourseByStudent')
      .addToUi();
}

function listAllCourses() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.insertSheet();
  var optionalArgs = { };
  listCourses( optionalArgs, sheet);  
}

function listCourses(optionalArgs, sheet) {
  var response = Classroom.Courses.list(optionalArgs);
  var courses = response.courses;
  if (courses && courses.length > 0) {
    sheet.appendRow(["course.id",
                     "course.name",
                     "course.section",
                     "course.descriptionHeading",
                     "course.description",
                     "course.room",
                     "course.ownerId",                       
                     "course.creationTime",
                     "course.updateTime",
                     "course.enrollmentCode",
                     "course.courseState",
                     "course.alternateLink"
                    ]);    
    for (i = 0; i < courses.length; i++) {
      var course = courses[i];        
      var ownerEmail ="";
      try {
        ownerEmail = Classroom.UserProfiles.get(course.ownerId).emailAddress;
      } catch (e) {}
      sheet.appendRow([course.id,
                       course.name,
                       course.section,
                       course.descriptionHeading,
                       course.description,
                       course.room,
                       ownerEmail,                       
                       course.creationTime,
                       course.updateTime,
                       course.enrollmentCode,
                       course.courseState,
                       course.alternateLink
                      ]);

    }
  } else {
    Logger.log('No courses found.');
  }
}


function listCoursesReport(optionalArgs, sheet) {
  var response = Classroom.Courses.list(optionalArgs);
  var courses = response.courses;
  if (courses && courses.length > 0) {
    sheet.appendRow(["course.id",
                     "course.name",
                     "course.section",
                     "countStudents"
                    ]);    
    for (i = 0; i < courses.length; i++) {
      var course = courses[i];        
      var ownerEmail ="";
      try {
        ownerEmail = Classroom.UserProfiles.get(course.ownerId).emailAddress;
      } catch (e) {}
      
      var countStudents=0;
      try {
        countStudents = Classroom.Courses.Students.list(course.id).students.length;
      } catch (e) {}
      sheet.appendRow([course.id,
                       course.name,
                       course.section,
                       countStudents
                      ]);

    }
  } else {
    Logger.log('No courses found.');
  }
}

function getUserProfiles(userId, sheet){
    var response = Classroom.UserProfiles.get(userId);
        sheet.appendRow(["id"          , response.id]);
        sheet.appendRow(["fullName"    , response.name.fullName]);
        sheet.appendRow(["emailAddress", response.emailAddress]);
        sheet.appendRow(["permissions" , response.permissions.toString()]);
        sheet.appendRow(["photoUrl"    , response.photoUrl]);
}

function listCourseByTeacher(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var selected=sheet.getActiveRange();
  var teachers = [];
  
  for (i=0;i<selected.getHeight();i++){
     for (j=0;j<selected.getWidth();j++){
        if (selected.getValues()[i][j] != "") {
          teachers.push(selected.getValues()[i][j]);        
        }
     }
  }  
  
  for (m=0;m<teachers.length;m++) {
    var optionalArgs = { 
      teacherId: teachers[m]
    };
    var response = Classroom.Courses.list(optionalArgs);
      if (ss.getSheetByName(teachers[m]) == null) {
        var teacherSheet = ss.setActiveSheet(ss.insertSheet(teachers[m]));
        getUserProfiles(teachers[m], teacherSheet);
        listCoursesReport(optionalArgs, teacherSheet);      
      }
  }  
}

function listCourseByStudent(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var selected=sheet.getActiveRange();
  var students = [];
  
  for (i=0;i<selected.getHeight();i++){
     for (j=0;j<selected.getWidth();j++){
        if (selected.getValues()[i][j] != "") {
          students.push(selected.getValues()[i][j]);        
        }
     }
  }  
  
  for (s=0;s<students.length;s++) {
    var optionalArgs = { 
      studentId: students[s]
    };
    var response = Classroom.Courses.list(optionalArgs);
      if (ss.getSheetByName(students[s]) == null) {
        var studentSheet = ss.setActiveSheet(ss.insertSheet(students[s]));
        getUserProfiles(students[s], studentSheet);
        listCoursesReport(optionalArgs, studentSheet);      
      }
  }  
}

