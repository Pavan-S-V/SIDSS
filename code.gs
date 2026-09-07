const CONFIG = {
  APP_NAME: 'SIDSS',
  VERSION: '1.0.0',

  SHEETS: {
    USERS: 'Users',
    STUDENTS: 'Student_Master',
    ACADEMICS: 'Academics',
    SUBJECT_MARKS: 'Subject_Marks',
    ATTENDANCE: 'Attendance',
    PLACEMENT: 'Placement',
    SKILLS: 'Skills',
    TRAINING: 'Training_Certifications',
    DEPARTMENTS: 'Departments',
    COMPANIES: 'Companies',
    EVENTS: 'Events',
    AUDIT: 'Audit_Log'
  }
};


/* =====================================================
   WEB APP
===================================================== */

function doGet() {

  const template = HtmlService.createTemplateFromFile('Index');

  return template
    .evaluate()
    .setTitle('SIDSS | Student Intelligence & Decision Support System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function include(filename) {

  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();
}


/* =====================================================
   DATABASE
===================================================== */

function getSpreadsheet() {
  const SPREADSHEET_ID = '1Sn6HAHna6xpx4D3ntLZiSu5JRvcR8U-DfLnSI3BU1Og';
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getUserManagementData(user) {

  try {

    // Only Institution Admin can access User Management
    if (
      !user ||
      String(user.role || '')
        .trim()
        .toLowerCase() !== 'institution admin'
    ) {
      throw new Error(
        'Only Institution Admin can access User Management.'
      );
    }


    const spreadsheet =
      getSpreadsheet();


    const sheet =
      spreadsheet.getSheetByName('Users');


    if (!sheet) {
      throw new Error(
        'Users sheet was not found.'
      );
    }


    const lastRow =
      sheet.getLastRow();


    // No users
    if (lastRow < 2) {

      return {
        success: true,
        records: [],
        totalUsers: 0,
        activeUsers: 0,
        inactiveUsers: 0
      };

    }


    /*
      USERS SHEET

      A User_ID
      B Full_Name
      C Email
      D Password
      E Role
      F Department
      G Status
    */

    const values =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          7
        )
        .getValues();


    const records =
      values.map(function(row) {

        return {

          userId:
            String(row[0] || '').trim(),

          fullName:
            String(row[1] || '').trim(),

          email:
            String(row[2] || '').trim(),

          // Password intentionally NOT sent to browser

          role:
            String(row[4] || '').trim(),

          department:
            String(row[5] || '').trim(),

          status:
            String(row[6] || '').trim()

        };

      });


    const activeUsers =
      records.filter(function(record) {

        return record.status
          .toLowerCase() === 'active';

      }).length;


    return {

      success: true,

      records: records,

      totalUsers:
        records.length,

      activeUsers:
        activeUsers,

      inactiveUsers:
        records.length - activeUsers

    };


  } catch (error) {

    console.error(
      'getUserManagementData error:',
      error
    );


    return {

      success: false,

      message:
        error && error.message
          ? error.message
          : 'Unable to load users.'

    };

  }

}

function addNewUser(user, newUser) {

  try {

    // Institution Admin only
    if (
      !user ||
      String(user.role || '')
        .trim()
        .toLowerCase() !== 'institution admin'
    ) {
      throw new Error(
        'Only Institution Admin can add users.'
      );
    }


    if (!newUser) {
      throw new Error(
        'User information was not received.'
      );
    }


    const fullName =
      String(newUser.fullName || '').trim();

    const email =
      String(newUser.email || '')
        .trim()
        .toLowerCase();

    const password =
      String(newUser.password || '').trim();

    const role =
      String(newUser.role || '').trim();

    const department =
      String(newUser.department || '').trim();

    const status =
      String(newUser.status || 'Active').trim();


    if (
      !fullName ||
      !email ||
      !password ||
      !role ||
      !department
    ) {

      throw new Error(
        'Please complete all required fields.'
      );

    }


    const spreadsheet =
      getSpreadsheet();


    const sheet =
      spreadsheet.getSheetByName('Users');


    if (!sheet) {
      throw new Error(
        'Users sheet was not found.'
      );
    }


    const lastRow =
      sheet.getLastRow();


    /*
     * Check existing users and
     * generate next User ID.
     */

    let nextNumber = 1;


    if (lastRow >= 2) {

      const existingData =
        sheet
          .getRange(
            2,
            1,
            lastRow - 1,
            7
          )
          .getValues();


      for (
        let i = 0;
        i < existingData.length;
        i++
      ) {

        const existingId =
          String(
            existingData[i][0] || ''
          ).trim();


        const existingEmail =
          String(
            existingData[i][2] || ''
          )
            .trim()
            .toLowerCase();


        // Prevent duplicate email
        if (existingEmail === email) {

          throw new Error(
            'A user with this email already exists.'
          );

        }


        // Find highest USR number
        const match =
          existingId.match(
            /^USR(\d+)$/i
          );


        if (match) {

          const number =
            Number(match[1]);


          if (
            !isNaN(number) &&
            number >= nextNumber
          ) {

            nextNumber =
              number + 1;

          }

        }

      }

    }


    const userId =
      'USR' +
      String(nextNumber)
        .padStart(3, '0');


    /*
     * Users sheet:
     *
     * A User_ID
     * B Full_Name
     * C Email
     * D Password
     * E Role
     * F Department
     * G Status
     */

    sheet.appendRow([

      userId,
      fullName,
      email,
      password,
      role,
      department,
      status

    ]);


    SpreadsheetApp.flush();

    const auditSheet =
      spreadsheet.getSheetByName(
        'Audit_Log'
      );


    if (auditSheet) {

      const now =
        new Date();


      const logId =
        'LOG' +
        Utilities.formatDate(
          now,
          Session.getScriptTimeZone(),
          'yyyyMMddHHmmss'
        );


      auditSheet.appendRow([

        logId,                           // Log_ID

        now,                             // Timestamp

        String(user.email || ''),        // User_Email

        String(user.role || ''),         // User_Role

        'CREATE',                        // Action

        'Users',                         // Module

        userId,                          // Record_ID

        'Created user ' +
          fullName +
          ' (' + email + ')',            // Description

        'Web App',                       // IP_Address / Source

        'Success'                        // Status

      ]);


      SpreadsheetApp.flush();

    }

    return {

      success: true,

      message:
        'User added successfully.',

      userId:
        userId

    };


  } catch (error) {

    console.error(
      'addNewUser error:',
      error
    );


    return {

      success: false,

      message:
        error && error.message
          ? error.message
          : 'Unable to add user.'

    };

  }

}


function updateUser(user, updatedUser) {

  try {

    // Institution Admin only
    if (
      !user ||
      String(user.role || '')
        .trim()
        .toLowerCase() !== 'institution admin'
    ) {
      throw new Error(
        'Only Institution Admin can update users.'
      );
    }


    if (!updatedUser) {
      throw new Error(
        'Updated user information was not received.'
      );
    }


    const userId =
      String(updatedUser.userId || '')
        .trim()
        .toUpperCase();

    const fullName =
      String(updatedUser.fullName || '')
        .trim();

    const email =
      String(updatedUser.email || '')
        .trim()
        .toLowerCase();

    const role =
      String(updatedUser.role || '')
        .trim();

    const department =
      String(updatedUser.department || '')
        .trim()
        .toUpperCase();

    const status =
      String(updatedUser.status || '')
        .trim();


    if (
      !userId ||
      !fullName ||
      !email ||
      !role ||
      !department ||
      !status
    ) {

      throw new Error(
        'Please complete all required fields.'
      );

    }


    const spreadsheet =
      getSpreadsheet();

    const sheet =
      spreadsheet.getSheetByName('Users');


    if (!sheet) {
      throw new Error(
        'Users sheet was not found.'
      );
    }


    const lastRow =
      sheet.getLastRow();


    if (lastRow < 2) {
      throw new Error(
        'No users were found.'
      );
    }


    const values =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          7
        )
        .getValues();


    let targetRow = -1;


    for (
      let i = 0;
      i < values.length;
      i++
    ) {

      const existingId =
        String(values[i][0] || '')
          .trim()
          .toUpperCase();

      const existingEmail =
        String(values[i][2] || '')
          .trim()
          .toLowerCase();


      // Find selected user
      if (existingId === userId) {

        targetRow = i + 2;

      }


      // Prevent duplicate email
      if (
        existingId !== userId &&
        existingEmail === email
      ) {

        throw new Error(
          'Another user already uses this email.'
        );

      }

    }


    if (targetRow === -1) {

      throw new Error(
        'User ' +
        userId +
        ' was not found.'
      );

    }


    /*
     * IMPORTANT:
     * We update B, C, E, F, G.
     *
     * Password column D is NOT changed.
     */

    sheet
      .getRange(targetRow, 2)
      .setValue(fullName);

    sheet
      .getRange(targetRow, 3)
      .setValue(email);

    sheet
      .getRange(targetRow, 5)
      .setValue(role);

    sheet
      .getRange(targetRow, 6)
      .setValue(department);

    sheet
      .getRange(targetRow, 7)
      .setValue(status);


    SpreadsheetApp.flush();

    const auditSheet =
      spreadsheet.getSheetByName(
        'Audit_Log'
      );


    if (auditSheet) {

      const now =
        new Date();


      const logId =
        'LOG' +
        Utilities.formatDate(
          now,
          Session.getScriptTimeZone(),
          'yyyyMMddHHmmss'
        );


      auditSheet.appendRow([

        logId,                           // Log_ID

        now,                             // Timestamp

        String(user.email || ''),        // User_Email

        String(user.role || ''),         // User_Role

        'UPDATE',                        // Action

        'Users',                         // Module

        userId,                          // Record_ID

        'Updated user ' +
          fullName +
          ' (' + email + ')',            // Description

        'Web App',                       // IP_Address / Source

        'Success'                        // Status

      ]);


      SpreadsheetApp.flush();

    }



    return {

      success: true,

      message:
        'User ' +
        userId +
        ' updated successfully.'

    };


  } catch (error) {

    console.error(
      'updateUser error:',
      error
    );


    return {

      success: false,

      message:
        error && error.message
          ? error.message
          : 'Unable to update user.'

    };

  }

}


function getAuditLogsData(user) {

  try {

    /*
     * INSTITUTION ADMIN ONLY
     */

    if (
      !user ||
      String(user.role || '')
        .trim()
        .toLowerCase() !== 'institution admin'
    ) {

      throw new Error(
        'Only Institution Admin can access Audit Logs.'
      );

    }


    const spreadsheet =
      getSpreadsheet();


    const sheet =
      spreadsheet.getSheetByName(
        'Audit_Log'
      );


    if (!sheet) {

      throw new Error(
        'Audit_Log sheet was not found.'
      );

    }


    const lastRow =
      sheet.getLastRow();


    /*
     * EMPTY AUDIT LOG
     */

    if (lastRow < 2) {

      return {

        success: true,

        totalLogs: 0,
        successLogs: 0,
        failedLogs: 0,

        records: []

      };

    }


    /*
     * READ A:J
     */

    const values =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          10
        )
        .getValues();


    const records =
      values.map(function(row) {

        /*
         * FORMAT TIMESTAMP
         */

        let timestamp = row[1];


        if (
          timestamp instanceof Date
        ) {

          timestamp =
            Utilities.formatDate(
              timestamp,
              Session.getScriptTimeZone(),
              'yyyy-MM-dd HH:mm:ss'
            );

        }

        else {

          timestamp =
            String(
              timestamp || ''
            );

        }


        return {

          logId:
            String(row[0] || ''),

          timestamp:
            timestamp,

          userEmail:
            String(row[2] || ''),

          userRole:
            String(row[3] || ''),

          action:
            String(row[4] || ''),

          module:
            String(row[5] || ''),

          recordId:
            String(row[6] || ''),

          description:
            String(row[7] || ''),

          ipAddress:
            String(row[8] || ''),

          status:
            String(row[9] || '')

        };

      });


    /*
     * NEWEST LOGS FIRST
     */

    records.reverse();


    const successLogs =
      records.filter(function(record) {

        return String(
          record.status || ''
        )
          .trim()
          .toLowerCase() ===
          'success';

      }).length;


    const failedLogs =
      records.filter(function(record) {

        return String(
          record.status || ''
        )
          .trim()
          .toLowerCase() ===
          'failed';

      }).length;


    /*
     * UNIQUE USERS
     */

    const uniqueUsers =
      new Set(
        records
          .map(function(record) {

            return String(
              record.userEmail || ''
            )
              .trim()
              .toLowerCase();

          })
          .filter(Boolean)
      ).size;


    return {

      success: true,

      totalLogs:
        records.length,

      successLogs:
        successLogs,

      failedLogs:
        failedLogs,

      uniqueUsers:
        uniqueUsers,

      records:
        records

    };


  } catch (error) {

    console.error(
      'getAuditLogsData error:',
      error
    );


    return {

      success: false,

      message:
        error && error.message
          ? error.message
          : 'Unable to load Audit Logs.'

    };

  }

}



function getSheetData(sheetName) {

  const sheet = getSpreadsheet().getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('Sheet not found: ' + sheetName);
  }

  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return [];
  }

  const headers = values[0].map(function(header) {
    return String(header).trim();
  });

  return values.slice(1).map(function(row) {

    const obj = {};

    headers.forEach(function(header, index) {
      obj[header] = row[index];
    });

    return obj;
  });
}


/* =====================================================
   LOGIN
===================================================== */

function loginUser(email, password) {

  try {

    email = String(email || '').trim().toLowerCase();
    password = String(password || '').trim();

    if (!email || !password) {

      return {
        success: false,
        message: 'Please enter email and password.'
      };
    }


    const users = getSheetData(CONFIG.SHEETS.USERS);


    const user = users.find(function(item) {

      return (
        String(item.Email || '').trim().toLowerCase() === email &&
        String(item.Password || '').trim() === password
      );

    });


    if (!user) {

      logActivity(
        email,
        'Unknown',
        'LOGIN',
        'Users',
        '',
        'Invalid login attempt',
        'Failed'
      );

      return {
        success: false,
        message: 'Invalid email or password.'
      };
    }


    if (
      String(user.Status || '')
        .trim()
        .toLowerCase() !== 'active'
    ) {

      return {
        success: false,
        message: 'Your account is inactive.'
      };
    }


    const sessionUser = {

      userId: String(user.User_ID || ''),
      facultyId: String(user.User_ID || ''),   // FAC001, FAC002...
      name: String(user.Full_Name || ''),
      email: String(user.Email || ''),
      role: String(user.Role || ''),
      department: String(user.Department || 'ALL'),
      status: String(user.Status || '')

    };


    logActivity(
      sessionUser.email,
      sessionUser.role,
      'LOGIN',
      'Users',
      sessionUser.userId,
      'User logged into SIDSS',
      'Success'
    );


    return {
      success: true,
      message: 'Login successful.',
      user: sessionUser
    };


  } catch (error) {

    return {
      success: false,
      message: error.message
    };
  }
}



function getFacultySubjects(facultyId) {

  const facultyData = getSheetData("Faculty_Subjects");

  const facultyIdText = String(facultyId || "").trim();

  return facultyData
    .filter(function(row) {
      return String(row.Faculty_ID || "").trim() === facultyIdText;
    })
    .map(function(row) {

      return {
        Faculty_ID: row.Faculty_ID,
        Department: row.Department,
        Semester: row.Semester,
        Subject: row.Subject,
        Subject_Code: row.Subject_Code,
        Subject_Name: row.Subject_Name,
        Subject_Category: row.Subject_Category || ""
      };

    });
}


function getFacultyStudents(facultyId) {

  const facultyData =
    getSheetData("Faculty_Subjects");

  const students =
    getSheetData("Student_Master");

  const evaluationData =
    getSheetData("Subject_Evaluation");


  const facultyIdStr =
    String(facultyId || "").trim();


  // ==========================================
  // SUBJECT EVALUATION LOOKUP
  // ==========================================

  const evaluationMap = {};

  evaluationData.forEach(function(row) {

    const code =
      String(row.Subject_Code || "").trim();

    if (!code) return;


    evaluationMap[code] = {

      assignmentMax:
        Number(row.Assignment_Max || 0),

      reportMax:
        Number(row.Report_Max || 0),

      labMax:
        Number(row.Lab_Max || 0),

      internal1Max:
        Number(row.Internal_1_Max || 0),

      internal2Max:
        Number(row.Internal_2_Max || 0),

      externalMax:
        Number(row.External_Max || 0),

      externalReducedMax:
        Number(row.External_Reduced_Max || 0),

      subjectCategory:
        String(row.Subject_Category || "").trim()

    };

  });


  // ==========================================
  // FACULTY SUBJECTS
  // ==========================================

  const facultySubjects =
    facultyData.filter(function(row) {

      return String(row.Faculty_ID || "").trim()
        === facultyIdStr;

    });


  if (facultySubjects.length === 0) {
    return [];
  }


  // ==========================================
  // STUDENT LOOKUP
  // ==========================================

  const studentMap = {};

  students.forEach(function(student) {

    const department =
      String(student.Department || "").trim();

    const semester =
      String(student.Semester || "").trim();

    const key =
      normalizeFacultyDepartment(department) +
      "|" +
      semester;


    if (!studentMap[key]) {
      studentMap[key] = [];
    }

    studentMap[key].push(student);

  });


  // ==========================================
  // BUILD RESULT
  // ==========================================

  const result = [];


  facultySubjects.forEach(function(subject) {

    const department =
      String(subject.Department || "").trim();

    const semester =
      String(subject.Semester || "").trim();

    const subjectCode =
      String(subject.Subject_Code || "").trim();

    const key =
      normalizeFacultyDepartment(department) +
      "|" +
      semester;


    const evaluation =
      evaluationMap[subjectCode] || {

        assignmentMax: 0,
        reportMax: 0,
        labMax: 0,
        internal1Max: 50,
        internal2Max: 50,
        externalMax: 100,
        externalReducedMax: 50,
        subjectCategory: ""

      };


    const matchingStudents =
      studentMap[key] || [];


    matchingStudents.forEach(function(student) {

      result.push({

        usn:
          student.USN,

        studentName:
          student.Student_Name,

        department:
          student.Department,

        semester:
          student.Semester,

        subjectCode:
          subjectCode,

        subjectName:
          subject.Subject_Name || "",


        // ==================================
        // APPLICABILITY
        // ==================================

        showAssignment:
          evaluation.assignmentMax > 0,

        showReport:
          evaluation.reportMax > 0,

        showLab:
          evaluation.labMax > 0,


        // ==================================
        // MAX MARKS
        // ==================================

        assignmentMax:
          evaluation.assignmentMax,

        reportMax:
          evaluation.reportMax,

        labMax:
          evaluation.labMax,

        internal1Max:
          evaluation.internal1Max,

        internal2Max:
          evaluation.internal2Max,

        externalMax:
          evaluation.externalMax,

        externalReducedMax:
          evaluation.externalReducedMax,


        subjectCategory:
          evaluation.subjectCategory

      });

    });

  });


  return result;

}


/* =====================================================
   ROLE CONFIGURATION
===================================================== */

function getRolePermissions(role) {

  const permissions = {

    'Institution Admin': {
      dashboard: true,
      students: true,
      academics: true,
      attendance: true,
      placement: true,
      skills: true,
      training: true,
      departments: true,
      companies: true,
      events: true,
      reports: true,
      users: true,
      audit: true,
      edit: true
    },


    'Principal': {
      dashboard: true,
      students: true,
      academics: true,
      attendance: true,
      placement: true,
      skills: true,
      training: true,
      departments: true,
      companies: true,
      events: true,
      reports: true,
      users: false,
      audit: false,
      edit: false
    },


    'HOD': {
      dashboard: true,
      students: true,
      academics: true,
      attendance: true,
      placement: true,
      skills: true,
      training: true,
      departments: false,
      companies: false,
      events: true,
      reports: true,
      users: false,
      audit: false,
      edit: false
    },


    'Department Admin': {
      dashboard: true,
      students: true,
      academics: true,
      attendance: true,
      placement: true,
      skills: true,
      training: true,
      departments: false,
      companies: false,
      events: true,
      reports: true,
      users: false,
      audit: false,
      edit: true
    },


    'Placement Officer': {
      dashboard: true,
      students: true,
      academics: true,
      attendance: true,
      placement: true,
      skills: true,
      training: true,
      departments: false,
      companies: true,
      events: true,
      reports: true,
      users: false,
      audit: false,
      edit: true
    }

  };


  return permissions[role] || {};
}


/* =====================================================
   AUDIT LOG
===================================================== */

function logActivity(
  email,
  role,
  action,
  module,
  recordId,
  description,
  status
) {

  try {

    const sheet =
      getSpreadsheet().getSheetByName(CONFIG.SHEETS.AUDIT);

    if (!sheet) return;


    const logId =
      'LOG' +
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        'yyyyMMddHHmmss'
      );


    sheet.appendRow([

      logId,

      new Date(),

      email,

      role,

      action,

      module,

      recordId,

      description,

      'Web App',

      status

    ]);

  } catch (error) {

    console.log(error);

  }
}


/* =====================================================
   DATABASE TEST
===================================================== */

function testDatabaseConnection() {

  try {

    const ss = getSpreadsheet();

    return {

      success: true,

      spreadsheet: ss.getName(),

      sheets: ss
        .getSheets()
        .map(function(sheet) {
          return sheet.getName();
        })

    };

  } catch (error) {

    return {
      success: false,
      message: error.message
    };

  }
}

/* =====================================================
   DASHBOARD
===================================================== */

function getDashboardData(user) {

  try {

    if (!user || !user.role) {
      throw new Error('Invalid user session.');
    }


    const students =
      getSheetData(CONFIG.SHEETS.STUDENTS) || [];

    const academics =
      getSheetData(CONFIG.SHEETS.ACADEMICS) || [];

    const attendance =
      getSheetData(CONFIG.SHEETS.ATTENDANCE) || [];

    const placements =
      getSheetData(CONFIG.SHEETS.PLACEMENT) || [];


    const department =
      String(user.department || 'ALL')
        .trim()
        .toUpperCase();


    const restricted =
      user.role === 'HOD' ||
      user.role === 'Department Admin';


    // ============================================
    // STUDENTS
    // ============================================

    const filteredStudents =
      restricted
        ? filterByDepartment(
            students,
            department
          )
        : students;


    // ============================================
    // ALLOWED USNs
    // ============================================

    const allowedUSNs =
      new Set(
        filteredStudents.map(function(student) {

          return String(
            student.USN ||
            student.usn ||
            ''
          )
            .trim()
            .toUpperCase();

        })
      );


    // ============================================
    // FILTER OTHER MODULES
    // ============================================

    function filterByAllowedUSN(records) {

      if (!restricted) {
        return records;
      }

      return records.filter(function(record) {

        const usn =
          String(
            record.USN ||
            record.usn ||
            ''
          )
            .trim()
            .toUpperCase();

        return allowedUSNs.has(usn);

      });

    }


    const filteredAcademics =
      filterByAllowedUSN(academics);


    const filteredAttendance =
      filterByAllowedUSN(attendance);


    const filteredPlacements =
      filterByAllowedUSN(placements);


    // ============================================
    // NORMAL DASHBOARD VALUES
    // ============================================

    const averageCGPA =
      calculateAverageCGPA(
        filteredAcademics
      );


    const averageAttendance =
      calculateAverageAttendance(
        filteredAttendance
      );


    const placedStudents =
      calculatePlacedStudents(
        filteredPlacements
      );


    // ============================================
    // IMPORTANT:
    // USE THE SAME AT-RISK ENGINE
    // ============================================

    const riskResult =
      getAtRiskStudents(user);


    const atRiskCount =
      riskResult &&
      riskResult.success
        ? Number(riskResult.count || 0)
        : 0;


    // ============================================
    // RETURN
    // ============================================

    return {

      success: true,

      scope:
        restricted
          ? department
          : 'Institution',


      stats: {

        students:
          filteredStudents.length,

        averageCGPA:
          averageCGPA,

        attendance:
          averageAttendance,

        placed:
          placedStudents,

        // SAME VALUE USED BY
        // AT-RISK STUDENTS PAGE
        atRisk:
          atRiskCount

      },


      permissions:
        getRolePermissions(
          user.role
        )

    };


  } catch (error) {

    console.error(
      'Dashboard Error:',
      error
    );


    return {

      success: false,

      message:
        error.message

    };

  }

}


function filterByDepartment(data, department) {
  return data.filter(function(row) {

    const rowDepartment =
      row.Department ||
      row.Dept ||
      row.Branch ||
      '';

    return String(rowDepartment)
      .trim()
      .toUpperCase() === department;
  });
}


function calculateAverageCGPA(data) {
  const values = data
    .map(function(row) {
      return Number(
        row.CGPA ||
        row.Current_CGPA ||
        row.Overall_CGPA
      );
    })
    .filter(function(value) {
      return !isNaN(value) && value > 0;
    });

  if (!values.length) return 0;

  const total = values.reduce(function(sum, value) {
    return sum + value;
  }, 0);

  return Number((total / values.length).toFixed(2));
}


function calculateAverageAttendance(data) {

  const values = data
    .map(function(row) {

      return Number(
        row['Overall_Attendance_%'] ||
        row.Overall_Attendance ||
        row.Attendance_Percentage ||
        row.Attendance ||
        row.Percentage ||
        0
      );

    })
    .filter(function(value) {

      return !isNaN(value) &&
             value >= 0 &&
             value <= 100;

    });


  if (!values.length) {
    return 0;
  }


  const total =
    values.reduce(function(sum, value) {
      return sum + value;
    }, 0);


  return Number(
    (total / values.length).toFixed(1)
  );
}


function calculatePlacedStudents(data) {

  return data.filter(function(row) {

    const offerStatus = String(
      row.Offer_Status ||
      row.OfferStatus ||
      ''
    )
      .trim()
      .toLowerCase();


    const interviewStatus = String(
      row.Interview_Status ||
      row.InterviewStatus ||
      ''
    )
      .trim()
      .toLowerCase();


    const selectedCompany = String(
      row.Selected_Company ||
      row.SelectedCompany ||
      ''
    ).trim();


    const packageLPA = Number(
      row.Package_LPA ||
      row.PackageLPA ||
      0
    );


    return (
      offerStatus === 'yes' ||
      interviewStatus === 'selected' ||
      (
        selectedCompany !== '' &&
        packageLPA > 0
      )
    );

  }).length;
}

/* =====================================================
   STUDENTS MODULE
===================================================== */

function getStudentsData(user) {

  try {

    if (!user || !user.role) {
      throw new Error('Invalid user session.');
    }

    let students =
      getSheetData(CONFIG.SHEETS.STUDENTS);


    const department =
      String(user.department || 'ALL')
        .trim()
        .toUpperCase();


    // HOD and Department Admin can only access
    // students belonging to their department.

    if (
      user.role === 'HOD' ||
      user.role === 'Department Admin'
    ) {

      students =
        filterByDepartment(
          students,
          department
        );

    }


    return {

      success: true,

      scope:
        (
          user.role === 'HOD' ||
          user.role === 'Department Admin'
        )
          ? department
          : 'Institution',

      count: students.length,

      students: students

    };


  } catch (error) {

    return {
      success: false,
      message: error.message
    };

  }

}

/* =====================================================
   360° STUDENT PROFILE
===================================================== */

function getStudent360(user, usn) {

  try {

    if (!user || !user.role) {
      throw new Error('Invalid user session.');
    }

    if (!usn) {
      throw new Error('Student USN is required.');
    }


    usn =
      String(usn)
        .trim()
        .toUpperCase();


    // ==========================================
    // LOAD DATASETS
    // ==========================================

    const students =
      getSheetData(CONFIG.SHEETS.STUDENTS) || [];

    const academics =
      getSheetData(CONFIG.SHEETS.ACADEMICS) || [];

    const attendance =
      getSheetData(CONFIG.SHEETS.ATTENDANCE) || [];

    const placement =
      getSheetData(CONFIG.SHEETS.PLACEMENT) || [];

    const skills =
      getSheetData(CONFIG.SHEETS.SKILLS) || [];

    const training =
      getSheetData(CONFIG.SHEETS.TRAINING) || [];

    const subjectMarks =
      getSheetData(CONFIG.SHEETS.SUBJECT_MARKS) || [];


    // ==========================================
    // FIND STUDENT
    // ==========================================

    const student =
      students.find(function(row) {

        const rowUSN =
          getBackendField(
            row,
            [
              'USN',
              'Student_ID',
              'Student_Id'
            ]
          );


        return String(rowUSN || '')
          .trim()
          .toUpperCase() === usn;

      });


    if (!student) {

      throw new Error(
        'Student record not found.'
      );

    }


    // ==========================================
    // DEPARTMENT SECURITY
    // ==========================================

    if (
      user.role === 'HOD' ||
      user.role === 'Department Admin'
    ) {

      const studentDepartment =
        String(
          getBackendField(
            student,
            [
              'Department',
              'Dept',
              'Branch'
            ]
          ) || ''
        )
          .trim()
          .toUpperCase();


      const userDepartment =
        String(
          user.department || ''
        )
          .trim()
          .toUpperCase();


      if (
        studentDepartment !==
        userDepartment
      ) {

        throw new Error(
          'You are not authorized to view this student.'
        );

      }

    }


    // ==========================================
    // RELATED RECORDS
    // ==========================================

    const studentAcademics =
      filterRecordsByUSN(
        academics,
        usn
      ) || [];

    const studentSubjectMarks =
      filterRecordsByUSN(
        subjectMarks,
        usn
      ) || [];      


    const studentAttendance =
      filterRecordsByUSN(
        attendance,
        usn
      ) || [];


    const studentPlacement =
      filterRecordsByUSN(
        placement,
        usn
      ) || [];


    const studentSkills =
      filterRecordsByUSN(
        skills,
        usn
      ) || [];


    const studentTraining =
      filterRecordsByUSN(
        training,
        usn
      ) || [];


    // ==========================================
    // SAFE DATA FOR GOOGLE.SCRIPT.RUN
    // ==========================================

    return {

      success: true,

      student:
        makeStudent360Safe(student),

      academics:
        makeStudent360Safe(studentAcademics),

      subjectMarks:
        makeStudent360Safe(studentSubjectMarks),

      attendance:
        makeStudent360Safe(studentAttendance),

      placement:
        makeStudent360Safe(studentPlacement),

      skills:
        makeStudent360Safe(studentSkills),

      training:
        makeStudent360Safe(studentTraining)

    };


  }
  catch (error) {

    return {

      success: false,

      message:
        error && error.message
          ? error.message
          : String(error)

    };

  }

}

function makeStudent360Safe(value) {

  // Array
  if (Array.isArray(value)) {

    return value.map(function(item) {
      return makeStudent360Safe(item);
    });

  }


  // Date
  if (
    Object.prototype.toString.call(value)
      === '[object Date]'
  ) {

    if (isNaN(value.getTime())) {
      return '';
    }


    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );

  }


  // Object
  if (
    value !== null &&
    typeof value === 'object'
  ) {

    const safeObject = {};


    Object.keys(value)
      .forEach(function(key) {

        safeObject[key] =
          makeStudent360Safe(
            value[key]
          );

      });


    return safeObject;

  }


  // Null / undefined
  if (
    value === null ||
    value === undefined
  ) {

    return '';

  }


  return value;

}


/* =====================================================
   FILTER RECORDS USING USN
===================================================== */

function filterRecordsByUSN(data, usn) {

  return data.filter(function(row) {

    const rowUSN =
      getBackendField(row, [
        'USN',
        'Student_USN',
        'Student_ID',
        'Student_Id'
      ]);


    return String(rowUSN)
      .trim()
      .toUpperCase() === usn;

  });

}


/* =====================================================
   BACKEND FIELD HELPER
===================================================== */

function getBackendField(
  object,
  fields
) {

  for (
    let i = 0;
    i < fields.length;
    i++
  ) {

    if (
      object[fields[i]] !== undefined &&
      object[fields[i]] !== null &&
      object[fields[i]] !== ''
    ) {

      return object[fields[i]];

    }

  }

  return '';

}

/* =====================================================
   ACADEMIC INTELLIGENCE MODULE
===================================================== */

function getAcademicsData(user) {

  try {

    if (!user || !user.role) {
      throw new Error('Invalid user session.');
    }

    let students =
      getSheetData(CONFIG.SHEETS.STUDENTS);

    let academics =
      getSheetData(CONFIG.SHEETS.ACADEMICS);


    const department =
      String(user.department || 'ALL')
        .trim()
        .toUpperCase();


    /* =============================================
       ROLE BASED DEPARTMENT FILTER
    ============================================= */

    if (
      user.role === 'HOD' ||
      user.role === 'Department Admin'
    ) {

      students =
        filterByDepartment(
          students,
          department
        );

    }


    /* =============================================
       GET ALLOWED STUDENT USNs
    ============================================= */

    const allowedUSNs =
      students.map(function(student) {

        return String(
          getBackendField(
            student,
            [
              'USN',
              'Student_ID',
              'Student_Id'
            ]
          )
        )
          .trim()
          .toUpperCase();

      });


    academics =
      academics.filter(function(record) {

        const usn =
          String(
            getBackendField(
              record,
              [
                'USN',
                'Student_USN',
                'Student_ID'
              ]
            )
          )
            .trim()
            .toUpperCase();

        return allowedUSNs.includes(usn);

      });


    /* =============================================
       COMBINE STUDENT + ACADEMIC DATA
    ============================================= */

    const records =
      academics.map(function(record) {

        const usn =
          String(
            getBackendField(
              record,
              [
                'USN',
                'Student_USN',
                'Student_ID'
              ]
            )
          )
            .trim()
            .toUpperCase();


        const student =
          students.find(function(s) {

            return String(
              getBackendField(
                s,
                [
                  'USN',
                  'Student_ID',
                  'Student_Id'
                ]
              )
            )
              .trim()
              .toUpperCase() === usn;

          }) || {};


        return {

          usn: usn,

          name:
            getBackendField(
              student,
              [
                'Student_Name',
                'Name',
                'Full_Name'
              ]
            ),

          department:
            getBackendField(
              student,
              [
                'Department',
                'Dept',
                'Branch'
              ]
            ),

          semester:
            getBackendField(
              record,
              [
                'Semester',
                'Sem'
              ]
            ),

          sgpa:
            getBackendField(
              record,
              [
                'SGPA',
                'Semester_GPA'
              ]
            ),

          branchRank:
            getBackendField(
              record,
              [
                'Branch_Rank',
                'Rank'
              ]
            ),

          attempts:
            getBackendField(
              record,
              [
                'Attempts',
                'Attempt'
              ]
            ),

          puc:
            getBackendField(
              record,
              [
                'PUC/Diploma_%',
                'PUC_Diploma_%',
                'PUC_Percentage'
              ]
            )

        };

      });


    /* =============================================
       ANALYTICS
    ============================================= */

    const sgpaValues =
      records
        .map(function(record) {
          return Number(record.sgpa);
        })
        .filter(function(value) {
          return !isNaN(value) && value > 0;
        });


    const averageSGPA =
      sgpaValues.length
        ? sgpaValues.reduce(
            function(total, value) {
              return total + value;
            },
            0
          ) / sgpaValues.length
        : 0;


    const topSGPA =
      sgpaValues.length
        ? Math.max.apply(
            null,
            sgpaValues
          )
        : 0;


    /* SGPA below 6.5 = Needs Attention */

    const attention =
      records.filter(function(record) {

        const sgpa =
          Number(record.sgpa);

        return (
          !isNaN(sgpa) &&
          sgpa > 0 &&
          sgpa < 6.5
        );

      }).length;


    return {

      success: true,

      scope:
        (
          user.role === 'HOD' ||
          user.role === 'Department Admin'
        )
          ? department
          : 'Institution',

      students:
        new Set(allowedUSNs).size,

      records: records,

      averageSGPA:
        averageSGPA,

      topSGPA:
        topSGPA,

      attention:
        attention

    };


  } catch (error) {

    return {
      success: false,
      message: error.message
    };

  }

}

/* =====================================================
   ATTENDANCE INTELLIGENCE MODULE
===================================================== */

function getAttendanceData(user) {

  try {

    if (!user || !user.role) {
      throw new Error('Invalid user session.');
    }


    let students =
      getSheetData(CONFIG.SHEETS.STUDENTS);

    let attendance =
      getSheetData(CONFIG.SHEETS.ATTENDANCE);


    const department =
      String(user.department || 'ALL')
        .trim()
        .toUpperCase();


    /* =============================================
       ROLE BASED FILTERING
    ============================================= */

    if (
      user.role === 'HOD' ||
      user.role === 'Department Admin'
    ) {

      students =
        filterByDepartment(
          students,
          department
        );

    }


    /* =============================================
       ALLOWED STUDENT USNs
    ============================================= */

    const allowedUSNs =
      students.map(function(student) {

        return String(
          getBackendField(
            student,
            [
              'USN',
              'Student_ID',
              'Student_Id'
            ]
          )
        )
          .trim()
          .toUpperCase();

      });


    /* =============================================
       FILTER ATTENDANCE
    ============================================= */

    attendance =
      attendance.filter(function(record) {

        const usn =
          String(
            getBackendField(
              record,
              [
                'USN',
                'Student_USN',
                'Student_ID'
              ]
            )
          )
            .trim()
            .toUpperCase();


        return allowedUSNs.includes(usn);

      });


    /* =============================================
       COMBINE STUDENT + ATTENDANCE
    ============================================= */

    const records =
      attendance.map(function(record) {

        const usn =
          String(
            getBackendField(
              record,
              [
                'USN',
                'Student_USN',
                'Student_ID'
              ]
            )
          )
            .trim()
            .toUpperCase();


        const student =
          students.find(function(s) {

            const studentUSN =
              String(
                getBackendField(
                  s,
                  [
                    'USN',
                    'Student_ID',
                    'Student_Id'
                  ]
                )
              )
                .trim()
                .toUpperCase();


            return studentUSN === usn;

          }) || {};


        return {

          usn: usn,

          name:
            getBackendField(
              student,
              [
                'Student_Name',
                'Name',
                'Full_Name'
              ]
            ),

          department:
            getBackendField(
              student,
              [
                'Department',
                'Dept',
                'Branch'
              ]
            ),

          semester:
            getBackendField(
              record,
              ['Semester']
            ),

          month:
            getBackendField(
              record,
              ['Month']
            ),

          percentage:
            getBackendField(
              record,
              [
                'Overall_Attendance_%'
              ]
            ),

          math:
            getBackendField(
              record,
              ['Math_%']
            ),

          physics:
            getBackendField(
              record,
              ['Physics_%']
            ),

          programming:
            getBackendField(
              record,
              ['Programming_%']
            ),

          electronics:
            getBackendField(
              record,
              ['Electronics_%']
            ),

          english:
            getBackendField(
              record,
              ['English_%']
            ),

          medicalLeave:
            getBackendField(
              record,
              ['Medical_Leave']
            ),

          attendanceStatus:
            getBackendField(
              record,
              ['Status']
            )

        };

      });


    /* =============================================
       ANALYTICS
    ============================================= */

    const percentages =
      records
        .map(function(record) {

          return Number(
            record.percentage
          );

        })
        .filter(function(value) {

          return (
            !isNaN(value) &&
            value >= 0
          );

        });


    const averageAttendance =
      percentages.length
        ? percentages.reduce(
            function(total, value) {
              return total + value;
            },
            0
          ) / percentages.length
        : 0;


    const highestAttendance =
      percentages.length
        ? Math.max.apply(
            null,
            percentages
          )
        : 0;


    /*
       Count unique students whose latest/minimum
       available attendance is below 75%.
    */

    const lowAttendanceUSNs = {};


    records.forEach(function(record) {

      const percentage =
        Number(record.percentage);


      if (
        !isNaN(percentage) &&
        percentage < 75
      ) {

        lowAttendanceUSNs[
          record.usn
        ] = true;

      }

    });


    return {

      success: true,

      scope:
        (
          user.role === 'HOD' ||
          user.role === 'Department Admin'
        )
          ? department
          : 'Institution',

      students:
        new Set(
          allowedUSNs
        ).size,

      records: records,

      averageAttendance:
        averageAttendance,

      highestAttendance:
        highestAttendance,

      below75:
        Object.keys(
          lowAttendanceUSNs
        ).length

    };


  } catch (error) {

    return {

      success: false,
      message: error.message

    };

  }

}

/* =====================================================
   SIDSS - PLACEMENT INTELLIGENCE
===================================================== */

function getPlacementData(user) {

  try {

    if (!user || !user.role) {
      throw new Error('Invalid user session.');
    }


    let students =
      getSheetData(CONFIG.SHEETS.STUDENTS);

    let placements =
      getSheetData(CONFIG.SHEETS.PLACEMENT);

    let academics =
      getSheetData(CONFIG.SHEETS.ACADEMICS);


    const department =
      String(user.department || 'ALL')
        .trim()
        .toUpperCase();


    /* =================================================
       ROLE BASED DEPARTMENT FILTER
    ================================================= */

    if (
      user.role === 'HOD' ||
      user.role === 'Department Admin'
    ) {

      students =
        filterByDepartment(
          students,
          department
        );

    }


    /* =================================================
       ALLOWED STUDENTS
    ================================================= */

    const allowedUSNs =
      students.map(function(student) {

        return String(
          getBackendField(
            student,
            [
              'USN',
              'Student_ID',
              'Student_Id'
            ]
          )
        )
          .trim()
          .toUpperCase();

      });


    /* =================================================
       FILTER PLACEMENT DATA
    ================================================= */

    placements =
      placements.filter(function(record) {

        const usn =
          String(
            getBackendField(
              record,
              ['USN']
            )
          )
            .trim()
            .toUpperCase();


        return allowedUSNs.includes(usn);

      });


    /* =================================================
       BUILD PLACEMENT RECORDS
    ================================================= */

    const records =
      placements.map(function(record) {

        const usn =
          String(
            getBackendField(
              record,
              ['USN']
            )
          )
            .trim()
            .toUpperCase();


        const student =
          students.find(function(s) {

            return String(
              getBackendField(
                s,
                [
                  'USN',
                  'Student_ID',
                  'Student_Id'
                ]
              )
            )
              .trim()
              .toUpperCase() === usn;
          

          }) || {};


        return {

          usn: usn,


          name:
            getBackendField(
              student,
              [
                'Student_Name',
                'Name',
                'Full_Name'
              ]
            ),


          department:
            getBackendField(
              student,
              [
                'Department',
                'Dept',
                'Branch'
              ]
            ),


          eligibility:
            getBackendField(
              record,
              ['Eligibility']
            ),


          resume:
            getBackendField(
              record,
              ['Resume_Link']
            ),


          aptitude:
            getBackendField(
              record,
              ['Aptitude_Score']
            ),


          coding:
            getBackendField(
              record,
              ['Coding_Score']
            ),


          communication:
            getBackendField(
              record,
              ['Communication_Score']
            ),


          interviewStatus:
            getBackendField(
              record,
              ['Interview_Status']
            ),


          companyApplied:
            getBackendField(
              record,
              ['Companies_Applied']
            ),


          selectedCompany:
            getBackendField(
              record,
              ['Selected_Company']
            ),


          offerStatus:
            getBackendField(
              record,
              ['Offer_Status']
            ),


          packageLPA:
            getBackendField(
              record,
              ['Package_LPA']
            ),


          readiness:
            getBackendField(
              record,
              ['Placement_Readiness_Score']
            )

        };

      });


    /* =================================================
       ANALYTICS
    ================================================= */

    const eligibleStudents =
      records.filter(function(record) {

        return String(
          record.eligibility
        )
          .trim()
          .toLowerCase() === 'yes';

      }).length;


    const placedStudents =
      records.filter(function(record) {

        return String(
          record.offerStatus
        )
          .trim()
          .toLowerCase() === 'yes';

      }).length;


    const readinessValues =
      records
        .map(function(record) {

          return Number(
            record.readiness
          );

        })
        .filter(function(value) {

          return !isNaN(value);

        });


    const averageReadiness =
      readinessValues.length
        ? readinessValues.reduce(
            function(total, value) {

              return total + value;

            },
            0
          ) / readinessValues.length
        : 0;


    const packages =
      records
        .map(function(record) {

          return Number(
            record.packageLPA
          );

        })
        .filter(function(value) {

          return (
            !isNaN(value) &&
            value > 0
          );

        });


    const highestPackage =
      packages.length
        ? Math.max.apply(
            null,
            packages
          )
        : 0;


    const averageCoding =
      calculatePlacementAverage(
        records,
        'coding'
      );


    const averageAptitude =
      calculatePlacementAverage(
        records,
        'aptitude'
      );


    const averageCommunication =
      calculatePlacementAverage(
        records,
        'communication'
      );


    return {

      success: true,


      scope:
        (
          user.role === 'HOD' ||
          user.role === 'Department Admin'
        )
          ? department
          : 'Institution',


      students:
        new Set(
          allowedUSNs
        ).size,


      records: records,


      eligibleStudents:
        eligibleStudents,


      placedStudents:
        placedStudents,


      averageReadiness:
        averageReadiness,


      highestPackage:
        highestPackage,


      averageCoding:
        averageCoding,


      averageAptitude:
        averageAptitude,


      averageCommunication:
        averageCommunication

    };


  } catch (error) {

    return {

      success: false,
      message: error.message

    };

  }

}


/* =====================================================
   PLACEMENT AVERAGE HELPER
===================================================== */

function calculatePlacementAverage(
  records,
  field
) {

  const values =
    records
      .map(function(record) {

        return Number(
          record[field]
        );

      })
      .filter(function(value) {

        return !isNaN(value);

      });


  if (!values.length) {
    return 0;
  }


  return values.reduce(
    function(total, value) {

      return total + value;

    },
    0
  ) / values.length;

}

/* =====================================================
   SIDSS - SKILLS INTELLIGENCE
===================================================== */

function getSkillsData(user) {

  try {

    if (!user || !user.role) {
      throw new Error('Invalid user session.');
    }

    let students =
      getSheetData(CONFIG.SHEETS.STUDENTS);

    let skills =
      getSheetData(CONFIG.SHEETS.SKILLS);


    const department =
      String(user.department || 'ALL')
        .trim()
        .toUpperCase();


    /* =================================================
       ROLE BASED DEPARTMENT FILTER
    ================================================= */

    if (
      user.role === 'HOD' ||
      user.role === 'Department Admin'
    ) {

      students =
        filterByDepartment(
          students,
          department
        );

    }


    /* =================================================
       ALLOWED STUDENTS
    ================================================= */

    const allowedUSNs =
      students.map(function(student) {

        return String(
          getBackendField(
            student,
            [
              'USN',
              'Student_ID',
              'Student_Id'
            ]
          )
        )
          .trim()
          .toUpperCase();

      });


    /* =================================================
       FILTER SKILLS
    ================================================= */

    skills =
      skills.filter(function(record) {

        const usn =
          String(
            getBackendField(
              record,
              ['USN']
            )
          )
            .trim()
            .toUpperCase();

        return allowedUSNs.includes(usn);

      });


    /* =================================================
       BUILD RECORDS
    ================================================= */

    const records =
      skills.map(function(record) {

        const usn =
          String(
            getBackendField(
              record,
              ['USN']
            )
          )
            .trim()
            .toUpperCase();


        const student =
          students.find(function(s) {

            return String(
              getBackendField(
                s,
                [
                  'USN',
                  'Student_ID',
                  'Student_Id'
                ]
              )
            )
              .trim()
              .toUpperCase() === usn;

          }) || {};


        return {

          usn: usn,

          name:
            getBackendField(
              student,
              [
                'Student_Name',
                'Name',
                'Full_Name'
              ]
            ),

          department:
            getBackendField(
              student,
              [
                'Department',
                'Dept',
                'Branch'
              ]
            ),

          skill:
            getBackendField(
              record,
              ['Skill_Name']
            ),

          level:
            getBackendField(
              record,
              ['Skill_Level']
            ),

          verified:
            getBackendField(
              record,
              ['Verified']
            ),

          experience:
            getBackendField(
              record,
              ['Years_of_Experience']
            ),

          lastUpdated:
            formatSkillDate(
              getBackendField(
                record,
                ['Last_Updated']
              )
            )

        };

      });


    /* =================================================
       ANALYTICS
    ================================================= */

    const uniqueStudents =
      new Set(
        records.map(function(record) {
          return record.usn;
        })
      ).size;


    const advancedSkills =
      records.filter(function(record) {

        return String(record.level)
          .trim()
          .toLowerCase() === 'advanced';

      }).length;


    const intermediateSkills =
      records.filter(function(record) {

        return String(record.level)
          .trim()
          .toLowerCase() === 'intermediate';

      }).length;


    const beginnerSkills =
      records.filter(function(record) {

        return String(record.level)
          .trim()
          .toLowerCase() === 'beginner';

      }).length;


    const verifiedSkills =
      records.filter(function(record) {

        return String(record.verified)
          .trim()
          .toLowerCase() === 'yes';

      }).length;


    /* =================================================
       SKILL POPULARITY
    ================================================= */

    const skillCounts = {};

    records.forEach(function(record) {

      const skill =
        String(record.skill || '').trim();

      if (!skill) return;

      if (!skillCounts[skill]) {
        skillCounts[skill] = 0;
      }

      skillCounts[skill]++;

    });


    const topSkills =
      Object.keys(skillCounts)
        .map(function(skill) {

          return {
            skill: skill,
            count: skillCounts[skill]
          };

        })
        .sort(function(a, b) {
          return b.count - a.count;
        });


    return {

      success: true,

      scope:
        (
          user.role === 'HOD' ||
          user.role === 'Department Admin'
        )
          ? department
          : 'Institution',

      totalStudents:
        uniqueStudents,

      totalSkills:
        records.length,

      advancedSkills:
        advancedSkills,

      intermediateSkills:
        intermediateSkills,

      beginnerSkills:
        beginnerSkills,

      verifiedSkills:
        verifiedSkills,

      topSkills:
        topSkills,

      records:
        records

    };


  } catch (error) {

    return {
      success: false,
      message: error.message
    };

  }

}

function formatSkillDate(value) {

  if (!value) {
    return '';
  }

  // If Google Sheets returned a Date object
  if (
    Object.prototype.toString.call(value) === '[object Date]' &&
    !isNaN(value.getTime())
  ) {

    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );

  }

  // Already text
  return String(value);

}

/* =====================================================
   SIDSS - TRAINING & CERTIFICATIONS INTELLIGENCE
===================================================== */

function getTrainingData(user) {

  try {

    if (!user || !user.role) {
      throw new Error('Invalid user session.');
    }

    let students =
      getSheetData(CONFIG.SHEETS.STUDENTS);

    let training =
      getSheetData(CONFIG.SHEETS.TRAINING);


    const department =
      String(user.department || 'ALL')
        .trim()
        .toUpperCase();


    /* ROLE BASED FILTER */

    if (
      user.role === 'HOD' ||
      user.role === 'Department Admin'
    ) {

      students =
        filterByDepartment(
          students,
          department
        );

    }


    /* ALLOWED STUDENTS */

    const allowedUSNs =
      students.map(function(student) {

        return String(
          getBackendField(
            student,
            ['USN', 'Student_ID', 'Student_Id']
          )
        )
        .trim()
        .toUpperCase();

      });


    /* FILTER TRAINING RECORDS */

    training =
      training.filter(function(record) {

        const usn =
          String(
            getBackendField(
              record,
              ['USN']
            )
          )
          .trim()
          .toUpperCase();

        return allowedUSNs.includes(usn);

      });


    /* BUILD RECORDS */

    const records =
      training.map(function(record) {

        const usn =
          String(
            getBackendField(
              record,
              ['USN']
            )
          )
          .trim()
          .toUpperCase();


        const student =
          students.find(function(s) {

            return String(
              getBackendField(
                s,
                ['USN', 'Student_ID', 'Student_Id']
              )
            )
            .trim()
            .toUpperCase() === usn;

          }) || {};


        return {

          usn: usn,

          name:
            getBackendField(
              student,
              ['Student_Name', 'Name', 'Full_Name']
            ),

          department:
            getBackendField(
              student,
              ['Department', 'Dept', 'Branch']
            ),

          category:
            getBackendField(
              record,
              ['Category']
            ),

          certificateName:
            getBackendField(
              record,
              ['Certificate_Name']
            ),

          provider:
            getBackendField(
              record,
              ['Provider']
            ),

          completionDate:
            formatBackendDate(
              getBackendField(
                record,
                ['Completion_Date']
              )
            ),

          certificateId:
            getBackendField(
              record,
              ['Certificate_ID']
            ),

          certificateLink:
            getBackendField(
              record,
              ['Certificate_Link']
            ),

          status:
            getBackendField(
              record,
              ['Status']
            )

        };

      });


    /* ANALYTICS */

    const uniqueStudents =
      new Set(
        records.map(function(r) {
          return r.usn;
        })
      ).size;


    const completed =
      records.filter(function(r) {

        return String(r.status)
          .trim()
          .toLowerCase() === 'completed';

      }).length;


    const inProgress =
      records.filter(function(r) {

        return String(r.status)
          .trim()
          .toLowerCase() === 'in progress';

      }).length;


    /* CATEGORY COUNTS */

    const categoryCounts = {};

    records.forEach(function(r) {

      const category =
        String(r.category || 'Other').trim();

      if (!categoryCounts[category]) {
        categoryCounts[category] = 0;
      }

      categoryCounts[category]++;

    });


    const categories =
      Object.keys(categoryCounts)
        .map(function(category) {

          return {
            category: category,
            count: categoryCounts[category]
          };

        })
        .sort(function(a, b) {
          return b.count - a.count;
        });


    /* PROVIDER COUNTS */

    const providerCounts = {};

    records.forEach(function(r) {

      const provider =
        String(r.provider || 'Other').trim();

      if (!providerCounts[provider]) {
        providerCounts[provider] = 0;
      }

      providerCounts[provider]++;

    });


    const providers =
      Object.keys(providerCounts)
        .map(function(provider) {

          return {
            provider: provider,
            count: providerCounts[provider]
          };

        })
        .sort(function(a, b) {
          return b.count - a.count;
        })
        .slice(0, 6);


    return {

      success: true,

      scope:
        (
          user.role === 'HOD' ||
          user.role === 'Department Admin'
        )
          ? department
          : 'Institution',

      totalStudents: uniqueStudents,

      totalRecords: records.length,

      completed: completed,

      inProgress: inProgress,

      completionRate:
        records.length
          ? (completed / records.length) * 100
          : 0,

      categories: categories,

      providers: providers,

      records: records

    };


  } catch (error) {

    return {
      success: false,
      message: error.message
    };

  }

}

function formatBackendDate(value) {

  if (!value) {
    return '';
  }

  if (
    Object.prototype.toString.call(value) ===
    '[object Date]'
  ) {

    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );

  }

  return String(value);
}


/* =====================================================
   SIDSS - EVENTS INTELLIGENCE
===================================================== */

function getEventsData(user) {

  try {

    if (!user || !user.role) {
      throw new Error('Invalid user session.');
    }

    let events =
      getSheetData(CONFIG.SHEETS.EVENTS);


    const department =
      String(user.department || 'ALL')
        .trim()
        .toUpperCase();


    /* =================================================
       ROLE BASED EVENT VISIBILITY

       HOD / Department Admin:
       - Own department events
       - ALL department events

       Principal / Institution Admin / Placement:
       - All events
    ================================================= */

    if (
      user.role === 'HOD' ||
      user.role === 'Department Admin'
    ) {

      events =
        events.filter(function(record) {

          const eventDepartment =
            String(
              getBackendField(
                record,
                ['Department']
              ) || ''
            )
              .trim()
              .toUpperCase();


          return (
            eventDepartment === department ||
            eventDepartment === 'ALL'
          );

        });

    }


    /* =================================================
       BUILD RECORDS
    ================================================= */

    const records =
      events.map(function(record) {

        return {

          eventId:
            getBackendField(
              record,
              ['Event_ID']
            ),

          eventName:
            getBackendField(
              record,
              ['Event_Name']
            ),

          eventType:
            getBackendField(
              record,
              ['Event_Type']
            ),

          department:
            getBackendField(
              record,
              ['Department']
            ),

          eventDate:
            getBackendField(
              record,
              ['Event_Date']
            ),

          venue:
            getBackendField(
              record,
              ['Venue']
            ),

          coordinator:
            getBackendField(
              record,
              ['Coordinator']
            ),

          eligibleStudents:
            getBackendField(
              record,
              ['Eligible_Students']
            ),

          participants:
            getBackendField(
              record,
              ['Participants']
            ),

          status:
            getBackendField(
              record,
              ['Status']
            ),

          remarks:
            getBackendField(
              record,
              ['Remarks']
            )

        };

      });


    /* =================================================
       ANALYTICS
    ================================================= */

    const totalEvents =
      records.length;


    const upcomingEvents =
      records.filter(function(record) {

        return String(record.status || '')
          .trim()
          .toLowerCase() === 'upcoming';

      }).length;


    const completedEvents =
      records.filter(function(record) {

        return String(record.status || '')
          .trim()
          .toLowerCase() === 'completed';

      }).length;


    const ongoingEvents =
      records.filter(function(record) {

        const status =
          String(record.status || '')
            .trim()
            .toLowerCase();

        return (
          status === 'ongoing' ||
          status === 'in progress'
        );

      }).length;


    const totalParticipants =
      records.reduce(function(total, record) {

        const value =
          Number(record.participants || 0);

        return total +
          (isNaN(value) ? 0 : value);

      }, 0);


    /* =================================================
       EVENT TYPE DISTRIBUTION
    ================================================= */

    const typeCounts = {};


    records.forEach(function(record) {

      const type =
        String(
          record.eventType || 'Other'
        ).trim();


      if (!typeCounts[type]) {
        typeCounts[type] = 0;
      }


      typeCounts[type]++;

    });


    const eventTypes =
      Object.keys(typeCounts)
        .map(function(type) {

          return {
            type: type,
            count: typeCounts[type]
          };

        })
        .sort(function(a, b) {
          return b.count - a.count;
        });


    return {

      success: true,

      scope:
        (
          user.role === 'HOD' ||
          user.role === 'Department Admin'
        )
          ? department
          : 'Institution',

      totalEvents:
        totalEvents,

      upcomingEvents:
        upcomingEvents,

      completedEvents:
        completedEvents,

      ongoingEvents:
        ongoingEvents,

      totalParticipants:
        totalParticipants,

      eventTypes:
        eventTypes,

      records:
        records

    };


  } catch (error) {

    return {

      success: false,
      message: error.message

    };

  }

}


function addStudent(student, currentUser) {

  try {

    // Only Department Admin can add students
    if (
      !currentUser ||
      currentUser.role !== 'Department Admin'
    ) {
      return {
        success: false,
        message: 'You do not have permission to add students.'
      };
    }


    const ss =
      SpreadsheetApp.openById(
        '1Sn6HAHna6xpx4D3ntLZiSu5JRvcR8U-DfLnSI3BU1Og'
      );

    const sheet =
      ss.getSheetByName('Student_Master');


    if (!sheet) {
      return {
        success: false,
        message: 'Student_Master sheet not found.'
      };
    }


    const usn =
      String(student.usn || '')
        .trim()
        .toUpperCase();

    const name =
      String(student.name || '').trim();

    const email =
      String(student.email || '')
        .trim()
        .toLowerCase();

    const semester =
      String(student.semester || '').trim();

    const status =
      String(student.status || 'Active').trim();


    if (!usn || !name || !email || !semester) {

      return {
        success: false,
        message: 'Please fill all required fields.'
      };

    }


    // Check duplicate USN
    const lastRow = sheet.getLastRow();

    if (lastRow > 1) {

      const existingUSNs =
        sheet
          .getRange(2, 1, lastRow - 1, 1)
          .getValues()
          .flat()
          .map(function(value) {
            return String(value)
              .trim()
              .toUpperCase();
          });


      if (existingUSNs.includes(usn)) {

        return {
          success: false,
          message: 'Student with this USN already exists.'
        };

      }

    }


    // Department is automatically taken
    // from logged-in Department Admin
    const department =
      currentUser.department;


    // Student_Master column order:
    // USN
    // Admission_No
    // Student_Name
    // Gender
    // DOB
    // Phone
    // Email
    // Department
    // Program
    // Semester
    // Section
    // Batch
    // Admission_Year
    // Photo_URL
    // Status

    sheet.appendRow([

      usn,               // USN
      '',                // Admission_No
      name,              // Student_Name
      '',                // Gender
      '',                // DOB
      '',                // Phone
      email,             // Email
      department,        // Department
      'B.E.',            // Program
      semester,          // Semester
      '',                // Section
      '',                // Batch
      new Date().getFullYear(), // Admission_Year
      '',                // Photo_URL
      status             // Status

    ]);


    return {
      success: true,
      message: 'Student added successfully.'
    };


  }
  catch (error) {

    return {
      success: false,
      message: error.message
    };

  }

}

function addAcademicRecord(record, currentUser) {

  try {

    // Only Department Admin can add academic records
    if (
      !currentUser ||
      currentUser.role !== 'Department Admin'
    ) {
      return {
        success: false,
        message: 'You do not have permission to add academic records.'
      };
    }


    // Open SIDSS database
    const ss =
      SpreadsheetApp.openById(
        '1Sn6HAHna6xpx4D3ntLZiSu5JRvcR8U-DfLnSI3BU1Og'
      );


    const academicSheet =
      ss.getSheetByName('Academics');

    const studentSheet =
      ss.getSheetByName('Student_Master');


    if (!academicSheet) {
      return {
        success: false,
        message: 'Academics sheet not found.'
      };
    }


    if (!studentSheet) {
      return {
        success: false,
        message: 'Student_Master sheet not found.'
      };
    }


    // Clean input
    const usn =
      String(record.usn || '')
        .trim()
        .toUpperCase();

    const semester =
      String(record.semester || '')
        .trim();

    const sgpa =
      Number(record.sgpa);

    const cgpa =
      Number(record.cgpa);

    const backlogs =
      Number(record.backlogs || 0);

    const attempts =
      Number(record.attempts || 1);

    const classRank =
      record.classRank === ''
        ? ''
        : Number(record.classRank);

    const branchRank =
      record.branchRank === ''
        ? ''
        : Number(record.branchRank);

    const sslc =
      record.sslc === ''
        ? ''
        : Number(record.sslc);

    const puc =
      record.puc === ''
        ? ''
        : Number(record.puc);


    // Required fields
    if (
      !usn ||
      !semester ||
      record.sgpa === '' ||
      record.cgpa === ''
    ) {
      return {
        success: false,
        message: 'Please fill all required fields.'
      };
    }


    // SGPA / CGPA validation
    if (
      isNaN(sgpa) ||
      sgpa < 0 ||
      sgpa > 10
    ) {
      return {
        success: false,
        message: 'SGPA must be between 0 and 10.'
      };
    }


    if (
      isNaN(cgpa) ||
      cgpa < 0 ||
      cgpa > 10
    ) {
      return {
        success: false,
        message: 'CGPA must be between 0 and 10.'
      };
    }


    // -----------------------------------------
    // VERIFY STUDENT + DEPARTMENT
    // Student_Master:
    // Column A = USN
    // Column H = Department
    // -----------------------------------------

    const studentLastRow =
      studentSheet.getLastRow();


    if (studentLastRow <= 1) {

      return {
        success: false,
        message: 'No students found in Student_Master.'
      };

    }


    const studentData =
      studentSheet
        .getRange(
          2,
          1,
          studentLastRow - 1,
          15
        )
        .getValues();


    let studentFound = false;
    let studentDepartment = '';


    for (
      let i = 0;
      i < studentData.length;
      i++
    ) {

      const studentUSN =
        String(studentData[i][0] || '')
          .trim()
          .toUpperCase();


      if (studentUSN === usn) {

        studentFound = true;

        // Column H = index 7
        studentDepartment =
          String(studentData[i][7] || '')
            .trim();

        break;

      }

    }


    if (!studentFound) {

      return {
        success: false,
        message: 'Student USN not found in Student_Master.'
      };

    }


    // Department security
    if (
      studentDepartment.toUpperCase() !==
      String(currentUser.department || '')
        .trim()
        .toUpperCase()
    ) {

      return {
        success: false,
        message:
          'You can only add academic records for students in your department.'
      };

    }


    // -----------------------------------------
    // CHECK DUPLICATE USN + SEMESTER
    // -----------------------------------------

    const academicLastRow =
      academicSheet.getLastRow();


    if (academicLastRow > 1) {

      const existingRecords =
        academicSheet
          .getRange(
            2,
            1,
            academicLastRow - 1,
            2
          )
          .getValues();


      for (
        let i = 0;
        i < existingRecords.length;
        i++
      ) {

        const existingUSN =
          String(existingRecords[i][0] || '')
            .trim()
            .toUpperCase();

        const existingSemester =
          String(existingRecords[i][1] || '')
            .trim();


        if (
          existingUSN === usn &&
          existingSemester === semester
        ) {

          return {
            success: false,
            message:
              'Academic record already exists for this student and semester.'
          };

        }

      }

    }


    // -----------------------------------------
    // SAVE TO ACADEMICS
    // -----------------------------------------
    //
    // A  USN
    // B  Semester
    // C  SGPA
    // D  CGPA
    // E  Backlogs
    // F  Attempts
    // G  Class_Rank
    // H  Branch_Rank
    // I  SSLC_%
    // J  PUC/Diploma_%
    // -----------------------------------------

    academicSheet.appendRow([

      usn,
      Number(semester),
      sgpa,
      cgpa,
      backlogs,
      attempts,
      classRank,
      branchRank,
      sslc,
      puc

    ]);


    return {
      success: true,
      message: 'Academic record added successfully.'
    };


  }
  catch (error) {

    return {
      success: false,
      message: error.message
    };

  }

}


function updateStudent(student, currentUser) {

  try {

    // Only Department Admin can edit students
    if (
      !currentUser ||
      currentUser.role !== 'Department Admin'
    ) {

      return {
        success: false,
        message: 'You do not have permission to edit students.'
      };

    }


    const ss =
      SpreadsheetApp.openById(
        '1Sn6HAHna6xpx4D3ntLZiSu5JRvcR8U-DfLnSI3BU1Og'
      );


    const sheet =
      ss.getSheetByName('Student_Master');


    if (!sheet) {

      return {
        success: false,
        message: 'Student_Master sheet not found.'
      };

    }


    const usn =
      String(student.usn || '')
        .trim()
        .toUpperCase();


    const name =
      String(student.name || '')
        .trim();


    const email =
      String(student.email || '')
        .trim()
        .toLowerCase();


    const phone =
      String(student.phone || '')
        .trim();


    const semester =
      String(student.semester || '')
        .trim();


    const section =
      String(student.section || '')
        .trim()
        .toUpperCase();


    const status =
      String(student.status || 'Active')
        .trim();


    if (
      !usn ||
      !name ||
      !email ||
      !semester
    ) {

      return {
        success: false,
        message: 'Please fill all required fields.'
      };

    }


    const lastRow =
      sheet.getLastRow();


    if (lastRow < 2) {

      return {
        success: false,
        message: 'Student record not found.'
      };

    }


    // USN is column A
    const usnValues =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          1
        )
        .getValues();


    let studentRow = -1;


    for (
      let i = 0;
      i < usnValues.length;
      i++
    ) {

      const existingUSN =
        String(usnValues[i][0])
          .trim()
          .toUpperCase();


      if (existingUSN === usn) {

        studentRow =
          i + 2;

        break;

      }

    }


    if (studentRow === -1) {

      return {
        success: false,
        message: 'Student record not found.'
      };

    }


    // Security:
    // Department Admin can edit only their department.
    // Department is column H = 8.

    const studentDepartment =
      String(
        sheet
          .getRange(studentRow, 8)
          .getValue()
      )
        .trim()
        .toUpperCase();


    const adminDepartment =
      String(currentUser.department || '')
        .trim()
        .toUpperCase();


    if (
      studentDepartment !==
      adminDepartment
    ) {

      return {
        success: false,
        message:
          'You cannot edit students from another department.'
      };

    }


    /*
      Student_Master columns:

      A  USN
      B  Admission_No
      C  Student_Name
      D  Gender
      E  DOB
      F  Phone
      G  Email
      H  Department
      I  Program
      J  Semester
      K  Section
      L  Batch
      M  Admission_Year
      N  Photo_URL
      O  Status
    */


    // C - Student Name
    sheet
      .getRange(studentRow, 3)
      .setValue(name);


    // F - Phone
    sheet
      .getRange(studentRow, 6)
      .setValue(phone);


    // G - Email
    sheet
      .getRange(studentRow, 7)
      .setValue(email);


    // J - Semester
    sheet
      .getRange(studentRow, 10)
      .setValue(semester);


    // K - Section
    sheet
      .getRange(studentRow, 11)
      .setValue(section);


    // O - Status
    sheet
      .getRange(studentRow, 15)
      .setValue(status);


    return {

      success: true,

      message:
        'Student updated successfully.'

    };


  }
  catch (error) {

    return {

      success: false,

      message: error.message

    };

  }

}


function addSkill(skill, currentUser) {

  try {

    // Department Admin only
    if (
      !currentUser ||
      currentUser.role !== 'Department Admin'
    ) {

      return {
        success: false,
        message: 'You do not have permission to add skills.'
      };

    }


    const ss =
      SpreadsheetApp.openById(
        '1Sn6HAHna6xpx4D3ntLZiSu5JRvcR8U-DfLnSI3BU1Og'
      );


    const skillsSheet =
      ss.getSheetByName('Skills');


    const studentSheet =
      ss.getSheetByName('Student_Master');


    if (!skillsSheet) {

      return {
        success: false,
        message: 'Skills sheet not found.'
      };

    }


    if (!studentSheet) {

      return {
        success: false,
        message: 'Student_Master sheet not found.'
      };

    }


    const usn =
      String(skill.usn || '')
        .trim()
        .toUpperCase();


    const skillName =
      String(skill.skillName || '')
        .trim();


    const skillLevel =
      String(skill.skillLevel || '')
        .trim();


    const verified =
      String(skill.verified || 'No')
        .trim();


    const yearsExperience =
      Number(skill.yearsExperience || 0);


    if (
      !usn ||
      !skillName ||
      !skillLevel
    ) {

      return {
        success: false,
        message: 'Please fill all required fields.'
      };

    }


    // -----------------------------
    // CHECK STUDENT
    // -----------------------------

    const studentLastRow =
      studentSheet.getLastRow();


    if (studentLastRow < 2) {

      return {
        success: false,
        message: 'No students found.'
      };

    }


    const students =
      studentSheet
        .getRange(
          2,
          1,
          studentLastRow - 1,
          15
        )
        .getValues();


    let studentFound = false;


    for (let i = 0; i < students.length; i++) {

      const studentUSN =
        String(students[i][0] || '')
          .trim()
          .toUpperCase();


      // Department column = H = index 7
      const studentDepartment =
        String(students[i][7] || '')
          .trim()
          .toUpperCase();


      if (studentUSN === usn) {

        if (
          studentDepartment !==
          String(currentUser.department || '')
            .trim()
            .toUpperCase()
        ) {

          return {
            success: false,
            message:
              'This student does not belong to your department.'
          };

        }


        studentFound = true;

        break;
      }

    }


    if (!studentFound) {

      return {
        success: false,
        message:
          'Student USN not found in Student_Master.'
      };

    }


    // -----------------------------
    // ADD SKILL
    // -----------------------------

    const today =
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        'yyyy-MM-dd'
      );


    skillsSheet.appendRow([

      usn,               // USN
      skillName,         // Skill_Name
      skillLevel,        // Skill_Level
      verified,          // Verified
      yearsExperience,   // Years_of_Experience
      today              // Last_Updated

    ]);


    return {
      success: true,
      message: 'Skill added successfully.'
    };


  }
  catch (error) {

    return {
      success: false,
      message: error.message
    };

  }

}


function addTrainingCertificate(training, user) {

  try {

    if (!training) {
      throw new Error('Certificate data is missing.');
    }

    if (!user || !user.role) {
      throw new Error('Invalid user session.');
    }

    const usn =
      String(training.usn || '')
        .trim()
        .toUpperCase();

    if (!usn) {
      throw new Error('USN is required.');
    }

    const ss = getSpreadsheet();

    if (!ss) {
      throw new Error('Unable to open SIDSS database.');
    }

    const sheet =
      ss.getSheetByName(
        CONFIG.SHEETS.TRAINING
      );

    if (!sheet) {
      throw new Error(
        'Training_Certifications sheet not found.'
      );
    }


    // Check student
    const students =
      getSheetData(
        CONFIG.SHEETS.STUDENTS
      );

    const student =
      students.find(function(row) {

        const rowUSN =
          getBackendField(row, [
            'USN',
            'Student_ID',
            'Student_Id'
          ]);

        return String(rowUSN || '')
          .trim()
          .toUpperCase() === usn;

      });


    if (!student) {
      throw new Error(
        'Student USN not found.'
      );
    }


    // Department restriction
    if (
      user.role === 'Department Admin' ||
      user.role === 'HOD'
    ) {

      const studentDepartment =
        String(
          getBackendField(student, [
            'Department',
            'Dept',
            'Branch'
          ]) || ''
        )
          .trim()
          .toUpperCase();

      const userDepartment =
        String(user.department || '')
          .trim()
          .toUpperCase();


      if (
        studentDepartment !==
        userDepartment
      ) {

        throw new Error(
          'Student does not belong to your department.'
        );

      }

    }


    // IMPORTANT:
    // Store date as text to avoid date serialization problems
    const completionDate =
      String(
        training.completionDate || ''
      ).trim();


    sheet.appendRow([

      usn,

      String(
        training.category || ''
      ).trim(),

      String(
        training.certificateName || ''
      ).trim(),

      String(
        training.provider || ''
      ).trim(),

      completionDate,

      String(
        training.certificateID || ''
      ).trim(),

      String(
        training.certificateLink || ''
      ).trim(),

      String(
        training.status || ''
      ).trim()

    ]);


    return {

      success: true,

      message:
        'Certificate added successfully.'

    };


  } catch (error) {

    return {

      success: false,

      message:
        error && error.message
          ? error.message
          : String(error)

    };

  }

}


function addEvent(eventData, user) {

  try {

    // -----------------------------
    // VALIDATE USER
    // -----------------------------

    if (!user || !user.role) {
      return {
        success: false,
        message: 'Invalid user session.'
      };
    }


    if (
      user.role !== 'Department Admin' &&
      user.role !== 'HOD'
    ) {

      return {
        success: false,
        message: 'You are not authorized to add events.'
      };

    }


    // -----------------------------
    // VALIDATE DATA
    // -----------------------------

    if (!eventData) {

      return {
        success: false,
        message: 'Event data is missing.'
      };

    }


    // -----------------------------
    // OPEN DATABASE
    // -----------------------------

    const ss = getSpreadsheet();

    const sheet =
      ss.getSheetByName('Events');


    if (!sheet) {

      return {
        success: false,
        message: 'Events sheet was not found.'
      };

    }


    // -----------------------------
    // READ VALUES
    // -----------------------------

    const eventName =
      String(
        eventData.eventName || ''
      ).trim();


    const eventType =
      String(
        eventData.eventType || ''
      ).trim();


    const eventDate =
      String(
        eventData.eventDate || ''
      ).trim();


    const venue =
      String(
        eventData.venue || ''
      ).trim();


    const coordinator =
      String(
        eventData.coordinator || ''
      ).trim();


    /*
     IMPORTANT:
     Keep this as text.

     Your form may contain:
     CSE 1st Year

     So DON'T use Number() here.
    */

    const eligibleStudents =
      String(
        eventData.eligibleStudents || ''
      ).trim();


    /*
     Participants should be numeric.
    */

    let participants =
      Number(
        eventData.participants || 0
      );


    if (isNaN(participants)) {
      participants = 0;
    }


    const status =
      String(
        eventData.status || ''
      ).trim();


    const remarks =
      String(
        eventData.remarks || ''
      ).trim();


    // -----------------------------
    // REQUIRED FIELDS
    // -----------------------------

    if (!eventName) {

      return {
        success: false,
        message: 'Event Name is required.'
      };

    }


    if (!eventType) {

      return {
        success: false,
        message: 'Event Type is required.'
      };

    }


    if (!eventDate) {

      return {
        success: false,
        message: 'Event Date is required.'
      };

    }


    if (!venue) {

      return {
        success: false,
        message: 'Venue is required.'
      };

    }


    if (!coordinator) {

      return {
        success: false,
        message: 'Coordinator is required.'
      };

    }


    if (!status) {

      return {
        success: false,
        message: 'Status is required.'
      };

    }


    // -----------------------------
    // DEPARTMENT
    // -----------------------------

    const department =
      String(
        user.department || ''
      )
      .trim()
      .toUpperCase();


    if (!department) {

      return {
        success: false,
        message: 'Department could not be determined.'
      };

    }


    // -----------------------------
    // GENERATE EVENT ID
    // -----------------------------

    const lastRow =
      sheet.getLastRow();


    let highestNumber = 0;


    if (lastRow > 1) {

      const ids =
        sheet
          .getRange(
            2,
            1,
            lastRow - 1,
            1
          )
          .getDisplayValues();


      ids.forEach(function(row) {

        const id =
          String(row[0] || '')
            .trim()
            .toUpperCase();


        const match =
          id.match(/^EVT(\d+)$/);


        if (match) {

          const num =
            parseInt(
              match[1],
              10
            );


          if (num > highestNumber) {
            highestNumber = num;
          }

        }

      });

    }


    const eventId =
      'EVT' +
      String(
        highestNumber + 1
      ).padStart(3, '0');


    // -----------------------------
    // WRITE TO GOOGLE SHEET
    // -----------------------------

    const newRow = [

      eventId,          // A Event_ID

      eventName,        // B Event_Name

      eventType,        // C Event_Type

      department,       // D Department

      eventDate,        // E Event_Date

      venue,            // F Venue

      coordinator,      // G Coordinator

      eligibleStudents, // H Eligible_Students

      participants,     // I Participants

      status,           // J Status

      remarks           // K Remarks

    ];


    sheet.appendRow(newRow);


    // -----------------------------
    // RETURN IMMEDIATELY
    // -----------------------------

    return {

      success: true,

      message:
        eventId +
        ' added successfully.',

      eventId: eventId

    };


  } catch (error) {

    console.error(
      'addEvent ERROR:',
      error
    );


    return {

      success: false,

      message:
        error && error.message
          ? error.message
          : 'Unable to add event.'

    };

  }

}


function addPlacement(placementData, user) {

  try {

    if (!user || !user.role) {
      throw new Error('Invalid user session.');
    }


    const role =
      String(user.role || '')
        .trim()
        .toLowerCase();


    if (
      role !== 'department admin' &&
      role !== 'hod' &&
      role !== 'placement officer'
    ) {
      throw new Error(
        'You are not authorized to add placement records.'
      );
    }


    if (!placementData) {
      throw new Error(
        'Placement data is missing.'
      );
    }


    const usn =
      String(placementData.usn || '')
        .trim()
        .toUpperCase();


    if (!usn) {
      throw new Error(
        'Student USN is required.'
      );
    }


    const spreadsheet =
      getSpreadsheet();


    const sheet =
      spreadsheet.getSheetByName(
        'Placement'
      );


    if (!sheet) {
      throw new Error(
        'Placement sheet was not found.'
      );
    }


    /*
     * CHECK DUPLICATE USN
     */

    const lastRow =
      sheet.getLastRow();


    if (lastRow >= 2) {

      const existingUSNs =
        sheet
          .getRange(
            2,
            1,
            lastRow - 1,
            1
          )
          .getValues();


      const duplicate =
        existingUSNs.some(function(row) {

          return String(row[0] || '')
            .trim()
            .toUpperCase() === usn;

        });


      if (duplicate) {

        throw new Error(
          'Placement record already exists for ' +
          usn +
          '.'
        );

      }

    }


    const eligibility =
      String(
        placementData.eligibility || 'No'
      ).trim();


    const resumeLink =
      String(
        placementData.resumeLink || ''
      ).trim();


    const aptitude =
      Number(
        placementData.aptitude || 0
      );


    const coding =
      Number(
        placementData.coding || 0
      );


    const communication =
      Number(
        placementData.communication || 0
      );


    const interviewStatus =
      String(
        placementData.interviewStatus ||
        'Not Applied'
      ).trim();


    const companiesApplied =
      String(
        placementData.companiesApplied || ''
      ).trim();


    const selectedCompany =
      String(
        placementData.selectedCompany || ''
      ).trim();


    const offerStatus =
      String(
        placementData.offerStatus || 'No'
      ).trim();


    const packageLPA =
      Number(
        placementData.packageLPA || 0
      );


    /*
     * VALIDATE SCORES
     */

    if (
      aptitude < 0 ||
      aptitude > 100 ||
      coding < 0 ||
      coding > 100 ||
      communication < 0 ||
      communication > 100
    ) {

      throw new Error(
        'Scores must be between 0 and 100.'
      );

    }


    /*
     * CALCULATE READINESS
     *
     * Average of:
     * Aptitude
     * Coding
     * Communication
     */

    const readiness =
      Number(
        (
          (
            aptitude +
            coding +
            communication
          ) / 3
        ).toFixed(1)
      );


    /*
     * PLACEMENT SHEET
     *
     * A  USN
     * B  Eligibility
     * C  Resume_Link
     * D  Aptitude_Score
     * E  Coding_Score
     * F  Communication_Score
     * G  Interview_Status
     * H  Companies_Applied
     * I  Selected_Company
     * J  Offer_Status
     * K  Package_LPA
     * L  Placement_Readiness_Score
     */


    sheet.appendRow([

      usn,
      eligibility,
      resumeLink,
      aptitude,
      coding,
      communication,
      interviewStatus,
      companiesApplied,
      selectedCompany,
      offerStatus,
      packageLPA,
      readiness

    ]);


    SpreadsheetApp.flush();


    return {

      success: true,

      message:
        'Placement record added successfully.',

      usn: usn,

      readiness: readiness

    };


  } catch (error) {

    console.error(
      'addPlacement error:',
      error
    );


    return {

      success: false,

      message:
        error && error.message
          ? error.message
          : 'Unable to add placement record.'

    };

  }

}

function updatePlacementRecord(
  placementData,
  user
) {

  try {

    if (!placementData) {
      throw new Error(
        'Placement data is missing.'
      );
    }


    if (!user || !user.role) {
      throw new Error(
        'Invalid user session.'
      );
    }


    const role =
      String(user.role || '')
        .trim()
        .toLowerCase();

    if (
      role !== 'department admin' &&
      role !== 'hod' &&
      role !== 'placement officer'
    ) {

      throw new Error(
        'You are not authorized to update placement records.'
      );

    }

    const usn =
      String(
        placementData.usn || ''
      )
        .trim()
        .toUpperCase();


    if (!usn) {

      throw new Error(
        'Student USN is required.'
      );

    }


    const spreadsheet =
      getSpreadsheet();


    if (!spreadsheet) {

      throw new Error(
        'Unable to open SIDSS database.'
      );

    }


    const sheet =
      spreadsheet.getSheetByName(
        'Placement'
      );


    if (!sheet) {

      throw new Error(
        'Placement sheet was not found.'
      );

    }


    const lastRow =
      sheet.getLastRow();


    if (lastRow < 2) {

      throw new Error(
        'No placement records found.'
      );

    }


    /*
     * YOUR PLACEMENT SHEET:
     *
     * A  USN
     * B  Eligibility
     * C  Resume_Link
     * D  Aptitude_Score
     * E  Coding_Score
     * F  Communication_Score
     * G  Interview_Status
     * H  Companies_Applied
     * I  Selected_Company
     * J  Offer_Status
     * K  Package_LPA
     * L  Placement_Readiness_Score
     */


    const usnValues =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          1
        )
        .getValues();


    let targetRow = -1;


    for (
      let i = 0;
      i < usnValues.length;
      i++
    ) {

      const sheetUSN =
        String(
          usnValues[i][0] || ''
        )
          .trim()
          .toUpperCase();


      if (sheetUSN === usn) {

        targetRow =
          i + 2;

        break;

      }

    }


    if (targetRow === -1) {

      throw new Error(
        'Placement record not found for ' +
        usn +
        '.'
      );

    }


    const aptitude =
      Number(
        placementData.aptitude || 0
      );


    const coding =
      Number(
        placementData.coding || 0
      );


    const communication =
      Number(
        placementData.communication || 0
      );


    /*
     * READINESS SCORE
     */

    const readiness =
      Math.round(
        (
          (
            aptitude +
            coding +
            communication
          ) / 3
        ) * 10
      ) / 10;


    /*
     * Update B:L.
     * USN in column A stays unchanged.
     */

    sheet
      .getRange(
        targetRow,
        2,
        1,
        11
      )
      .setValues([[

        String(
          placementData.eligibility ||
          'No'
        ),

        String(
          placementData.resumeLink ||
          ''
        ),

        aptitude,

        coding,

        communication,

        String(
          placementData.interviewStatus ||
          'Not Applied'
        ),

        String(
          placementData.companiesApplied ||
          ''
        ),

        String(
          placementData.selectedCompany ||
          ''
        ),

        String(
          placementData.offerStatus ||
          'No'
        ),

        Number(
          placementData.packageLPA || 0
        ),

        readiness

      ]]);


    SpreadsheetApp.flush();


    return {

      success: true,

      message:
        'Placement record updated successfully.',

      usn: usn,

      readiness: readiness

    };


  } catch (error) {

    console.error(
      'updatePlacementRecord error:',
      error
    );


    return {

      success: false,

      message:
        error && error.message
          ? error.message
          : 'Unable to update placement record.'

    };

  }

}


function updateEvent(eventData, user) {

  try {

    if (!user || !user.role) {
      throw new Error(
        'Invalid user session.'
      );
    }


    if (
      user.role !== 'Department Admin' &&
      user.role !== 'HOD'
    ) {

      throw new Error(
        'You are not authorized to edit events.'
      );

    }


    if (!eventData) {
      throw new Error(
        'Event data is missing.'
      );
    }


    const eventId =
      String(
        eventData.eventId || ''
      )
      .trim()
      .toUpperCase();


    if (!eventId) {
      throw new Error(
        'Event ID is missing.'
      );
    }


    const spreadsheet =
      getSpreadsheet();


    const sheet =
      spreadsheet.getSheetByName(
        'Events'
      );


    if (!sheet) {
      throw new Error(
        'Events sheet was not found.'
      );
    }


    const lastRow =
      sheet.getLastRow();


    if (lastRow < 2) {
      throw new Error(
        'No event records found.'
      );
    }


    // Find Event_ID in column A

    const ids =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          1
        )
        .getValues();


    let sheetRow = -1;


    for (
      let i = 0;
      i < ids.length;
      i++
    ) {

      const currentId =
        String(ids[i][0] || '')
          .trim()
          .toUpperCase();


      if (currentId === eventId) {

        sheetRow =
          i + 2;

        break;

      }

    }


    if (sheetRow === -1) {

      throw new Error(
        'Event not found: ' +
        eventId
      );

    }


    /*
     * SECURITY:
     * Department Admin/HOD can edit only
     * events belonging to their department
     * or ALL department events.
     */

    const existingDepartment =
      String(
        sheet
          .getRange(sheetRow, 4)
          .getValue() || ''
      )
      .trim()
      .toUpperCase();


    const userDepartment =
      String(
        user.department || ''
      )
      .trim()
      .toUpperCase();


    if (
      existingDepartment !== 'ALL' &&
      existingDepartment !== userDepartment
    ) {

      throw new Error(
        'You are not authorized to edit this event.'
      );

    }


    const eventName =
      String(
        eventData.eventName || ''
      ).trim();


    const eventType =
      String(
        eventData.eventType || ''
      ).trim();


    const eventDate =
      String(
        eventData.eventDate || ''
      ).trim();


    const venue =
      String(
        eventData.venue || ''
      ).trim();


    const coordinator =
      String(
        eventData.coordinator || ''
      ).trim();


    const eligibleStudents =
      Number(
        eventData.eligibleStudents || 0
      );


    const participants =
      Number(
        eventData.participants || 0
      );


    const status =
      String(
        eventData.status || ''
      ).trim();


    const remarks =
      String(
        eventData.remarks || ''
      ).trim();


    if (!eventName) {
      throw new Error(
        'Event Name is required.'
      );
    }


    if (!eventType) {
      throw new Error(
        'Event Type is required.'
      );
    }


    if (!eventDate) {
      throw new Error(
        'Event Date is required.'
      );
    }


    if (!venue) {
      throw new Error(
        'Venue is required.'
      );
    }


    if (!coordinator) {
      throw new Error(
        'Coordinator is required.'
      );
    }


    if (!status) {
      throw new Error(
        'Status is required.'
      );
    }


    /*
     * Columns:
     *
     * A Event_ID
     * B Event_Name
     * C Event_Type
     * D Department
     * E Event_Date
     * F Venue
     * G Coordinator
     * H Eligible_Students
     * I Participants
     * J Status
     * K Remarks
     *
     * Keep A and D unchanged.
     */


    sheet
      .getRange(
        sheetRow,
        2,
        1,
        10
      )
      .setValues([[

        eventName,            // B
        eventType,            // C
        existingDepartment,   // D
        eventDate,            // E
        venue,                // F
        coordinator,          // G
        eligibleStudents,     // H
        participants,         // I
        status,               // J
        remarks               // K

      ]]);


    SpreadsheetApp.flush();


    return {

      success: true,

      message:
        eventId +
        ' updated successfully.',

      eventId:
        eventId

    };


  } catch (error) {

    console.error(
      'updateEvent error:',
      error
    );


    return {

      success: false,

      message:
        error && error.message
          ? error.message
          : 'Unable to update event.'

    };

  }

}


function getEventsData(user) {

  try {

    if (!user || !user.role) {
      throw new Error('Invalid user session.');
    }


    let events =
      getSheetData('Events');


    const department =
      String(user.department || 'ALL')
        .trim()
        .toUpperCase();


    /*
     * ROLE BASED FILTER
     */

    if (
      user.role === 'HOD' ||
      user.role === 'Department Admin'
    ) {

      events =
        events.filter(function(record) {

          const eventDepartment =
            String(
              getBackendField(
                record,
                ['Department']
              ) || ''
            )
            .trim()
            .toUpperCase();


          /*
           * Department Admin/HOD can see:
           *
           * CSE
           * CSE, ISE
           * ALL
           */

          if (eventDepartment === 'ALL') {
            return true;
          }


          const departments =
            eventDepartment
              .split(',')
              .map(function(item) {
                return item.trim();
              });


          return departments.includes(
            department
          );

        });

    }


    /*
     * BUILD EVENT RECORDS
     */

    const records =
      events.map(function(record) {

        return {

          eventId:
            getBackendField(
              record,
              ['Event_ID']
            ),

          eventName:
            getBackendField(
              record,
              ['Event_Name']
            ),

          eventType:
            getBackendField(
              record,
              ['Event_Type']
            ),

          department:
            getBackendField(
              record,
              ['Department']
            ),

          eventDate:
            formatEventDate(
              getBackendField(
                record,
                ['Event_Date']
              )
            ),

          venue:
            getBackendField(
              record,
              ['Venue']
            ),

          coordinator:
            getBackendField(
              record,
              ['Coordinator']
            ),

          eligibleStudents:
            getBackendField(
              record,
              ['Eligible_Students']
            ),

          participants:
            Number(
              getBackendField(
                record,
                ['Participants']
              ) || 0
            ),

          status:
            getBackendField(
              record,
              ['Status']
            ),

          remarks:
            getBackendField(
              record,
              ['Remarks']
            )

        };

      });


    /*
     * ANALYTICS
     */

    const totalEvents =
      records.length;


    const upcomingEvents =
      records.filter(function(record) {

        return String(record.status || '')
          .trim()
          .toLowerCase() === 'upcoming';

      }).length;


    const completedEvents =
      records.filter(function(record) {

        return String(record.status || '')
          .trim()
          .toLowerCase() === 'completed';

      }).length;


    const ongoingEvents =
      records.filter(function(record) {

        const status =
          String(record.status || '')
            .trim()
            .toLowerCase();

        return (
          status === 'ongoing' ||
          status === 'in progress'
        );

      }).length;


    const totalParticipants =
      records.reduce(function(total, record) {

        return total +
          Number(record.participants || 0);

      }, 0);


    /*
     * EVENT TYPE DISTRIBUTION
     */

    const typeCounts = {};


    records.forEach(function(record) {

      const type =
        String(
          record.eventType || 'Other'
        ).trim();


      if (!typeCounts[type]) {
        typeCounts[type] = 0;
      }


      typeCounts[type]++;

    });


    const eventTypes =
      Object.keys(typeCounts)

        .map(function(type) {

          return {

            type: type,

            count:
              typeCounts[type]

          };

        })

        .sort(function(a, b) {

          return b.count - a.count;

        });


    /*
     * FINAL RESPONSE
     */

    return {

      success: true,

      scope:
        (
          user.role === 'HOD' ||
          user.role === 'Department Admin'
        )
          ? department
          : 'Institution',

      totalEvents:
        totalEvents,

      upcomingEvents:
        upcomingEvents,

      completedEvents:
        completedEvents,

      ongoingEvents:
        ongoingEvents,

      totalParticipants:
        totalParticipants,

      eventTypes:
        eventTypes,

      records:
        records

    };


  } catch (error) {

    console.error(
      'getEventsData error:',
      error
    );


    return {

      success: false,

      message:
        error.message ||
        'Unable to load events data.'

    };

  }

}

function formatEventDate(value) {

  if (!value) {
    return '';
  }


  try {

    if (
      Object.prototype.toString.call(value) ===
      '[object Date]'
    ) {

      return Utilities.formatDate(
        value,
        Session.getScriptTimeZone(),
        'yyyy-MM-dd'
      );

    }


    return String(value);

  } catch (error) {

    return String(value);

  }

}


function getFacultyDashboardStats(facultyId) {

  // Use the actual sheet names directly
  const facultySubjects =
    getSheetData("Faculty_Subjects");

  const students =
    getSheetData("Student_Master");

  const marks =
    getSheetData("Subject_Marks");


  const facultyIdStr =
    String(facultyId || "").trim();


  // Find the faculty's assigned subject
  const assignment =
    facultySubjects.find(function(row) {

      return String(row.Faculty_ID || "").trim()
        === facultyIdStr;

    });


  // If no assignment is found
  if (!assignment) {

    return {
      totalStudents: 0,
      totalSubjects: 0,
      updated: 0,
      pending: 0,
      pass: 0,
      fail: 0,
      department: "",
      semester: "",
      subjectCode: "",
      subjectName: ""
    };

  }


  const department =
    String(assignment.Department || "").trim();

  const semester =
    String(assignment.Semester || "").trim();

  const subjectCode =
    String(assignment.Subject_Code || "").trim();

  const subjectName =
    String(assignment.Subject_Name || "").trim();


  // Students in this faculty's department + semester
  const facultyStudents =
    students.filter(function(student) {

      return (
        normalizeFacultyDepartment(
          student.Department
        ) ===
        normalizeFacultyDepartment(
          department
        ) &&

        String(student.Semester || "").trim() ===
        semester
      );

    });


  const totalStudents =
    facultyStudents.length;


  // Get only USNs belonging to these students
  const studentUSNs =
    new Set(
      facultyStudents.map(function(student) {

        return String(student.USN || "").trim();

      })
    );


  // Get marks only for this subject and these students
  const subjectMarks =
    marks.filter(function(mark) {

      const usn =
        String(mark.USN || "").trim();

      const code =
        String(mark.Subject_Code || "").trim();

      return (
        code === subjectCode &&
        studentUSNs.has(usn)
      );

    });


  // Count each student only once
  const updatedUSNs =
    new Set();


  subjectMarks.forEach(function(mark) {

    const usn =
      String(mark.USN || "").trim();

    const total =
      String(mark.Total || "").trim();


    if (
      usn &&
      total !== "" &&
      Number(total) > 0
    ) {

      updatedUSNs.add(usn);

    }

  });


  const updated =
    updatedUSNs.size;


  const pending =
    Math.max(
      totalStudents - updated,
      0
    );


  // PASS / FAIL
  const passedUSNs =
    new Set();

  const failedUSNs =
    new Set();


  subjectMarks.forEach(function(mark) {

    const usn =
      String(mark.USN || "").trim();

    const result =
      String(mark.Result || "")
        .trim()
        .toUpperCase();


    if (result === "PASS") {
      passedUSNs.add(usn);
    }

    if (result === "FAIL") {
      failedUSNs.add(usn);
    }

  });


  const passed =
    passedUSNs.size;

  const failed =
    failedUSNs.size;


  const passPercent =
    totalStudents > 0
      ? Math.round(
          (passed / totalStudents) * 100
        )
      : 0;


  const failPercent =
    totalStudents > 0
      ? Math.round(
          (failed / totalStudents) * 100
        )
      : 0;


  return {

    totalStudents:
      totalStudents,

    totalSubjects:
      1,

    updated:
      updated,

    pending:
      pending,

    pass:
      passPercent,

    fail:
      failPercent,

    department:
      department,

    semester:
      semester,

    subjectCode:
      subjectCode,

    subjectName:
      subjectName

  };

}


function normalizeFacultyDepartment(value) {

  const department =
    String(value || "")
      .trim()
      .toUpperCase();

  if (department === "AD") {
    return "AIDS";
  }

  if (department === "IS") {
    return "ISE";
  }

  if (department === "EC") {
    return "ECE";
  }

  if (department === "ME") {
    return "MECH";
  }

  return department;
}



function getFacultyDashboardDetails(facultyId, type) {

  // Use the actual sheet names directly
  const facultySubjects = getSheetData("Faculty_Subjects");
  const students = getSheetData("Student_Master");
  const marks = getSheetData("Subject_Marks");

  const facultyIdStr = String(facultyId || "").trim();

  // Get subjects assigned to this faculty
  const assignedSubjects = facultySubjects.filter(function(row) {

    return String(row.Faculty_ID || "").trim() === facultyIdStr;

  });


  if (assignedSubjects.length === 0) {
    return [];
  }


  // =====================================================
  // SUBJECTS ASSIGNED
  // =====================================================

  if (type === "subjects") {

    return assignedSubjects.map(function(subject) {

      return {

        subjectCode:
          String(subject.Subject_Code || "").trim(),

        subjectName:
          String(subject.Subject_Name || "").trim(),

        department:
          String(subject.Department || "").trim(),

        semester:
          String(subject.Semester || "").trim()

      };

    });

  }


  // =====================================================
  // STUDENT DETAILS
  // =====================================================

  const result = [];


  assignedSubjects.forEach(function(subject) {

    const department =
      String(subject.Department || "").trim();

    const semester =
      String(subject.Semester || "").trim();

    const subjectCode =
      String(subject.Subject_Code || "").trim();

    const subjectName =
      String(subject.Subject_Name || "").trim();


    // Students belonging to this faculty's
    // department + semester
    const matchingStudents =
      students.filter(function(student) {

        return (
          normalizeFacultyDepartment(
            student.Department
          ) ===
          normalizeFacultyDepartment(
            department
          ) &&

          String(student.Semester || "").trim() ===
          semester
        );

      });


    matchingStudents.forEach(function(student) {

      const studentUSN =
        String(student.USN || "").trim();


      // Find marks for this student + subject
      const studentMark =
        marks.find(function(mark) {

          return (
            String(mark.USN || "").trim() === studentUSN &&
            String(mark.Subject_Code || "").trim() === subjectCode
          );

        });


      const total =
        studentMark
          ? Number(studentMark.Total || 0)
          : 0;


      const resultStatus =
        studentMark
          ? String(studentMark.Result || "")
              .trim()
              .toUpperCase()
          : "";


      const hasMarks =
        studentMark &&
        (
          Number(studentMark.Internal_1 || 0) > 0 ||
          Number(studentMark.Internal_2 || 0) > 0 ||
          Number(studentMark.Assignment || 0) > 0 ||
          Number(studentMark.Report || 0) > 0 ||
          Number(studentMark.Lab || 0) > 0 ||
          Number(studentMark.External || 0) > 0 ||
          Number(studentMark.Total || 0) > 0
        );


      let include = false;


      // ALL STUDENTS
      if (type === "students") {
        include = true;
      }


      // MARKS UPDATED
      else if (type === "updated") {
        include = hasMarks;
      }


      // MARKS PENDING
      else if (type === "pending") {
        include = !hasMarks;
      }


      // PASSED
      else if (type === "pass") {
        include = resultStatus === "PASS";
      }


      // FAILED
      else if (type === "fail") {
        include = resultStatus === "FAIL";
      }


      if (include) {

        result.push({

          usn: studentUSN,

          studentName:
            String(student.Student_Name || "").trim(),

          department: department,

          semester: semester,

          subjectCode: subjectCode,

          subjectName: subjectName,

          cie: studentMark
            ? Number(studentMark.CIE || 0)
            : 0,

          external: studentMark
            ? Number(studentMark.External_Reduced || 0)
            : 0,

          total: total,

          result: resultStatus || "-"

        });

      }

    });

  });


  return result;
}


function getSubjectEvaluation(subjectCode) {

  // Use the configured sheet name if available.
  // Otherwise use the actual Subject_Evaluation sheet name.
  const sheetName =
    CONFIG &&
    CONFIG.SHEETS &&
    CONFIG.SHEETS.SUBJECT_EVALUATION
      ? CONFIG.SHEETS.SUBJECT_EVALUATION
      : "Subject_Evaluation";


  const rows =
    getSheetData(sheetName);


  const targetCode =
    String(subjectCode || "")
      .trim()
      .toUpperCase();


  const subject =
    rows.find(function(row) {

      return String(
        row.Subject_Code || ""
      )
        .trim()
        .toUpperCase() === targetCode;

    });


  if (!subject) {

    throw new Error(
      "Evaluation structure not found for subject: " +
      subjectCode
    );

  }


  return subject;

}



function saveFacultyMarksToSheet(marks) {

  try {

    const ss =
      getSpreadsheet();

    const sheet =
      ss.getSheetByName(
        "Subject_Marks"
      );


    if (!sheet) {

      throw new Error(
        "Subject_Marks sheet not found."
      );

    }


    if (!marks) {

      throw new Error(
        "Marks data is missing."
      );

    }


    const usn =
      String(
        marks.usn || ""
      ).trim();


    const subjectCode =
      String(
        marks.subjectCode || ""
      ).trim();


    if (!usn) {

      throw new Error(
        "USN is missing."
      );

    }


    if (!subjectCode) {

      throw new Error(
        "Subject Code is missing."
      );

    }


    // ========================================
    // READ MARKS
    // ========================================

    const internal1 =
      Number(marks.internal1 || 0);

    const internal2 =
      Number(marks.internal2 || 0);

    const assignment =
      Number(marks.assignment || 0);

    const report =
      Number(marks.report || 0);

    const lab =
      Number(marks.lab || 0);

    const external =
      Number(marks.external || 0);


    // ========================================
    // VALIDATION
    // ========================================

    if (
      internal1 < 0 ||
      internal1 > 50
    ) {

      throw new Error(
        "Internal 1 must be between 0 and 50."
      );

    }


    if (
      internal2 < 0 ||
      internal2 > 50
    ) {

      throw new Error(
        "Internal 2 must be between 0 and 50."
      );

    }


    if (
      marks.hasAssignment &&
      (
        assignment < 0 ||
        assignment > 25
      )
    ) {

      throw new Error(
        "Assignment must be between 0 and 25."
      );

    }


    if (
      marks.hasReport &&
      (
        report < 0 ||
        report > 15
      )
    ) {

      throw new Error(
        "Report must be between 0 and 15."
      );

    }


    if (
      marks.hasLab &&
      (
        lab < 0 ||
        lab > 25
      )
    ) {

      throw new Error(
        "Lab must be between 0 and 25."
      );

    }


    if (
      external < 0 ||
      external > 100
    ) {

      throw new Error(
        "External must be between 0 and 100."
      );

    }


    // ========================================
    // INTERNAL REDUCED
    // ========================================

    const internalReduced =
      Math.round(
        (internal1 + internal2) / 4
      );

    // ========================================
    // CIE
    // ========================================

    let cie =
      internalReduced;


    if (marks.hasAssignment) {

      cie += assignment;

    }


    if (marks.hasReport) {

      cie += report;

    }


    if (marks.hasLab) {

      cie += lab;

    }


    // ========================================
    // EXTERNAL REDUCED
    // ========================================

    const externalReduced =
      Math.round(
        external / 2
      );


    // ========================================
    // TOTAL
    // ========================================

    const total =
      cie + externalReduced;


    // ========================================
    // RESULT
    // ========================================

    const result =
      external === 0
        ? ""
        : (
            total >= 40 &&
            externalReduced >= 18
          )
            ? "Pass"
            : "Fail";


    // ========================================
    // FIND EXISTING ROW
    // ========================================

    const lastRow =
      sheet.getLastRow();


    if (lastRow < 2) {

      throw new Error(
        "Subject_Marks has no student rows."
      );

    }


    const data =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          3
        )
        .getValues();


    let existingRow =
      -1;


    for (
      let i = 0;
      i < data.length;
      i++
    ) {

      const sheetUSN =
        String(
          data[i][0] || ""
        ).trim();


      const sheetSubjectCode =
        String(
          data[i][2] || ""
        ).trim();


      if (
        sheetUSN === usn &&
        sheetSubjectCode === subjectCode
      ) {

        existingRow =
          i + 2;

        break;

      }

    }


    if (existingRow === -1) {

      throw new Error(
        "Student + Subject row not found in Subject_Marks."
        + "\nUSN: " + usn
        + "\nSubject: " + subjectCode
      );

    }


    // ========================================
    // WRITE EXACTLY E:O
    // ========================================

    const values = [

      internal1,

      internal2,

      internalReduced,

      marks.hasAssignment
        ? assignment
        : 0,

      marks.hasReport
        ? report
        : 0,

      marks.hasLab
        ? lab
        : 0,

      cie,

      external,

      externalReduced,

      total,

      result

    ];


    sheet
      .getRange(
        existingRow,
        5,
        1,
        11
      )
      .setValues([
        values
      ]);

    SpreadsheetApp.flush();


    // ========================================
    // VERIFY WRITE
    // ========================================




    return {

      success: true,

      message:
        "Marks saved successfully.",

      row:
        existingRow,

      internalReduced:
        internalReduced,

      cie:
        cie,

      externalReduced:
        externalReduced,

      total:
        total,

      result:
        result

    };


  } catch (error) {

    console.error(
      "saveFacultyMarksToSheet ERROR:",
      error
    );


    return {

      success: false,

      message:
        error &&
        error.message
          ? error.message
          : String(error)

    };

  }

}


function getFacultySubjectConfig(subjectCode) {

  const ss = SpreadsheetApp.openById("1Sn6HAHna6xpx4D3ntLZiSu5JRvcR8U-DfLnSI3BU1Og");

  const sheets = ss.getSheets();

  for (let s = 0; s < sheets.length; s++) {

    const sheet = sheets[s];
    const data = sheet.getDataRange().getValues();

    if (!data.length) continue;

    const headers = data[0].map(function(h) {
      return String(h).trim();
    });

    // Find the SUBJECT CONFIGURATION sheet
    const codeCol = headers.indexOf("Subject_Code");
    const categoryCol = headers.indexOf("Subject_Category");

    if (codeCol === -1 || categoryCol === -1) {
      continue;
    }

    const assignmentCol = headers.indexOf("Assignment_Max");
    const reportCol = headers.indexOf("Report_Max");
    const labCol = headers.indexOf("Lab_Max");
    const internal1Col = headers.indexOf("Internal_1_Max");
    const internal2Col = headers.indexOf("Internal_2_Max");

    for (let i = 1; i < data.length; i++) {

      if (
        String(data[i][codeCol]).trim() ===
        String(subjectCode).trim()
      ) {

        return {
          success: true,

          subjectCode: subjectCode,

          subjectCategory:
            String(data[i][categoryCol] || ""),

          internal1Max:
            Number(data[i][internal1Col] || 0),

          internal2Max:
            Number(data[i][internal2Col] || 0),

          assignmentMax:
            Number(data[i][assignmentCol] || 0),

          reportMax:
            Number(data[i][reportCol] || 0),

          labMax:
            Number(data[i][labCol] || 0)
        };
      }
    }
  }

  return {
    success: false,
    message: "Subject configuration not found for " + subjectCode
  };
}



function getSubjectCategory(subjectCode) {

  const sheet = getSpreadsheet().getSheetByName("Subject_Evaluation");

  if (!sheet) {
    throw new Error("Subject_Evaluation sheet not found.");
  }

  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return "";
  }

  const headers = values[0].map(function(header) {
    return String(header).trim();
  });

  const codeIndex = headers.indexOf("Subject_Code");
  const categoryIndex = headers.indexOf("Subject_Category");

  if (codeIndex === -1) {
    throw new Error("Subject_Code column not found.");
  }

  if (categoryIndex === -1) {
    throw new Error("Subject_Category column not found.");
  }

  const wantedCode = String(subjectCode || "").trim();

  for (let i = 1; i < values.length; i++) {

    const currentCode =
      String(values[i][codeIndex] || "").trim();

    if (currentCode === wantedCode) {

      return String(
        values[i][categoryIndex] || ""
      ).trim();

    }
  }

  return "";
}


function getExistingFacultyMarks(usn, subjectCode) {

  const sheet =
    getSpreadsheet().getSheetByName("Subject_Marks");

  if (!sheet) {

    return {
      success: false,
      message: "Subject_Marks sheet not found."
    };

  }


  const data =
    sheet.getDataRange().getDisplayValues();


  function normalize(value) {

    return String(value || "")
      .trim()
      .toUpperCase()
      .replace(/\s+/g, "");

  }


  const targetUSN =
    normalize(usn);

  const targetSubjectCode =
    normalize(subjectCode);


  for (let i = 1; i < data.length; i++) {

    const sheetUSN =
      normalize(data[i][0]);

    const sheetSubjectCode =
      normalize(data[i][2]);


    if (
      sheetUSN === targetUSN &&
      sheetSubjectCode === targetSubjectCode
    ) {

      return {

        success: true,

        // E - Internal 1
        internal1:
          data[i][4] !== ""
            ? data[i][4]
            : "",

        // F - Internal 2
        internal2:
          data[i][5] !== ""
            ? data[i][5]
            : "",

        // G - Internal Reduced
        internalReduced:
          data[i][6] !== ""
            ? data[i][6]
            : "",

        // H - Assignment
        assignment:
          data[i][7] !== ""
            ? data[i][7]
            : "",

        // I - Report
        report:
          data[i][8] !== ""
            ? data[i][8]
            : "",

        // J - Lab
        lab:
          data[i][9] !== ""
            ? data[i][9]
            : "",

        // K - CIE
        cie:
          data[i][10] !== ""
            ? data[i][10]
            : "",

        // L - External
        external:
          data[i][11] !== ""
            ? data[i][11]
            : "",

        // M - External Reduced
        externalReduced:
          data[i][12] !== ""
            ? data[i][12]
            : "",

        // N - Total
        total:
          data[i][13] !== ""
            ? data[i][13]
            : "",

        // O - Result
        result:
          data[i][14] !== ""
            ? data[i][14]
            : ""

      };

    }

  }


  // No marks entered yet

  return {

    success: true,

    internal1: "",
    internal2: "",
    internalReduced: "",

    assignment: "",
    report: "",
    lab: "",

    cie: "",

    external: "",
    externalReduced: "",

    total: "",
    result: ""

  };

}



function getFilterBuilderStudents(user) {

  if (!user || !user.role) {
    throw new Error("User information is required.");
  }


  // =====================================================
  // READ PHASE-1 DATA
  // =====================================================

  const students =
    getSheetData("Student_Master");

  const academics =
    getSheetData("Academics");

  const attendance =
    getSheetData("Attendance");

  const subjectMarks =
    getSheetData("Subject_Marks");

  const placement =
    getSheetData("Placement");

  const skills =
    getSheetData("Skills");

  const training =
    getSheetData("Training_Certifications");


  // =====================================================
  // HELPERS
  // =====================================================

  function text(value) {

    return String(value == null ? "" : value)
      .trim();

  }


  function upper(value) {

    return text(value)
      .toUpperCase();

  }


  function number(value) {

    if (
      value === "" ||
      value === null ||
      value === undefined
    ) {
      return null;
    }

    const n = Number(value);

    return Number.isFinite(n)
      ? n
      : null;

  }


  function field(row, names) {

    if (!row) {
      return "";
    }

    for (
      let i = 0;
      i < names.length;
      i++
    ) {

      const name = names[i];

      if (
        Object.prototype.hasOwnProperty.call(
          row,
          name
        )
      ) {

        const value = row[name];

        if (
          value !== "" &&
          value !== null &&
          value !== undefined
        ) {

          return value;

        }

      }

    }

    return "";

  }


  function usnOf(row) {

    return upper(
      field(
        row,
        [
          "USN",
          "Student_USN",
          "Student_ID",
          "Student_Id"
        ]
      )
    );

  }


  function recordsForUSN(data, usn) {

    return data.filter(function(row) {

      return usnOf(row) === usn;

    });

  }


  function latestRecord(records) {

    if (
      !records ||
      !records.length
    ) {

      return null;

    }

    return records[records.length - 1];

  }


  // =====================================================
  // BUILD LOOKUP MAPS
  // =====================================================

  function buildMap(data) {

    const map = {};

    data.forEach(function(row) {

      const usn =
        usnOf(row);

      if (!usn) {
        return;
      }

      if (!map[usn]) {
        map[usn] = [];
      }

      map[usn].push(row);

    });

    return map;

  }


  const academicMap =
    buildMap(academics);

  const attendanceMap =
    buildMap(attendance);

  const subjectMarksMap =
    buildMap(subjectMarks);

  const placementMap =
    buildMap(placement);

  const skillsMap =
    buildMap(skills);

  const trainingMap =
    buildMap(training);


  // =====================================================
  // ROLE SCOPE
  // =====================================================

  const role =
    text(user.role)
      .toLowerCase();

  const userDepartment =
    upper(user.department);


  const departmentScoped =
    role === "hod" ||
    role === "department admin";


  // =====================================================
  // BUILD STUDENT RECORDS
  // =====================================================

  const result = [];


  students.forEach(function(student) {

    const usn =
      usnOf(student);

    if (!usn) {
      return;
    }


    const department =
      text(
        field(
          student,
          [
            "Department",
            "Dept"
          ]
        )
      );


    // ---------------------------------------------------
    // SERVER-SIDE DEPARTMENT SECURITY
    // ---------------------------------------------------

    if (
      departmentScoped &&
      upper(department) !== userDepartment
    ) {

      return;

    }


    const studentAcademics =
      academicMap[usn] || [];

    const studentAttendance =
      attendanceMap[usn] || [];

    const studentSubjectMarks =
      subjectMarksMap[usn] || [];

    const studentPlacement =
      placementMap[usn] || [];

    const studentSkills =
      skillsMap[usn] || [];

    const studentTraining =
      trainingMap[usn] || [];


    // ===================================================
    // ACADEMIC
    // ===================================================

    const academicLatest =
      latestRecord(
        studentAcademics
      );


    const cgpa =
      number(
        field(
          academicLatest,
          [
            "CGPA",
            "Current_CGPA",
            "Overall_CGPA"
          ]
        )
      );


    const sgpa =
      number(
        field(
          academicLatest,
          [
            "SGPA"
          ]
        )
      );


    const backlogs =
      number(
        field(
          academicLatest,
          [
            "Backlogs",
            "Backlog_Count",
            "Backlog"
          ]
        )
      );


    const classRank =
      number(
        field(
          academicLatest,
          [
            "Class_Rank",
            "Class Rank"
          ]
        )
      );


    const branchRank =
      number(
        field(
          academicLatest,
          [
            "Branch_Rank",
            "Branch Rank"
          ]
        )
      );


    // ===================================================
    // ATTENDANCE
    // ===================================================

    const attendanceValues =
      studentAttendance
        .map(function(row) {

          return number(
            field(
              row,
              [
                "Overall_Attendance_%",
                "Overall %",
                "Attendance_Percentage",
                "Attendance",
                "Percentage"
              ]
            )
          );

        })
        .filter(function(value) {

          return value !== null;

        });


    let overallAttendance = null;


    if (attendanceValues.length) {

      overallAttendance =
        attendanceValues.reduce(
          function(total, value) {

            return total + value;

          },
          0
        ) /
        attendanceValues.length;

      overallAttendance =
        Number(
          overallAttendance.toFixed(1)
        );

    }


    // ===================================================
    // ATTENDANCE DEFAULTER
    // ===================================================

    let defaulterStatus = "";


    const latestAttendance =
      latestRecord(
        studentAttendance
      );


    if (latestAttendance) {

      const rawDefaulter =
        field(
          latestAttendance,
          [
            "Defaulter",
            "Defaulter_Status",
            "Defaulter Status"
          ]
        );


      if (text(rawDefaulter)) {

        defaulterStatus =
          text(rawDefaulter)
            .toLowerCase();

      }

      else if (
        overallAttendance !== null
      ) {

        defaulterStatus =
          overallAttendance < 75
            ? "yes"
            : "no";

      }

    }


    // ===================================================
    // ATTENDANCE TREND
    // ===================================================

    let attendanceTrend = "";


    if (
      attendanceValues.length >= 2
    ) {

      const previous =
        attendanceValues[
          attendanceValues.length - 2
        ];

      const current =
        attendanceValues[
          attendanceValues.length - 1
        ];


      if (current > previous) {

        attendanceTrend =
          "improving";

      }

      else if (current < previous) {

        attendanceTrend =
          "declining";

      }

      else {

        attendanceTrend =
          "stable";

      }

    }


    // ===================================================
    // PLACEMENT
    // ===================================================

    const placementLatest =
      latestRecord(
        studentPlacement
      );


    const eligibility =
      text(
        field(
          placementLatest,
          [
            "Eligibility",
            "Placement_Eligible",
            "Placement Eligibility"
          ]
        )
      );


    const companiesApplied =
      text(
        field(
          placementLatest,
          [
            "Companies_Applied",
            "Companies Applied"
          ]
        )
      );


    const offerStatus =
      text(
        field(
          placementLatest,
          [
            "Offer_Status",
            "Offer Status"
          ]
        )
      );


    const packageLPA =
      number(
        field(
          placementLatest,
          [
            "Package_LPA",
            "Package",
            "Package LPA"
          ]
        )
      );


    // ===================================================
    // SKILLS
    // ===================================================

    const skillRecords =
      studentSkills.map(function(row) {

        return {

          skill:
            text(
              field(
                row,
                [
                  "Skill",
                  "Skill_Name",
                  "Skill Name"
                ]
              )
            ),

          proficiency:
            text(
              field(
                row,
                [
                  "Proficiency_Level",
                  "Proficiency Level",
                  "Proficiency"
                ]
              )
            ),

          verified:
            text(
              field(
                row,
                [
                  "Verified",
                  "Verified_YN"
                ]
              )
            )

        };

      });


    // ===================================================
    // TRAINING / CERTIFICATIONS
    // ===================================================

    const trainingRecords =
      studentTraining.map(function(row) {

        return {

          category:
            text(
              field(
                row,
                [
                  "Category"
                ]
              )
            ),

          title:
            text(
              field(
                row,
                [
                  "Title",
                  "Certification",
                  "Certificate"
                ]
              )
            ),

          provider:
            text(
              field(
                row,
                [
                  "Provider"
                ]
              )
            )

        };

      });


    const certifications =
      trainingRecords.filter(function(item) {

        return (
          item.category
            .toLowerCase()
            .indexOf("cert") !== -1 ||

          item.category
            .toLowerCase()
            .indexOf("nptel") !== -1 ||

          item.category
            .toLowerCase()
            .indexOf("coursera") !== -1 ||

          item.category
            .toLowerCase()
            .indexOf("aws") !== -1 ||

          item.category
            .toLowerCase()
            .indexOf("cisco") !== -1

        );

      });


    const hackathonParticipation =
      trainingRecords.some(function(item) {

        return (
          item.category
            .toLowerCase()
            .indexOf("hackathon") !== -1
        );

      });


    // ===================================================
    // DEMOGRAPHIC / ADMINISTRATIVE
    // ===================================================

    const branch =
      text(
        field(
          student,
          [
            "Branch",
            "Program",
            "Course"
          ]
        )
      );


    const batch =
      text(
        field(
          student,
          [
            "Batch"
          ]
        )
      );


    const section =
      text(
        field(
          student,
          [
            "Section"
          ]
        )
      );


    /*
     * Restricted fields are intentionally
     * returned only to Institution Admin
     * and Principal.
     */

    const restrictedFieldsAllowed =
      role === "institution admin" ||
      role === "principal";


    const record = {

      // -----------------------------------------------
      // BASIC
      // -----------------------------------------------

      usn:
        usn,

      studentName:
        text(
          field(
            student,
            [
              "Student_Name",
              "Name",
              "Full_Name"
            ]
          )
        ),

      department:
        department,

      semester:
        text(
          field(
            student,
            [
              "Semester"
            ]
          )
        ),

      branch:
        branch,

      batch:
        batch,

      section:
        section,


      // -----------------------------------------------
      // ACADEMIC
      // -----------------------------------------------

      cgpa:
        cgpa,

      sgpa:
        sgpa,

      backlogs:
        backlogs,

      classRank:
        classRank,

      branchRank:
        branchRank,


      // -----------------------------------------------
      // ATTENDANCE
      // -----------------------------------------------

      overallAttendance:
        overallAttendance,

      defaulterStatus:
        defaulterStatus,

      attendanceTrend:
        attendanceTrend,


      // -----------------------------------------------
      // PLACEMENT
      // -----------------------------------------------

      eligibility:
        eligibility,

      companiesApplied:
        companiesApplied,

      offerStatus:
        offerStatus,

      packageLPA:
        packageLPA,


      // -----------------------------------------------
      // SUBJECT MARKS
      // -----------------------------------------------

      subjectMarks:
        studentSubjectMarks.map(
          function(row) {

            return {

              subjectCode:
                text(
                  field(
                    row,
                    [
                      "Subject_Code",
                      "Subject Code"
                    ]
                  )
                ),

              subjectName:
                text(
                  field(
                    row,
                    [
                      "Subject_Name",
                      "Subject Name"
                    ]
                  )
                ),

              total:
                number(
                  field(
                    row,
                    [
                      "Total"
                    ]
                  )
                ),

              result:
                text(
                  field(
                    row,
                    [
                      "Result"
                    ]
                  )
                )

            };

          }
        ),


      // -----------------------------------------------
      // SKILLS
      // -----------------------------------------------

      skills:
        skillRecords,


      // -----------------------------------------------
      // TRAINING
      // -----------------------------------------------

      certifications:
        certifications,

      hackathonParticipation:
        hackathonParticipation

    };


    // =================================================
    // RESTRICTED FIELDS
    // =================================================

    if (restrictedFieldsAllowed) {

      record.gender =
        text(
          field(
            student,
            [
              "Gender"
            ]
          )
        );


      record.category =
        text(
          field(
            student,
            [
              "Category"
            ]
          )
        );


      record.scholarshipStatus =
        text(
          field(
            student,
            [
              "Scholarship_Status",
              "Scholarship Status"
            ]
          )
        );


      record.incomeBracket =
        text(
          field(
            student,
            [
              "Income_Bracket",
              "Income Bracket"
            ]
          )
        );

    }


    result.push(record);

  });


  return result;

}

function calculateAttendanceTrend(records) {

  const monthly = {};

  (records || []).forEach(function(record) {

    const rawDate =
      record.Date ||
      record.date ||
      record.Attendance_Date ||
      record.attendanceDate ||
      record.Month ||
      record.month ||
      '';

    const rawAttendance =
      record.Attendance ||
      record.attendance ||
      record.Attendance_Percentage ||
      record.attendancePercentage ||
      record.Overall_Attendance ||
      record.overallAttendance ||
      '';

    const attendance =
      Number(
        String(rawAttendance)
          .replace('%', '')
          .trim()
      );

    if (!Number.isFinite(attendance)) {
      return;
    }

    let date = null;

    if (rawDate instanceof Date) {
      date = rawDate;
    } else if (rawDate) {
      date = new Date(rawDate);
    }

    if (
      !date ||
      isNaN(date.getTime())
    ) {
      return;
    }

    const monthKey =
      Utilities.formatDate(
        date,
        Session.getScriptTimeZone(),
        'yyyy-MM'
      );

    const monthLabel =
      Utilities.formatDate(
        date,
        Session.getScriptTimeZone(),
        'MMM yyyy'
      );

    if (!monthly[monthKey]) {

      monthly[monthKey] = {
        month: monthLabel,
        total: 0,
        count: 0
      };

    }

    monthly[monthKey].total += attendance;
    monthly[monthKey].count++;

  });

  return Object.keys(monthly)
    .sort()
    .map(function(key) {

      const item = monthly[key];

      return {
        month: item.month,
        average:
          item.count
            ? Number(
                (
                  item.total /
                  item.count
                ).toFixed(1)
              )
            : 0
      };

    });

}



function getAtRiskStudents(user) {

  try {

    if (!user || !user.role) {
      throw new Error('Invalid user session.');
    }

    let students =
      getSheetData(CONFIG.SHEETS.STUDENTS) || [];

    const academics =
      getSheetData(CONFIG.SHEETS.ACADEMICS) || [];

    const attendance =
      getSheetData(CONFIG.SHEETS.ATTENDANCE) || [];


    const department =
      String(user.department || 'ALL')
        .trim()
        .toUpperCase();


    // Department restriction
    if (
      user.role === 'HOD' ||
      user.role === 'Department Admin'
    ) {

      students =
        filterByDepartment(
          students,
          department
        );

    }


    // ============================================
    // USN HELPER
    // ============================================

    function getUSN(row) {

      return String(
        row.USN ||
        row.Student_ID ||
        row.Student_Id ||
        ''
      )
        .trim()
        .toUpperCase();

    }


    // ============================================
    // NUMBER HELPER
    // ============================================

    function toNumber(value) {

      if (
        value === null ||
        value === undefined ||
        value === ''
      ) {
        return null;
      }

      const number = Number(value);

      return Number.isFinite(number)
        ? number
        : null;

    }


    // ============================================
    // ACADEMIC MAP
    // READ EXACT SHEET COLUMNS
    // ============================================

    const academicMap = {};


    academics.forEach(function(record) {

      const usn = getUSN(record);

      if (!usn) {
        return;
      }


      // IMPORTANT:
      // Read EXACTLY from Academics sheet
      const cgpa =
        toNumber(record['CGPA']);


      const backlogs =
        toNumber(record['Backlogs']);


      academicMap[usn] = {

        cgpa: cgpa,

        backlogs: backlogs

      };

    });


    // ============================================
    // ATTENDANCE MAP
    // ============================================

    const attendanceMap = {};

    const attendanceTotals = {};
    const attendanceCounts = {};

    attendance.forEach(function(record) {

      const usn = getUSN(record);

      if (!usn) {
        return;
      }

      // EXACT column name from Attendance sheet
      const raw =
        record['Overall_Attendance_%'];

      if (
        raw === null ||
        raw === undefined ||
        raw === ''
      ) {
        return;
      }

      const percentage =
        Number(raw);

      if (!Number.isFinite(percentage)) {
        return;
      }

      if (!attendanceTotals[usn]) {
        attendanceTotals[usn] = 0;
        attendanceCounts[usn] = 0;
      }

      attendanceTotals[usn] += percentage;
      attendanceCounts[usn]++;

    });


    // Calculate average attendance for each student

    Object.keys(attendanceTotals).forEach(function(usn) {

      if (attendanceCounts[usn] > 0) {

        attendanceMap[usn] =
          attendanceTotals[usn] /
          attendanceCounts[usn];

      }

    });


    // ============================================
    // FIND AT-RISK STUDENTS
    // ============================================

    const atRiskStudents =
      students
        .map(function(student) {

          const usn =
            getUSN(student);


          const academic =
            academicMap[usn] || {};


          const cgpa =
            academic.cgpa;


          const backlogs =
            academic.backlogs;


          const attendancePercentage =
            attendanceMap[usn];


          const reasons = [];


          // ----------------------------------------
          // BACKLOG RISK
          // 2 OR MORE BACKLOGS
          // ----------------------------------------

          if (
            backlogs !== null &&
            backlogs >= 2
          ) {

            reasons.push({

              type: 'BACKLOG',

              label: 'Backlog Risk',

              value:
                backlogs +
                (
                  backlogs === 1
                    ? ' active backlog'
                    : ' active backlogs'
                )

            });

          }


          // ----------------------------------------
          // CGPA RISK
          // BELOW 7.5
          // ----------------------------------------

          if (
            cgpa !== null &&
            cgpa < 7.5
          ) {

            reasons.push({

              type: 'CGPA',

              label: 'CGPA Risk',

              value:
                'CGPA ' +
                Number(cgpa).toFixed(2)

            });

          }


          // ----------------------------------------
          // ATTENDANCE RISK
          // BELOW 75%
          // ----------------------------------------

          if (
            attendancePercentage !== null &&
            attendancePercentage < 75
          ) {

            reasons.push({

              type: 'ATTENDANCE',

              label: 'Attendance Risk',

              value:
                'Attendance ' +
                Number(
                  attendancePercentage
                ).toFixed(1) +
                '%'

            });

          }


          // No risk
          if (!reasons.length) {
            return null;
          }


          // ----------------------------------------
          // RISK LEVEL
          // ----------------------------------------

          let riskLevel =
            'MODERATE';


          if (reasons.length >= 2) {
            riskLevel = 'HIGH';
          }


          if (reasons.length >= 3) {
            riskLevel = 'CRITICAL';
          }


          return {

            ...student,

            _risk: {

              level:
                riskLevel,

              reasons:
                reasons,

              cgpa:
                cgpa,

              backlogs:
                backlogs,

              attendance:
                attendancePercentage

            }

          };

        })
        .filter(function(student) {

          return student !== null;

        });


    // ============================================
    // RETURN
    // ============================================

    return {

      success: true,

      scope:
        (
          user.role === 'HOD' ||
          user.role === 'Department Admin'
        )
          ? department
          : 'Institution',

      count:
        atRiskStudents.length,

      students:
        atRiskStudents

    };


  } catch (error) {

    console.error(
      'At-Risk Students Error:',
      error
    );


    return {

      success: false,

      message:
        error.message

    };

  }

}


function getStudentAttendanceByUSN(attendanceRecords, usn) {

  const targetUSN =
    String(usn || '')
      .trim()
      .toUpperCase();

  if (!targetUSN) {
    return null;
  }

  const record =
    attendanceRecords.find(function(item) {

      const recordUSN =
        String(
          item.USN ||
          item.usn ||
          item.Student_ID ||
          item.Student_Id ||
          ''
        )
          .trim()
          .toUpperCase();

      return recordUSN === targetUSN;

    });

  if (!record) {
    return null;
  }


  const possibleFields = [
    'Overall_Attendance',
    'Overall Attendance',
    'OverallAttendance',
    'Attendance_Percentage',
    'Attendance Percentage',
    'Attendance',
    'Attendance_Percentage_%',
    'Percentage'
  ];


  for (let i = 0; i < possibleFields.length; i++) {

    const key =
      possibleFields[i];

    if (
      record[key] !== undefined &&
      record[key] !== null &&
      String(record[key]).trim() !== ''
    ) {

      let value =
        Number(
          String(record[key])
            .replace('%', '')
            .trim()
        );

      if (Number.isFinite(value)) {

        // Handles values such as 0.72
        if (value > 0 && value <= 1) {
          value = value * 100;
        }

        return value;
      }
    }
  }


  return null;
}


