/**
 * @OnlyCurrentDoc
 */

// --- doGet & Global Constants ---

/**
 * Serves the HTML file of the web app.
 * @returns {HtmlOutput} The HTML output for the web app.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('ระบบจัดการคะแนนความประพฤติ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

const TEACHERS_SHEET = 'Teachers';
const BEHAVIORS_SHEET = 'Behaviors';
const INFRACTIONS_SHEET = 'Infractions';
const STUDENTS_SHEET = 'Students';
const CLASSES_SHEET = 'Classes';

const HEADERS = {
  [TEACHERS_SHEET]: ['Name', 'Username', 'Password'],
  [BEHAVIORS_SHEET]: ['ID', 'Name', 'Score', 'Type'], // Type: 'positive' or 'negative'
  [INFRACTIONS_SHEET]: ['ID', 'StudentUUID', 'StudentName', 'StudentClass', 'Date', 'BehaviorID', 'Comment', 'Timestamp'],
  [STUDENTS_SHEET]: ['ID', 'StudentCode', 'Name', 'Class', 'InitialScore', 'DeductedScore', 'AddedScore'], 
  [CLASSES_SHEET]: ['ID', 'Name']
};

// --- UTILITY FUNCTIONS ---

/**
 * Gets a sheet by name, creating it with headers if it doesn't exist.
 * @param {string} sheetName The name of the sheet.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet object.
 */
function getSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (HEADERS[sheetName]) {
      sheet.getRange(1, 1, 1, HEADERS[sheetName].length).setValues([HEADERS[sheetName]]).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

/**
 * Generates a unique UUID.
 * @returns {string} A UUID string.
 */
function generateId() { return Utilities.getUuid(); }

/**
 * Logs errors to the Apps Script logger for debugging.
 * @param {string} functionName The name of the function where the error occurred.
 * @param {Error} error The error object.
 */
function logError(functionName, error) { Logger.log(`Error in ${functionName}: ${error.message} \n ${error.stack}`); }

// --- AUTHENTICATION ---

/**
 * Authenticates a teacher based on username and password.
 * @param {string} username The teacher's username.
 * @param {string} password The teacher's password.
 * @returns {{success: boolean, message: string|undefined}} Result object.
 */
function teacherLogin(username, password) {
  try {
    const data = getSheet(TEACHERS_SHEET).getDataRange().getValues().slice(1);
    const user = data.find(row => row[1] === username && row[2] === password);
    return user ? { success: true } : { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
  } catch (e) {
    logError('teacherLogin', e);
    return { success: false, message: `เกิดข้อผิดพลาดในการเชื่อมต่อ: ${e.message}` };
  }
}

/**
 * Retrieves a student's data by their student code.
 * @param {string} studentCode The student's public code.
 * @returns {{success: boolean, student: object|undefined, message: string|undefined}} Result object.
 */
function getStudentByCode(studentCode) {
  try {
    if (!studentCode || studentCode.trim() === '') return { success: false, message: 'กรุณาระบุรหัสนักเรียน' };
    const studentRow = getSheet(STUDENTS_SHEET).getDataRange().getValues().slice(1).find(row => row[1] && row[1].toString().trim() === studentCode.trim());
    if (studentRow) {
      return {  
        success: true,  
        student: {  
          id: studentRow[0], studentCode: studentRow[1], name: studentRow[2], class: studentRow[3],  
          initialScore: parseInt(studentRow[4], 10) || 0,  
          deductedScore: parseInt(studentRow[5], 10) || 0,
          addedScore: parseInt(studentRow[6], 10) || 0
        }  
      };
    }
    return { success: false, message: 'ไม่พบรหัสนักเรียนนี้ในระบบ' };
  } catch (e) {
    logError('getStudentByCode', e);
    return { success: false, message: `เกิดข้อผิดพลาด: ${e.message}` };
  }
}

// --- DASHBOARD DATA ---

/**
 * Gathers and computes all data needed for the main dashboard.
 * @returns {object} An object containing dashboard statistics.
 */
function getDashboardData() {
  try {
    const infractions = getSheet(INFRACTIONS_SHEET).getDataRange().getValues().slice(1);
    const students = getSheet(STUDENTS_SHEET).getDataRange().getValues().slice(1);
    const behaviors = getSheet(BEHAVIORS_SHEET).getDataRange().getValues().slice(1);
    const behaviorMap = new Map(behaviors.map(b => [b[0], {name: b[1], score: b[2], type: b[3]}]));

    // 1. Top Behavior
    const behaviorCounts = infractions.reduce((acc, inf) => {
      const behaviorId = inf[5];
      acc[behaviorId] = (acc[behaviorId] || 0) + 1;
      return acc;
    }, {});
    let topBehavior = { name: 'N/A', count: 0 };
    if (Object.keys(behaviorCounts).length > 0) {
      const topId = Object.keys(behaviorCounts).reduce((a, b) => behaviorCounts[a] > behaviorCounts[b] ? a : b);
      if (behaviorMap.has(topId)) {
        topBehavior = { name: behaviorMap.get(topId).name, count: behaviorCounts[topId] };
      }
    }

    // 2. Lowest Scoring Student
    let lowestStudent = { name: 'N/A', initialScore: 0, deductedScore: 0, addedScore: 0, score: 0 };
    let minScore = Infinity;
    if (students.length > 0) {
      students.forEach(s => {
        const score = (parseInt(s[4]) || 0) - (parseInt(s[5]) || 0) + (parseInt(s[6]) || 0);
        if (score < minScore) {
          minScore = score;
          lowestStudent = { name: s[2], initialScore: s[4], deductedScore: s[5], addedScore: s[6], score: score };
        }
      });
    }

    // 3. Reports this month
    const now = new Date();
    const totalReportsInMonth = infractions.filter(inf => {
      const d = new Date(inf[4]);
      return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
    }).length;

    // 4. Chart data
    const behaviorStats = { labels: [], datasets: [{ label: 'จำนวนครั้ง', data: [], backgroundColor: [] }] };
    for (const [id, count] of Object.entries(behaviorCounts)) {
      if (behaviorMap.has(id)) {
        behaviorStats.labels.push(behaviorMap.get(id).name);
        behaviorStats.datasets[0].data.push(count);
        behaviorStats.datasets[0].backgroundColor.push(behaviorMap.get(id).type === 'positive' ? '#2ECC71' : '#E74C3C');
      }
    }
    
    const classScores = students.reduce((acc, s) => {
      const className = s[3];
      if (!acc[className]) acc[className] = { totalScore: 0, count: 0 };
      acc[className].totalScore += (parseInt(s[4]) || 0) - (parseInt(s[5]) || 0) + (parseInt(s[6]) || 0);
      acc[className].count++;
      return acc;
    }, {});

    const classChartColors = ['#3498DB', '#F1C40F', '#8E44AD', '#1ABC9C', '#E67E22', '#2980B9', '#F39C12'];
    const classStats = { labels: [], datasets: [{ data: [], backgroundColor: [] }] };
    let colorIndex = 0;
    for(const [name, stats] of Object.entries(classScores)) {
      classStats.labels.push(name);
      const averageScore = stats.count > 0 ? stats.totalScore / stats.count : 0;
      classStats.datasets[0].data.push(averageScore);
      classStats.datasets[0].backgroundColor.push(classChartColors[colorIndex % classChartColors.length]);
      colorIndex++;
    }
    
    return { success: true, topBehavior, lowestStudent, totalReportsInMonth, behaviorStats, classStats };
  } catch(e) {
    logError('getDashboardData', e);
    return { success: false, message: e.message };
  }
}

// --- BEHAVIORS CRUD ---

/**
 * Retrieves all behaviors.
 * @returns {Array<object>} An array of behavior objects.
 */
function getBehaviors() {
  try {
    return getSheet(BEHAVIORS_SHEET).getDataRange().getValues().slice(1)
      .map(row => ({ id: row[0], name: row[1], score: parseInt(row[2], 10) || 0, type: row[3] || 'negative' }))
      .filter(b => b.id && b.name);
  } catch (e) { logError('getBehaviors', e); return []; }
}

/**
 * Saves a behavior (creates or updates).
 * @param {object} behavior The behavior object to save.
 * @returns {{success: boolean, message: string|undefined}} Result object.
 */
function saveBehavior(behavior) {
  try {
    const score = parseInt(behavior.score, 10);
    if (isNaN(score)) return { success: false, message: 'กรุณากรอกคะแนนเป็นตัวเลข' };

    if (behavior.type === 'positive' && (score < 0 || score > 100)) {
      return { success: false, message: 'คะแนนสำหรับพฤติกรรมเชิงบวกต้องอยู่ในช่วง 0 ถึง 100' };
    }
    if (behavior.type === 'negative' && (score > 0 || score < -100)) {
      return { success: false, message: 'คะแนนสำหรับพฤติกรรมเชิงลบต้องอยู่ในช่วง -100 ถึง 0' };
    }

    const sheet = getSheet(BEHAVIORS_SHEET);
    const data = sheet.getDataRange().getValues();
    if (data.slice(1).some(row => row[1] && row[1].trim().toLowerCase() === behavior.name.trim().toLowerCase() && row[0] !== behavior.id)) {
      return { success: false, message: `ชื่อพฤติกรรม "${behavior.name}" มีอยู่แล้ว` };
    }
    if (behavior.id) {
      const rowIndex = data.findIndex(row => row[0] === behavior.id);
      if (rowIndex !== -1) {
        sheet.getRange(rowIndex + 1, 1, 1, 4).setValues([[behavior.id, behavior.name.trim(), score, behavior.type]]);
      }
    } else {
      sheet.appendRow([generateId(), behavior.name.trim(), score, behavior.type]);
    }
    return { success: true };
  } catch (e) { logError('saveBehavior', e); return { success: false, message: e.message }; }
}

/**
 * Retrieves a single behavior by its ID.
 * @param {string} id The UUID of the behavior.
 * @returns {object|null} The behavior object or null if not found.
 */
function getBehaviorById(id) {
  try {
    const row = getSheet(BEHAVIORS_SHEET).getDataRange().getValues().find(r => r[0] === id);
    return row ? { id: row[0], name: row[1], score: parseInt(row[2], 10), type: row[3] } : null;
  } catch (e) {
    logError('getBehaviorById', e);
    return null;
  }
}

/**
 * Deletes a behavior if it is not in use.
 * @param {string} id The UUID of the behavior to delete.
 * @returns {{success: boolean, message: string|undefined}} Result object.
 */
function deleteBehavior(id) {
  try {
    const infractions = getSheet(INFRACTIONS_SHEET).getDataRange().getValues().slice(1);
    if (infractions.some(inf => inf[5] === id)) {
      return { success: false, message: 'ไม่สามารถลบพฤติกรรมนี้ได้ เนื่องจากมีการใช้งานแล้ว' };
    }
    const sheet = getSheet(BEHAVIORS_SHEET);
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => row[0] === id);
    if (rowIndex > 0) {
      sheet.deleteRow(rowIndex + 1);
      return { success: true };
    }
    return { success: false, message: 'ไม่พบพฤติกรรมที่ต้องการลบ' };
  } catch (e) { logError('deleteBehavior', e); return { success: false, message: e.message }; }
}

// --- CLASSES CRUD ---

/**
 * Retrieves all classes with their student counts.
 * @returns {Array<object>} An array of class objects.
 */
function getClassesWithStudentCount() {
  try {
    const classSheet = getSheet(CLASSES_SHEET);
    if (classSheet.getLastRow() < 2) return [];
    const classes = classSheet.getDataRange().getValues().slice(1).map(row => ({ id: row[0], name: row[1], studentCount: 0 }));

    const studentSheet = getSheet(STUDENTS_SHEET);
    if (studentSheet.getLastRow() < 2) return classes;
    const students = studentSheet.getDataRange().getValues().slice(1);
    
    const classMap = new Map(classes.map(c => [c.name, c]));
    students.forEach(s => {
      const className = s[3];
      if (classMap.has(className)) {
        classMap.get(className).studentCount++;
      }
    });
    return Array.from(classMap.values());
  } catch(e) { logError('getClassesWithStudentCount', e); return []; }
}

/**
 * Retrieves all classes.
 * @returns {Array<object>} An array of class objects.
 */
function getClasses() {
  try {
    const sheet = getSheet(CLASSES_SHEET);
    if(sheet.getLastRow() < 2) return [];
    return sheet.getDataRange().getValues().slice(1)
      .map(row => ({ id: row[0], name: row[1] }))
      .filter(c => c.id && c.name);
  } catch (e) { logError('getClasses', e); return []; }
}

/**
 * Saves a class (creates or updates).
 * @param {object} classData The class object to save.
 * @returns {{success: boolean, message: string|undefined}} Result object.
 */
function saveClass(classData) {
  try {
    if (!classData.name || classData.name.trim() === '') {
      return { success: false, message: 'กรุณากรอกชื่อชั้นเรียน' };
    }
    const sheet = getSheet(CLASSES_SHEET);
    const data = sheet.getDataRange().getValues();
    if (data.slice(1).some(row => row[1] && row[1].trim().toLowerCase() === classData.name.trim().toLowerCase() && row[0] !== classData.id)) {
      return { success: false, message: `ชื่อชั้นเรียน "${classData.name}" มีอยู่แล้ว` };
    }
    if (classData.id) {
      const rowIndex = data.findIndex(row => row[0] === classData.id);
      if (rowIndex !== -1) {
        sheet.getRange(rowIndex + 1, 2).setValue(classData.name.trim());
      }
    } else {
      sheet.appendRow([generateId(), classData.name.trim()]);
    }
    return { success: true };
  } catch (e) { logError('saveClass', e); return { success: false, message: e.message }; }
}


/**
 * Deletes a class if it has no students.
 * @param {string} id The UUID of the class to delete.
 * @returns {{success: boolean, message: string|undefined}} Result object.
 */
function deleteClass(id) {
  try {
    const studentSheet = getSheet(STUDENTS_SHEET);
    const classSheet = getSheet(CLASSES_SHEET);

    const classDataRow = classSheet.getDataRange().getValues().find(r => r[0] === id);
    if (!classDataRow) return { success: false, message: 'ไม่พบชั้นเรียนที่ต้องการลบ' };
    const className = classDataRow[1];

    if (studentSheet.getLastRow() > 1) {
        const students = studentSheet.getDataRange().getValues().slice(1);
        if (students.some(s => s[3] === className)) {
          return { success: false, message: 'ไม่สามารถลบชั้นเรียนที่มีนักเรียนอยู่ได้' };
        }
    }

    const rowIndex = classSheet.getDataRange().getValues().findIndex(row => row[0] === id);
    if (rowIndex > 0) { // Check if it's not the header row
      classSheet.deleteRow(rowIndex + 1);
      return { success: true };
    }
    return { success: false, message: 'ไม่พบชั้นเรียนที่ต้องการลบ' };
  } catch (e) { logError('deleteClass', e); return { success: false, message: e.message }; }
}

// --- STUDENTS CRUD ---

/**
 * Retrieves data needed for the student management page.
 * @returns {object} An object containing students and classes.
 */
function getStudentsAndClasses() {
  try {
    const studentSheet = getSheet(STUDENTS_SHEET);
    let students = [];
    if (studentSheet.getLastRow() > 1) {
        students = studentSheet.getDataRange().getValues().slice(1)
          .map(row => ({ id: row[0], studentCode: row[1], name: row[2], class: row[3] }))
          .filter(s => s.id && s.name);
    }
    const classes = getClasses();
    return { students, classes };
  } catch (e) { logError('getStudentsAndClasses', e); return { students: [], classes: [] }; }
}

/**
 * Saves a student (creates or updates).
 * @param {object} student The student object to save.
 * @returns {{success: boolean, message: string|undefined}} Result object.
 */
function saveStudent(student) {
  try {
    const sheet = getSheet(STUDENTS_SHEET);
    const data = sheet.getDataRange().getValues();

    if (data.slice(1).some(row => row[1] && String(row[1]).trim().toLowerCase() === student.studentCode.trim().toLowerCase() && row[0] !== student.id)) {
      return { success: false, message: `รหัสนักเรียน "${student.studentCode}" มีอยู่แล้ว` };
    }

    const validClasses = new Set(getClasses().map(c => c.name));
    if (!validClasses.has(student.class)) {
      return { success: false, message: `ชั้นเรียน "${student.class}" ไม่มีในระบบ` };
    }
    if (student.id) {
      const rowIndex = data.findIndex(row => row[0] === student.id);
      if (rowIndex !== -1) {
        // Only update code, name, and class. Scores remain.
        sheet.getRange(rowIndex + 1, 2, 1, 3).setValues([[
          student.studentCode.trim(),
          student.name.trim(),
          student.class
        ]]);
      }
    } else {
      // New student starts with 0 scores
      sheet.appendRow([generateId(), student.studentCode.trim(), student.name.trim(), student.class, 0, 0, 0]);
    }
    return { success: true };
  } catch (e) { logError('saveStudent', e); return { success: false, message: e.message }; }
}


/**
 * Deletes a student if they have no infractions.
 * @param {string} id The UUID of the student to delete.
 * @returns {{success: boolean, message: string|undefined}} Result object.
 */
function deleteStudent(id) {
  try {
    const infractions = getSheet(INFRACTIONS_SHEET).getDataRange().getValues().slice(1);
    if (infractions.some(inf => inf[1] === id)) {
      return { success: false, message: 'ไม่สามารถลบนักเรียนที่มีประวัติการกระทำผิดได้' };
    }
    const sheet = getSheet(STUDENTS_SHEET);
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => row[0] === id);
    if (rowIndex > 0) {
      sheet.deleteRow(rowIndex + 1);
      return { success: true };
    }
    return { success: false, message: 'ไม่พบนักเรียนที่ต้องการลบ' };
  } catch (e) { logError('deleteStudent', e); return { success: false, message: e.message }; }
}

// --- SCORE CALCULATION ---

/**
 * Recalculates and updates total deducted/added scores for a student from their infractions history.
 * @param {string} studentUUID The UUID of the student.
 * @returns {{success: boolean, message: string|undefined}} Result object.
 */
function updateStudentScoresFromInfractions(studentUUID) {
  try {
    const infractionsSheet = getSheet(INFRACTIONS_SHEET);
    const behaviorsSheet = getSheet(BEHAVIORS_SHEET);
    const studentsSheet = getSheet(STUDENTS_SHEET);

    const allInfractions = infractionsSheet.getLastRow() > 1 ? infractionsSheet.getDataRange().getValues().slice(1) : [];
    const behaviorMap = behaviorsSheet.getLastRow() > 1 ? new Map(behaviorsSheet.getDataRange().getValues().slice(1).map(b => [b[0], { score: parseInt(b[2]) || 0, type: b[3] }])) : new Map();
    
    const studentInfractions = allInfractions.filter(inf => inf[1] === studentUUID);
    let totalDeducted = 0, totalAdded = 0;

    studentInfractions.forEach(infraction => {
      const behavior = behaviorMap.get(infraction[5]);
      if (behavior) {
        if (behavior.type === 'positive') totalAdded += behavior.score;
        else totalDeducted += Math.abs(behavior.score);
      }
    });
    
    const studentData = studentsSheet.getDataRange().getValues();
    const studentRowIndex = studentData.findIndex(row => row[0] === studentUUID);
    if (studentRowIndex > 0) {
      // Update deducted and added scores in columns F (6) and G (7)
      studentsSheet.getRange(studentRowIndex + 1, 6).setValue(totalDeducted);
      studentsSheet.getRange(studentRowIndex + 1, 7).setValue(totalAdded);
    }
    return { success: true };
  } catch(e) { 
    logError('updateStudentScoresFromInfractions', e); 
    return { success: false, message: e.message }; 
  }
}

// --- INFRACTIONS ---

/**
 * Retrieves data needed to populate the infraction logging form.
 * @returns {object} An object containing students and behaviors data.
 */
function getInfractionFormData() {
  const students = getStudentReportData().students;
  const behaviors = getBehaviors();
  return { students, behaviors };
}

/**
 * Saves an infraction and updates the student's score.
 * @param {object} infraction The infraction object from the client.
 * @returns {{success: boolean, message: string|undefined}} Result object.
 */
function saveInfraction(infraction) {
  try {
    const sheet = getSheet(INFRACTIONS_SHEET);
    const formattedDate = Utilities.formatDate(new Date(infraction.date), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd");
    
    const studentSheet = getSheet(STUDENTS_SHEET);
    const studentDataRange = studentSheet.getDataRange().getValues();
    const studentRow = studentDataRange.find(row => row[0] === infraction.studentUUID);
    if (!studentRow) {
      return { success: false, message: 'ไม่พบข้อมูลนักเรียน' };
    }
    
    const studentName = studentRow[2]; 
    const studentClass = studentRow[3];
    
    sheet.appendRow([generateId(), infraction.studentUUID, studentName, studentClass, formattedDate, infraction.behaviorId, infraction.comment, new Date()]);
    
    const updateResult = updateStudentScoresFromInfractions(infraction.studentUUID);
    if (!updateResult.success) {
      // Even if score update fails, the infraction is logged. Return the error message.
      return { success: true, message: `บันทึกข้อมูลสำเร็จ แต่เกิดข้อผิดพลาดในการอัปเดตคะแนน: ${updateResult.message}` };
    }
    
    return { success: true };
  } catch (e) { 
    logError('saveInfraction', e); 
    return { success: false, message: e.message }; 
  }
}


/**
 * Retrieves all infractions for a specific student.
 * @param {string} studentUUID The UUID of the student.
 * @returns {Array<object>} An array of the student's infraction objects.
 */
function getStudentInfractions(studentUUID) {
  try {
    const infractionsRaw = getSheet(INFRACTIONS_SHEET).getDataRange().getValues().slice(1);
    const behaviorsRaw = getSheet(BEHAVIORS_SHEET).getDataRange().getValues().slice(1);
    const behaviorMap = new Map(behaviorsRaw.map(b => [b[0], {name: b[1], score: b[2], type: b[3]}]));
    return infractionsRaw.filter(row => row[1] === studentUUID).map(row => {
        const behavior = behaviorMap.get(row[5]);
        return {
          id: row[0], date: new Date(row[4]).toISOString(),
          behaviorName: behavior ? behavior.name : 'ไม่พบข้อมูล',
          behaviorScore: behavior ? (parseInt(behavior.score, 10) || 0) : 0,
          behaviorType: behavior ? (behavior.type || 'negative') : 'negative',
          comment: row[6]
        };
      }).sort((a,b) => new Date(b.date) - new Date(a.date));
  } catch(e) { logError('getStudentInfractions', e); return []; }
}

// --- REPORTS & FILE HANDLING ---

/**
 * Retrieves all data needed for the main student report.
 * @returns {object} An object containing students and classes data.
 */
function getStudentReportData() {
  try {
    const studentSheet = getSheet(STUDENTS_SHEET);
    const students = studentSheet.getLastRow() < 2 ? [] : studentSheet.getDataRange().getValues().slice(1).map(row => ({
      id: row[0], studentCode: row[1], name: row[2], class: row[3],
      initialScore: parseInt(row[4], 10) || 0,
      deductedScore: parseInt(row[5], 10) || 0,
      addedScore: parseInt(row[6], 10) || 0
    })).filter(s => s.id && s.name);
    return { students, classes: getClasses() };
  } catch (e) { logError('getStudentReportData', e); return { students:[], classes:[] }; }
}

function filterStudentReport(search, classFilter) {
  let { students } = getStudentReportData();
  if (search) { 
    const term = search.toLowerCase(); 
    students = students.filter(s => {
      const studentCode = String(s.studentCode || '').toLowerCase();
      const studentName = String(s.name || '').toLowerCase();
      return studentCode.includes(term) || studentName.includes(term);
    }); 
  }
  if (classFilter && classFilter !== 'all') {
    students = students.filter(s => s.class === classFilter);
  }
  return students;
}

/**
 * Creates a student report file (PDF or CSV) and returns a download URL.
 * @param {string} type The type of file to export ('pdf' or 'csv').
 * @returns {string|null} The download URL for the generated file or null on failure.
 */
function exportReport(type) {
  try {
    const { students } = getStudentReportData();
    const timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd_HHmm");
    let fileBlob, fileName;

    if (type === 'csv') {
      let csv = '\uFEFF' + 'ลำดับ,รหัสนักเรียน,ชื่อ-สกุล,ชั้น,คะแนนเริ่มต้น,คะแนนที่เพิ่ม,คะแนนที่หัก,คะแนนคงเหลือ\n';
      students.forEach((s, i) => {
        const score = s.initialScore - s.deductedScore + s.addedScore;
        csv += `${i+1},"${s.studentCode}","${s.name}","${s.class}",${s.initialScore},${s.addedScore},${s.deductedScore},${score}\n`;
      });
      fileName = `student_report_${timestamp}.csv`;
      fileBlob = Utilities.newBlob(csv, MimeType.CSV, fileName);
    } else { // pdf
      let html = "<html><head><style>@import url('https://fonts.googleapis.com/css2?family=Sarabun&display=swap'); body{font-family: 'Sarabun', sans-serif;} table,th,td{border:1px solid black; border-collapse:collapse; padding:4px;}</style></head><body>"
      + `<h1>รายงานคะแนนความประพฤติ (${new Date().toLocaleDateString('th-TH')})</h1>`
      + "<table><tr><th>#</th><th>รหัส</th><th>ชื่อ-สกุล</th><th>ชั้น</th><th>คะแนนคงเหลือ</th></tr>";
      students.forEach((s, i) => {
        const score = s.initialScore - s.deductedScore + s.addedScore;
        html += `<tr><td>${i+1}</td><td>${s.studentCode}</td><td>${s.name}</td><td>${s.class}</td><td>${score}</td></tr>`;
      });
      html += "</table></body></html>";
      fileName = `student_report_${timestamp}.pdf`;
      fileBlob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF).setName(fileName);
    }
    
    const file = DriveApp.createFile(fileBlob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getDownloadUrl().replace('&e=download&gd=true', '');
  } catch (e) {
    logError('exportReport', e);
    return null;
  }
}

/**
 * Creates a CSV template file for importing students.
 * @returns {string|null} The download URL for the template file.
 */
function downloadStudentTemplate() {
  try {
    const csv = '\uFEFF' + 'รหัสนักเรียน,ชื่อ-นามสกุล,ชั้นเรียน\n' + 'ตัวอย่าง123,เด็กชายตัวอย่าง ตั้งใจเรียน,ม.1/1\n';
    const fileName = 'student_import_template.csv';
    const file = DriveApp.createFile(Utilities.newBlob(csv, MimeType.CSV, fileName));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getDownloadUrl().replace('&e=download&gd=true', '');
  } catch (e) { logError('downloadStudentTemplate', e); return null; }
}

/**
 * Imports students from a CSV file content.
 * @param {string} csvContent The string content of the uploaded CSV file.
 * @returns {{success: boolean, message: string}} Result object with a summary message.
 */
function importStudentsFromCSV(csvContent) {
  try {
    const studentSheet = getSheet(STUDENTS_SHEET);
    const existingStudents = studentSheet.getLastRow() < 2 ? [] : studentSheet.getDataRange().getValues().slice(1);
    const existingCodes = new Set(existingStudents.map(s => String(s[1]).trim()));
    const validClasses = new Set(getClasses().map(c => c.name));

    const rows = Utilities.parseCsv(csvContent);
    if (rows.length < 2) return { success: false, message: 'ไฟล์ CSV ไม่มีข้อมูลนักเรียน' };

    const newStudentsData = [];
    let errors = [];
    
    rows.slice(1).forEach((cols, i) => {
      if (cols.length < 3) {
        errors.push(`แถว ${i+2}: ข้อมูลไม่ครบถ้วน`); return;
      }
      const [code, name, className] = cols.map(c => c.trim());
      if (!code || !name || !className) { errors.push(`แถว ${i+2}: ข้อมูลบางช่องว่างเปล่า`); return; }
      if (existingCodes.has(code)) { errors.push(`แถว ${i+2}: รหัสนักเรียน ${code} มีอยู่แล้ว`); return; }
      if (!validClasses.has(className)) { errors.push(`แถว ${i+2}: ชั้นเรียน ${className} ไม่มีในระบบ`); return; }
      
      newStudentsData.push([generateId(), code, name, className, 0, 0, 0]);
      existingCodes.add(code); // Add to set to check for duplicates within the file itself
    });

    if (newStudentsData.length > 0) {
      studentSheet.getRange(studentSheet.getLastRow() + 1, 1, newStudentsData.length, newStudentsData[0].length).setValues(newStudentsData);
    }

    let message = `นำเข้าสำเร็จ ${newStudentsData.length} รายการ`;
    if (errors.length > 0) {
      message += `\nเกิดข้อผิดพลาด ${errors.length} รายการ:\n- ${errors.slice(0, 10).join('\n- ')}`; // Show first 10 errors
       if (errors.length > 10) message += `\n...และอีก ${errors.length - 10} รายการ`;
    }
    return { success: (errors.length === 0), message: message };

  } catch (e) {
    logError('importStudentsFromCSV', e);
    return { success: false, message: `เกิดข้อผิดพลาดในการประมวลผลไฟล์: ${e.message}` };
  }
}
