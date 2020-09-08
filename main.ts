/// <reference path="./classroom_v1.d.ts" />
type Course = GoogleAppsScript.Classroom.Schema.Course;
type CourseWork = GoogleAppsScript.Classroom.Schema.CourseWork;
type Student = GoogleAppsScript.Classroom.Schema.Student;
type Material = GoogleAppsScript.Classroom.Schema.Material;
type Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;

const SECOND = 1000;
const MINUTE = SECOND * 60;
const HOUR = MINUTE * 60;
const DAY = HOUR * 24;

const scriptProperties = PropertiesService.getScriptProperties();
const MMSBAND_COURSEID = scriptProperties.getProperty("MMSBAND_COURSEID");
const APPSSCRIPTTEST_COURSEID = scriptProperties.getProperty("APPSSCRIPTTEST_COURSEID");
const PRACTICECHART_TOPICID = scriptProperties.getProperty("PRACTICECHART_TOPICID");
const TIMEZONE = -4;
const COURSEADMIN = scriptProperties.getProperty("COURSEADMIN");

const numToShortMonth = (num: number): string => {
    switch (num) {
        case 0:
            return null;
        case 1:
            return "Jan";
        case 2:
            return "Feb";
        case 3:
            return "Mar";
        case 4:
            return "Apr";
        case 5:
            return "May";
        case 6:
            return "Jun";
        case 7:
            return "Jul";
        case 8:
            return "Aug";
        case 9:
            return "Sep";
        case 10:
            return "Oct";
        case 11:
            return "Nov";
        case 12:
            return "Dec";
    }
};

const numToWeekday = (num: number): string => {
    switch (num) {
        case 0:
            return "Sunday";
        case 1:
            return "Monday";
        case 2:
            return "Tuesday";
        case 3:
            return "Wednesday";
        case 4:
            return "Thursday";
        case 5:
            return "Friday";
        case 6:
            return "Saturday";
    }
};

const dateToISO = (date: Date): string => {
    const year = date.getFullYear();
    const month = date.getMonth();
    const day = date.getDate();
    return `${year}-${month}-${day}`;
};

const getCourseById = (id: string): Course => {
    return Classroom.Courses.get(id);
};

const getCourseByName = (name: string): Course => {
    const { courses } = Classroom.Courses.list();
    for (let i = 0; i < courses.length; i++) {
        const course = courses[i];
        if (course.name == name) return course;
    }
};

const getStudentsFromCourse = (id: string): Student[] => {
    return Classroom.Courses.Students.list(id).students;
};

const createAssignment = (
    title: string,
    description: string,
    dueDate: string,
    dueTime: string,
    attachments: [string, string][] //[id, shareMode] `id`: Drive File ID, `shareMode`: ShareMode {"VIEW", "EDIT", "STUDENT_COPY"}
): CourseWork => {
    const [year, month, day] = dueDate.split("-").map((str) => Number(str));
    const dueDateObject: GoogleAppsScript.Classroom.Schema.Date = { year, month, day };
    const [hours, minutes] = dueTime.split(":").map((str) => Number(str));
    const dueTimeObject: GoogleAppsScript.Classroom.Schema.TimeOfDay = { hours: hours - TIMEZONE, minutes };
    const materials: Material[] = attachments.map((attachment) => {
        const [id, shareMode] = attachment;
        return {
            driveFile: {
                driveFile: {
                    id,
                },
                shareMode,
            },
        };
    });
    return {
        title,
        description,
        dueDate: dueDateObject,
        dueTime: dueTimeObject,
        workType: "ASSIGNMENT",
        materials,
    } as CourseWork;
};

const postCourseWork = (courseWork: CourseWork, courseId: string, topicId: string = null): CourseWork => {
    courseWork.state = "PUBLISHED";
    courseWork.topicId = topicId;
    return Classroom.Courses.CourseWork.create(courseWork, courseId);
};

const createPracticeChart = (name: string, days: Date[]): Spreadsheet => {
    const spreadsheet = SpreadsheetApp.create(name);
    const sheet = spreadsheet.getSheets()[0];
    sheet.setName("Practice Chart");
    sheet.setHiddenGridlines(true);

    const styleBigAndBold = SpreadsheetApp.newTextStyle()
        .setBold(true)
        .setFontSize(18)
        .setFontFamily("Century Gothic")
        .build();
    const styleSmallAndBold = SpreadsheetApp.newTextStyle()
        .setBold(true)
        .setFontSize(12)
        .setFontFamily("Proxima Nova")
        .build();

    // TITLE
    sheet.setRowHeight(1, 40);
    const titleCell = spreadsheet.getRange("A1");
    titleCell.setValue("Practice Chart");
    titleCell.setTextStyle(styleBigAndBold);
    titleCell.setBackground("lightcyan");
    titleCell.setHorizontalAlignment("center");
    titleCell.setVerticalAlignment("middle");
    sheet.getRange("A1:G1").mergeAcross();

    // WEEKDAYS AND DATES
    sheet.setRowHeight(2, 20);
    sheet.setRowHeight(3, 20);
    const weekdays = sheet.getRange("A2:G3");
    weekdays.setValues([
        days.map((day) => `${numToWeekday(day.getDay())}`),
        days.map((day) => `${day.getMonth()}/${day.getDate()}`),
    ]);
    weekdays.setTextStyle(styleSmallAndBold);
    weekdays.setHorizontalAlignment("center");

    // MINUTES
    spreadsheet.setRowHeight(4, 75);
    for (let i = 1; i <= 7; i++) {
        spreadsheet.setColumnWidth(i, 90);
    }
    const minutes = spreadsheet.getRange("A4:G4");
    minutes.setTextStyle(styleBigAndBold);
    minutes.setHorizontalAlignment("center");
    minutes.setVerticalAlignment("middle");
    const total = sheet.getRange("H4");
    total.setFormula("=SUM(A4:G4)");
    total.setTextStyle(styleBigAndBold);

    //BLANK
    // ASSIGNMENT TITLE
    // ASSIGNMENT LIST
    sheet.deleteRows(8, sheet.getMaxRows() - 8);
    sheet.deleteColumns(10, 26 - 10);
    sheet.protect().addEditor(COURSEADMIN).setUnprotectedRanges([minutes]);
    return spreadsheet;
};

const main = () => {
    const startTime = new Date(2020, 9, 7);
    const days: Date[] = [];
    for (let i = 0; i < 7; i++) {
        days.push(new Date(startTime.valueOf() + DAY * i));
    }
    const dueDate = new Date(startTime.valueOf() + DAY * 7);
    const title = `Practice Chart (Due Monday, ${numToShortMonth(dueDate.getMonth())} ${dueDate.getDate()})`;
    const description = ``;
    const practiceChart = createPracticeChart(title, days);
    const courseWork = createAssignment(title, description, dateToISO(dueDate), "12:00", [
        [practiceChart.getId(), "STUDENT_COPY"],
    ]);
    //postCourseWork(courseWork, APPSSCRIPTTEST_COURSEID, PRACTICECHART_TOPICID);
    const index = SpreadsheetApp.create("Practice Chart Index");
    index.getRange("A1").setValue(practiceChart.getId());
    index.getRange("B1").setFormula(`=IMPORTRANGE($A1, "Practice Chart$H$4")`);
};
