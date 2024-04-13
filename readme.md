# utm-attendance-process

utm-attendance-process is a python package for processing attendance data in 
UTM, and produce basic statistics of student's attendance

## Usage

### For Windows user

1. Download `process_attendance.exe` from the release page and reminder letter forms inside folder `forms` and place them inside a
processing folder (eg `Attendance`) to process the attendance data. The
attendance data are the PDF files downloaded from the UTM QR attendance
system.
1. Download the PDF files of the attendance record, placing them inside
a folder named `pdf` created inside the processing folder (eg
`Attendance/pdf`). Name each of the attendance record according 'YYMMDD-HH-D' format, eg 240301-14-2 corresponds to date 01/03/2024, time 14:00, and duration of 2 hours.
1. Double click the downloaded `process_attendance.exe` to open windows where additional settings can be specified (path to attendance record files, attendance exclusions, etc).
1. `attendance_processed-YYMMDD-HH-D.xlsx` will be generated, containing the
processed attendance information. Reminder letters will be automatically generated inside `reminder_letter-generated` folder.
