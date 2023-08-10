# List of Students

## Description
The **List of Students** project is an Excel worksheet designed to streamline and automate the organization of students' activities using VBA (Visual Basic for Applications). This Excel tool offers two comprehensive plans to effectively manage student activities and class notes.

### First Plan: Activities
The Activities plan features a space to input student names, with each row containing checkboxes representing completed activities. To register a new activity, simply replace the text "Class activity" with the activity's name. Clicking the "Update" button verifies the filled fields – any name other than "Class activity" with a non-empty cell is counted as an activity. The total number of activities is displayed upon each "Update". Moreover, a calculation showcases the completed activities at the end of each row. For visual tracking, conditional formatting colorizes rows to illustrate student progress.

### Second Plan: Class Notes
The Class Notes plan is designed for methodical registration of classes. The "New entry" button launches a form, enabling easy data input. Upon clicking the "Save" button, cells are populated automatically. Entries can be removed by utilizing the "Delete entry" button, which prompts for the corresponding date.
<br><br>

## Execution
The project consists of two folders: "src" and "worksheet". The "worksheet" folder contains the Excel worksheet with embedded VBA code. Due to Microsoft's security measures, macros may be blocked when the file is initially opened. To resolve this, right-click the file, navigate to "Properties," and under the "Security" section, check the "Unblock" checkbox. Apply the change, then open the file and click "Enable content" in the warning message. The "src" folder contains all VBA code, including form-related code.

This project serves as a versatile tool to manage student activities and class notes, leveraging VBA automation for enhanced efficiency and organization.
