# VBExportHours
This is a project created for the Vandegrift ViperBots FIRST (For Inspiration and Recognition of Science and Technology) program.  Vandegtift High School is part of the Leander Independent School District in Texas.  

The ViperBots use a fingerprint scanner for tracking the time that their member students spend in the robot room and working on the robot.  As part of the University Interscholastic League, the students have a maximum amount of time that they may spend with extra-curricular programs, and the scanner helps the ViperBots program sponsors monitor that activity.  It also helps to let the students and the parents know what their hours are on a weekly basis so they can self-regulate by limiting their hours in the future, or increase their hours and, therby increasing their contribution to their team's activity and ultimate success on the competition field.

The specific fingerprint scanner in use has the ability to export the hours from a specific time period as an Excel-97 formatted file.  The output of that scanner has multiple audiences:
* The program sponsors to manage the time according with any maximum or minimums for success
* The parents and the students themselves, so they can see if they are nearing the maximum or are just not spending enough time to help out their team.

While the Excel formatted file saves a lot of information, there's really only a bit of information that is needed to provide to the coaches and the families, and manually culling the extraneous data and organizing the data became a very time-consuming process.  This program helps streamline that process by consuming the output of the fingerprint scanner and generating a new Excel-97 or Excel-2010+ style output that just has the information needed as two worksheets in the workbook.  One of the worksheets is intended for the coaches and includes the student's personal information (e.g., name), and the other is intended for the parents and only has the student's unique ID.

This program uses the Apache CLI and POI packages to provide the options for command line arguments and the ability to read/write Excel-formatted files.
