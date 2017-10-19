# VBExportHours
This is a project created for the [Vandegrift ViperBots](http://viperbots.org/) Robotics program ([FIRST](https://www.firstinspires.org/) For Inspiration and Recognition of Science and Technology).  [Vandegtift High School](http://vhs.leanderisd.org/) is part of the [Leander Independent School District](http://leanderisd.org/) in Texas.  The name is descriptive:  ViperBots Export Hours.

The ViperBots use a fingerprint scanner for tracking the time that their member students spend in the robot room and working on the robot.  As part of the [University Interscholastic League](http://www.uiltexas.org/), the students have a maximum amount of time that they may spend with extra-curricular programs, and the scanner helps the ViperBots program sponsors monitor that activity.  It also helps to let the students and the parents know what their hours are on a weekly basis so they can self-regulate by limiting their hours in the future, or increase their hours and, therby increasing their contribution to their team's activity and ultimate success on the competition field.

The fingerprint scanner in use has the ability to export the hours from a specific time period as an Excel-97 formatted file.  The output of that scanner has multiple audiences:
* The program sponsors to manage the time according with any maximum or minimums for success
* The parents and the students themselves, so they can see if they are nearing the maximum or are just not spending enough time to help out their team.

While the Excel formatted file saves a lot of information, there's really only a bit of information that is needed to provide to the coaches and the families.  When the ViperBots program was only a few teams, manually culling the extraneous data and organizing the data for sharing was not too cumbersome.  As the ViperBots grew to 9 teams and nearly 100 students, it quickly became a very time-consuming process.  

This program helps streamline that process by consuming the output of the fingerprint scanner and generating a new Excel-97 or Excel-2010+ style output that just has the information needed as two worksheets in the workbook.  One of the worksheets is intended for the coaches and includes the student's personal information (e.g., name), and the other is intended for the parents and only has the student's unique ID.  Each sheet is pre-formatted (and print ranges specified for Excel-2010+ formatted .xlsx files) so that you can just do "Save As PDF" on each sheet and then send it wherever its needed.

This program uses the Apache CLI and POI packages to provide the options for command line arguments and the ability to read/write Excel-formatted files.
