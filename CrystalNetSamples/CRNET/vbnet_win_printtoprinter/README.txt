________________________________________________________________ 

Crystal Decisions Technical Support - vbnet_win_PrintToPrinter.exe

Phone: (604) 669-8379

Answers by Email: 
http://support.crystaldecisions.com/support/answers.asp
  
________________________________________________________________ 

PRODUCT VERSION

This application is for use with Visual Basic .NET and Crystal Reports for Visual Studio .NET.

________________________________________________________________ 

DESCRIPTION

This Visual Basic .NET sample Windows application demonstrates how to print a report directly to a printer at runtime using the engine object model.

________________________________________________________________ 


FILES

- Form1.resx
- Form1.vb
- PrintToPrinter.sln
- PrintToPrinter.vbproj
- Readme.txt


________________________________________________________________ 

INSTALLATION

1. After having extracted this file to the default installation folder, C:\Crystal\crnet\vbnet_win_PrintToPrinter\, open the Visual Studio .NET IDE.

2. From the File menu, select 'Open'| 'Project' and browse to where you extracted the sample and select 'PrintToPrinter.sln'.  Click on the 'Open' button to load the project.

3. From the Build menu, select 'Build Solution' to compile the application.

4. From the Debug menu, select 'Start Without Debugging' to run the application.

Running the application will display a form with a button.  Click the button to print the report to the printer.  A dialog will be displayed when the printing is complete.
________________________________________________________________ 

ADDITIONAL INFO

1.  This sample application uses a sample report called, 'Income Statement.rpt', which is installed with Crystal Reports for Visual Studio .NET.  The sample report uses a Microsoft Access Database (Xtreme.mdb) installed by default to: 

C:\Program Files\Microsoft Visual Studio .NET\Crystal Reports\Samples\Databases\

2.  The sample report does not have a default printer specified.   You will have to modify the printer name in code to use your default printer.  To find the printer name, you can print out a test page (Control Panel; properties for your printer).

________________________________________________________________ 

TECH ID:  JAS
________________________________________________________________ 

Last updated on April 12, 2002
________________________________________________________________ 