________________________________________________________________ 

Crystal Decisions Technical Support - vbnet_win_dbengine.exe

Phone: (604) 669-8379

Answers by Email: 
http://support.crystaldecisions.com/support/answers.asp
  
________________________________________________________________ 

PRODUCT VERSION

This application is for use with Visual Basic .NET and Crystal Reports for Visual Studio .NET.

________________________________________________________________ 

DESCRIPTION

This Visual Basic .NET sample Windows application demonstrates how to connect to a secured database using the Crystal Reports for Visual Studio .NET Engine object model.

________________________________________________________________ 


FILES

- CrystalReport1.rpt
- CrystalReport1.vb
- Form1.resx
- Form1.vb
- DBConnectivity.sln
- DBConnectivity.vbproj
- Readme.txt



________________________________________________________________ 

INSTALLATION

1. Extract this file to the default installation folder, C:\Crystal\crnet\vbnet_win_dbengine. 

2. In the Visual Studio .NET IDE, click the 'File' menu and select 'Open'| 'Project'. Browse to where you extracted the sample files and select 'DBConnectivity.sln' to open the project.

3. From the Build menu, select 'Build Solution' to compile the application.

4. From the Debug menu, select "Start Without Debugging" to run the application.


A report appears after logging into the secured database using the engine object model.
________________________________________________________________ 

ADDITIONAL INFO

The report in this sample application uses a secured Microsoft SQL Server database called 'Pubs' (which is a sample database installed with MS SQL Server).  The report uses an OLE DB connection using the Microsoft SQL Server provider (SQLOLEDB).  

To preview the report you must have client connectivity to a MS SQL Server database that has the 'Pubs' database installed. You need to modify the logon code in the Form1 constructor, after the call to InitializeComponent().  The code sets database logon code for an object called 'crConnectionInfo', which is used in a 'With...' block. 

________________________________________________________________ 

TECH ID:  JAS
________________________________________________________________ 

Last updated on January 23, 2002
________________________________________________________________ 
