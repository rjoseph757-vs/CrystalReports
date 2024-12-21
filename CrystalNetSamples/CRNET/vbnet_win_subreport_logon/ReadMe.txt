________________________________________________________________ 

Crystal Decisions Technical Support - vbnet_win_subreport_logon.exe

Phone: (604) 669-8379

Answers by Email: 
http://support.crystaldecisions.com/support/answers.asp
  

________________________________________________________________ 

PRODUCT VERSION

This Windows Application is for use with Visual Basic .NET and Crystal Reports for Visual Studio .NET.

________________________________________________________________ 

DESCRIPTION

This VB .NET sample Windows application demonstrates how to pass SQL log on information to a
main report and a subreport.

This application uses an OLE DB connection to the Northwind database that SQL Server installs by default.  You can set the Server Name, UserId, Password at runtime through the interface.

________________________________________________________________ 

File	
		
LogOnParameters.sln	
LogOnParameters.vbproj	
Form1.vb		
Form2.vb		
CrystalReport1.vb	
AssemblyInfo.vb		
Form1.resx		
Form2.resx		
CrystalReport1.rpt	

________________________________________________________________ 

INSTALLATION

1. After having extracted this file to the default installation folder, C:\Crystal\crnet\vbnet_win_subreport_logon, open the Visual Studio .NET IDE.

2. From the 'File' menu, select 'Open'| 'Project' and browse to where you extracted the sample files and select 'LogOnParameters.sln'.  Click the 'Open' button to load the project.

3. From the Build menu, select 'Build Solution' to compile the application.

4. From the Debug menu, select 'Start Without Debugging' to run the application.

5. The 'Login form' appears.  Enter the Server name/DSN, User ID, Password, and Database.  Click 'Launch Report' to preview the report.

A report now appears displaying the main and subreport (after logging into the Northwind database).

________________________________________________________________ 

TECH ID:  901
________________________________________________________________ 

Last updated on April 23/2002
________________________________________________________________ 
