________________________________________________________________ 

Crystal Decisions Technical Support - vbnet_web_simplelogonViewer.exe

Phone: (604) 669-8379

Answers by Email: 
http://support.crystaldecisions.com/support/answers.asp
  
________________________________________________________________ 

PRODUCT VERSION

This application is for use with Visual Basic .NET and Crystal Reports for Visual Studio .NET.

________________________________________________________________ 

DESCRIPTION

This Visual Basic .NET sample Web application demonstrates how to log onto a secure database using the LogonInfo Property of the CrystalReportViewer control to establish a connection.  This application uses "Viewer" in its title because it uses objects in the CrystalDecisions.Shared assembly to set table logon information that can be passed through
the viewer. 

A connection to the database can also be established using the Database Table object of the ReportDocument class in the CrystalDecisions.CrystalReports.Engine assembly.

________________________________________________________________ 


FILES

- AssemblyInfo.vb
- CrystalReport1.rpt
- CrystalReport1.vb
- Global.asax
- Global.asax.resx
- Global.asax.vb
- README.txt
- Styles.css
- SimpleLogonViewer.sln
- SimpleLogonViewer.vbproj
- SimpleLogonViewer.vsdisco
- WebForm1.aspx
- WebForm1.aspx.resx
- WebForm1.aspx.vb
- Web.config


________________________________________________________________ 

INSTALLATION

1. After extracting the contents of this Zip file to the default installation folder, C:\Crystal\CRNET\crnet_web_vbnet_simplelogonviewer, create a virtual directory called 'CRNETSamples' for this application:

Note: ====

If you already created the 'CRNETSamples' virtual directory for previous samples,the following steps may not be necessary.

==========

To create a virtual directory:

a)  From the Start menu, select Run, and type 'inetmgr'.  Click 'OK' to load the IIS Manager.
b)  Expand the server node, and right-click on 'Default Web Site'.  Select 'New' | 'Virtual Directory'.
c)  Click the 'Next' button on the Virtual Directory Creation Wizard. 
d)  In the 'Alias' textbox, enter 'CRNETSamples' and click the 'Next' button.
e)  Browse to the 'C:\Crystal\CRNET\' folder, and click the 'Next' button.
f)  On the 'Access Permission' screen, click the 'Next' button to advance to the next screen.
g)  Click the 'Finish' button to create the virtual directory for this sample.

2.  Create application settings for this Directory in IIS.  Use the following steps to create the application settings:

a)  From the Start menu, select Run and type 'inetmgr'.  Click 'OK' to load the IIS Manager.
b)  Expand the server node and you will see the Default Web Site node.  For this node, expand the nodes, CRNETSamples | crnet_web_vbnet_simplelogonviewer.
c)  Right-click on the node 'crnet_web_vbnet_simplelogonviewer' and select 'Properties'.
d)  In the Properties window click the 'Create' button and click 'OK' to close the window.

3. Open the Visual Studio .NET IDE.

4. From the File menu, select 'Open' | 'Project', and browse to where you extracted the sample and choose the WebExportReport.sln.  Click the 'Open' button to load the project.

5. From the Build menu, select 'Build Solution' to compile the application.

6. From the Debug menu, select 'Start Without Debugging' to run the application. 

Note:======

To preview the report, you may need to right-click on the 'Webform1.aspx' page and select 'Set as Start Page' in the Solution Explorer.

===========

A report now appears with data from the Authors table in th at runtimee SQL Server Pubs sample database.
________________________________________________________________ 

ADDITIONAL INFO

1.  The report in the sample, uses the Pubs sample database that is installed by default with SQL Server.
2.  In order for the report provided with this application to run correctly, you will have to supply the
necessary information for connecting to your installation of SQL Server.  This information includes setting
the ServerName, UserID and Password properties for the ConnectionInfo object in the source code of the Webform1.aspx.vb file included with this application.  The report connects using OLE-DB and SQL Authentication. 


________________________________________________________________

TECH ID: DAK
________________________________________________________________ 

Last updated on Feb 8, 2002
________________________________________________________________ 
