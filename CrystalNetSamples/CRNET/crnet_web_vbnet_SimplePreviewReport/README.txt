________________________________________________________________ 

Crystal Decisions Technical Support - vbnet_web_SimplePreviewReport.exe

Phone: (604) 669-8379

Answers by Email: 
http://support.crystaldecisions.com/support/answers.asp
  
________________________________________________________________ 

PRODUCT VERSION

This application is for use with Visual Basic .NET and Crystal Reports for Visual Studio .NET.

________________________________________________________________ 

DESCRIPTION

This sample Web application demonstrates six different methods of loading a report in the Web Forms viewer.

The types of reports set as the reportsource for the viewer are:

- By Report Name (from string path)
- By Report Object (from string path)
- By Untyped Report Component (from string path)
- By Report Object (added to Project)
- By Strongly-Typed Report Component
- By Cached Report Component

Note: =======

There are other methods for binding reports to a viewer such as a ServerFileReport, Web Service report and through the databindings properties of the viewer, but are not covered in this sample application.

=============
________________________________________________________________ 


FILES

- AssemblyInfo.vb
- World Sales Report.rpt
- World Sales Report.vb
- Global.asax
- Global.asax.resx
- Global.asax.vb
- README.txt
- Styles.css
- SimplePreviewReport.sln
- SimplePreviewReport.vbproj
- SimplePreviewReport.vsdisco
- WebForm1.aspx
- WebForm1.aspx.resx
- WebForm1.aspx.vb
- Web.config



________________________________________________________________ 

INSTALLATION

1. After extracting the contents of this Zip file to the default installation folder, C:\Crystal\CRNET\crnet_web_vbnet_SimplePreviewReport, create a virtual directory called 'CRNETSamples' for this application:

Note: ====

If you already created the 'CRNETSamples' virtual directory for previous samples, the following steps may not be necessary.

==========

To create a virtual directory:

a)  From the Start menu, select Run, and type 'inetmgr'.  Click 'OK' to load the IIS Manager.
b)  Expand the server node, and right-click on 'Default Web Site'.  Select 'New' | 'Virtual Directory'.
c)  Click the 'Next' button on the Virtual Directory Creation Wizard. 
d)  In the 'Alias' textbox, enter 'CRNETSamples' and click the 'Next' button.
e)  Browse to the 'C:\Crystal\CRNET\' folder, and click the 'Next' button.
f)  On the 'Access Permission' screen, click the 'Next' button to advance to the next screen.
g)  Click the 'Finish' button to create the virtual directory for this sample.

2.  You need to create application settings for this Directory in IIS.  Use the following steps to create the application settings.

a)  From the Start menu, select Run and type 'inetmgr'.  Click 'OK' to load the IIS Manager.
b)  Expand the server node and you will see the Default Web Site node.  For this node, expand the nodes, CRNETSamples | crnet_web_vbnet_SimplePreviewReport.
c)  Right-click on the node 'crnet_web_vbnet_SimplePreviewReport' and select 'Properties'.
d)  In the Properties window click the 'Create' button and click 'OK' to close the window.


3. Open the Visual Studio .NET IDE.

4. From the File menu, select 'Open' | 'Project', and browse to where you extracted the sample and choose SimplePreviewReport.sln.  Click on the 'Open' button to load the project.

5. From the Build menu, select 'Build Solution' to compile the application.

6. From the Debug menu, select 'Start Without Debugging' to run the application.

Note:======

To preview the report, you may need to right-click on the 'Webform1.aspx' page and select 'Set as Start Page' in the Solution Explorer.

===========

Use the drop down menu to select a method of loading the report and click the 'Load Report' button to view the report.

________________________________________________________________ 


ADDITIONAL INFO

The report in the sample, uses an MS Access Database (Xtreme.mdb).  Xtreme.mdb is the sample database Crystal Reports uses and is installed by default to: 

C:\Program Files\Microsoft Visual Studio .NET\Crystal Reports\Samples\Databases\


________________________________________________________________

TECH ID: DAK
________________________________________________________________ 

Last updated on Jan 9, 2002
________________________________________________________________ 
