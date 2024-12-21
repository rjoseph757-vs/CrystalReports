________________________________________________________________ 

Crystal Decisions Technical Support - vbnet_web_adonet.exe

Self-serve Support:
http://support.crystaldecisions.com/

Email Support:
http://support.crystaldecisions.com/support/answers.asp 

Telephone Support:
http://www.crystaldecisions.com/contact/support.asp
  
________________________________________________________________ 

PRODUCT VERSION

This application is for use with Visual Basic .NET and Crystal Reports for Visual Studio .NET.

________________________________________________________________ 

DESCRIPTION

This Visual Basic .NET Web sample application demonstrates how to build and populate an ADO.NET dataset and pass the dataset to a report at runtime.
________________________________________________________________ 


FILES

- AssemblyInfo.vb
- CrystalReport1.rpt
- CrystalReport1.vb
- DataSet1.vb
- DataSet1.xsd
- DataSet1.xsx
- Global.asax
- Global.asax.resx
- Global.asax.vb
- README.txt
- Styles.css
- vbnet_web_adonet.sln
- vbnet_web_adonet.vbproj
- vbnet_web_adonet.vsdisco
- WebForm1.aspx
- WebForm1.aspx.resx
- WebForm1.aspx.vb
- Web.config


________________________________________________________________ 

INSTALLATION

1. After extracting the contents of this Zip file to the default installation folder, C:\Crystal\CRNET\vbnet_web_adonet, create a virtual directory called 'CRNETSamples' for this application:

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
b)  Expand the server node and you will see the Default Web Site node.  For this node, expand the nodes, CRNETSamples | vbnet_web_adonet.
c)  Right-click on the node 'vbnet_web_adonet' and select 'Properties'.
d)  In the Properties window click the 'Create' button and click 'OK' to close the window.


3. Open the Visual Studio .NET IDE.


4. From the File menu, select 'Open' | 'Project', and browse to where you extracted the sample and choose vbnet_web_adonet.sln.  Click on the 'Open' button to load the project.

5. From the Build menu, select 'Build Solution' to compile the application.

6. From the Debug menu, select 'Start Without Debugging' to run the application.

Note:======

To preview the report, you may need to right-click on the 'Webform1.aspx' page and select 'Set as Start Page' in the Solution Explorer.

===========

A report now appears with data from the Authors table in the SQL Server Pubs sample database. 
________________________________________________________________ 

ADDITIONAL INFO

The report in this sample application is based on a dataset schema file (Dataset1.xsd) that was designed using the Dataset designer in VS .NET.  

The dataset schema defines the 'authors' table from the Microsoft SQL Server sample database, 'pubs'.  This means you must have client connectivity to a Microsoft SQL Server database that has the 'Pubs' database installed.

You will be required to modify the ADO.NET code.  The code used to build the connection string needs to be modified to use your Microsoft SQL Server logon parameters.


________________________________________________________________

TECH ID: HAN

Last updated on July 1, 2003
________________________________________________________________ 
