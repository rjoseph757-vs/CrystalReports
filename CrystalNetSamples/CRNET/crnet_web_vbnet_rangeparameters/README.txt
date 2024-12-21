________________________________________________________________ 

Crystal Decisions Technical Support - vbnet_web_rangeparams.exe

Phone: (604) 669-8379

Answers by Email: 
http://support.crystaldecisions.com/support/answers.asp
  
________________________________________________________________ 

PRODUCT VERSION

This application is for use with Visual Basic .NET and Crystal Reports for Visual Studio .NET.

________________________________________________________________ 

DESCRIPTION

This sample Web application demonstrates how to pass multiple values to a single discrete parameter field.  In addition, this application passes a start and end value to a range parameter field.

________________________________________________________________ 


FILES

- AssemblyInfo.vb
- WebRangeParameter.rpt
- WebRangeParameter.vb
- Global.asax
- Global.asax.resx
- Global.asax.vb
- README.txt
- Styles.css
- Web.config
- WebRangeParameter.sln
- WebRangeParameter.vbproj
- webRangeParameter.vsdisco
- WebForm1.aspx
- WebForm1.aspx.resx
- WebForm1.aspx.vb

________________________________________________________________ 

INSTALLATION

1. After extracting the contents of this Zip file to the default installation folder, C:\Crystal\CRNET\crnet_web_vbnet_rangeparameters, create a virtual directory called 'CRNETSamples' for this application:

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
b)  Expand the server node and you will see the Default Web Site node.  For this node, expand the nodes, CRNETSamples | crnet_web_vbnet_rangeparameters.
c)  Right-click on the node 'crnet_web_vbnet_rangeparameters' and select 'Properties'.
d)  In the Properties window click the 'Create' button and click 'OK' to close the window.

3. Open the Visual Studio .NET IDE.

4. From the File menu, select 'Open' | 'Project', and browse to where you extracted the sample and choose WebRangeParameter.sln.  Click on the 'Open' button to load the project.

5. From the Build menu, select 'Build Solution' to compile the application.

6. From the Debug menu, select 'Start Without Debugging' to run the application.

A report now appears with the data specified by the parameter values passed in at runtime.

Note:======

To preview the report, you may need to right-click on the 'Webform1.aspx' page and select 'Set as Start Page' in the Solution Explorer.

===========

________________________________________________________________ 

ADDITIONAL INFO

The report in the sample, uses an MS Access Database (Xtreme.mdb).  Xtreme.mdb is the sample database Crystal Reports uses and is installed by default to: 

C:\Program Files\Microsoft Visual Studio .NET\Crystal Reports\Samples\Databases\


________________________________________________________________

TECH ID: DAK
________________________________________________________________ 

Last updated on Jan 3, 2002
________________________________________________________________ 
