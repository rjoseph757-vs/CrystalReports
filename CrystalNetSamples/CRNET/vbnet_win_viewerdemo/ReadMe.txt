________________________________________________________________ 

Crystal Decisions Technical Support - vbnet_win_viewerdemo.exe

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

This Microsoft Visual Basic .NET sample application demonstrates
how to use the Crystal Reports viewer to manipulate the report
through code.  
This application includes the following functionality:

- Open a report
- View a report
- set the connection information through the viewer
- set the parameters through the viewer
- drill down through code with the viewer
- set the record selection formula through the viewer

________________________________________________________________ 

FILES

AssemblyInfo.vb
CrystalReport1.rpt
CrystalReport1.vb
Form1.resx
Form1.vb
frmConn.resx
frmConn.vb
frmRecSelForm.resx
frmRecSelForm.vb
Module1.vb
vbnet_win_viewerdemo.sln
vbnet_win_viewerdemo.suo
vbnet_win_viewerdemo.vbproj
vbnet_win_viewerdemo.vbproj.user

(in the Bin directory)
AxInterop.CRVIEWER9Lib.dll
Interop.CRVIEWER9Lib.dll
vbnet_win_viewerdemo.exe
vbnet_win_viewerdemo.pdb

(in the Ojb directory)
AxInterop.CRVIEWER9Lib.dll
Interop.CRVIEWER9Lib.dll

(in the Obj/Debug directory)
vbnet_win_viewerdemo.exe
vbnet_win_viewerdemo.Form1.resources
vbnet_win_viewerdemo.frmConn.resources
vbnet_win_viewerdemo.frmRecSelForm.resources
vbnet_win_viewerdemo.pdb

(in the Obj/Debug/TempPE directory)
CrystalReport1.vb.dll

________________________________________________________________ 

INSTALLATION

1. After having extracted the contents of this file to the default installation folder, C:\Crystal\CRNET\vbnet_win_viewerdemo, open the Visual Studio .NET IDE.

2. From the File menu, select 'Open'| 'Project' and browse to where you extracted the sample and select 'vbnet_win_viewerdemo.sln'.  Click on the 'Open' button to load the project.

3. From the Build menu, select 'Build Solution' to compile the application.

4. From the Debug menu, select "Start Without Debugging" to run the application.

5. Under the 'File' menu, select 'Open Report'.  You can
   select either the sample report included with this application, or your
   own report.  If you select your own report, a connectivity error message
   will appear.  To resolve this error, select 'Set Connection' under the 
   'Viewer Options' menu, and set the connection information.

6. Under the 'File' menu, click 'View Report'.  Under the 'Viewer Options' menu, select 
   'Set Parameters', 'Set Selection Formula', and 'Drill Down' to set the parameter 
   fields, set the record selection formula, and drill down on a group,
   respectively.

________________________________________________________________ 

ADDITIONAL INFO 

The sample report uses a Microsoft Access Database (Xtreme.mdb).  Xtreme.mdb is the 
sample database installed by default to: 

C:\Program Files\Microsoft Visual Studio .NET\Crystal Reports\Samples\Databases\

The interface for entering the record selection formula is basic, and will
not assist you with constructing a valid selection formula.  You will need
to construct the formula from your own knowledge of the report.

If you use a report other than the sample report, you will need to adapt
the code to use your parameter fields and group information, when setting
the parameters and drill down on the groups, respectively.

________________________________________________________________ 

Tech ID LAM
Last updated on July 17, 2003
________________________________________________________________ 
