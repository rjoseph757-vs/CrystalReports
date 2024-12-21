________________________________________________________________ 

Crystal Decisions Technical Support - vbnet_win_adodotnet.exe

Phone: (604) 669-8379

Answers by Email: 
http://support.crystaldecisions.com/support/answers.asp
  
________________________________________________________________ 

PRODUCT VERSION

This application is for use with Visual Basic .NET and Crystal Reports for Visual Studio .NET.

________________________________________________________________ 

DESCRIPTION

This Visual Basic .NET sample Windows application demonstrates how to build and populate an ADO.NET dataset and pass the dataset to a report at runtime.

________________________________________________________________ 


FILES

- CrystalReport1.rpt
- CrystalReport1.vb
- Dataset1.vb
- Dataset1.xsd
- Dataset1.xsx
- Form1.resx
- Form1.vb
- UsingDataSet.sln
- UsingDataSet.vbproj
- Readme.txt



________________________________________________________________ 

INSTALLATION

1. After having extracted this file to the default installation folder, C:\Crystal\crnet\vbnet_win_adodotnet\, open the Visual Studio .NET IDE.

2. From the File menu, select 'Open'| 'Project' and browse to where you extracted the sample and select 'UsingDataSet.sln'.  Click on the 'Open' button to load the project.

3. From the Build menu, select 'Build Solution' to compile the application.

4. From the Debug menu, select "Start Without Debugging" to run the application.

A report now appears with after having logged onto the secured database using the engine object model.
________________________________________________________________ 

ADDITIONAL INFO

The report in this sample application is based on a dataset schema file (Dataset1.xsd) that was designed using the Dataset designer in VS .NET.  

The dataset schema defines the 'authors' table from the Microsoft SQL Server sample database, 'pubs'.  This means you must have client connectivity to a Microsoft SQL Server database that has the 'Pubs' database installed.

You will be required to modify the ADO.NET code in the Form1 constructor, after the call to InitializeComponent().  The code used to build the connection string needs to be modified to use your Microsoft SQL Server logon parameters.


________________________________________________________________ 

TECH ID:  JAS
________________________________________________________________ 

Last updated on March 21, 2002
________________________________________________________________ 
