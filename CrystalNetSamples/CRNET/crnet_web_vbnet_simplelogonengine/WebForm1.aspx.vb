'********************************************
'
'File Name:     	SimpleLogonEngine.sln
'Created:       	Jan 17, 2002
'Author ID:     	DAK
'Purpose:       	This Visual Basic .NET sample web application
'                   demonstrates how to log onto a secure SQL database 
'                   using the database table objects of the ReportDocument 
'                   class to establish the connection.  This application 
'                   uses "Engine" in its title because it uses objects in 
'                   the CrystalDecisions.CrystalReports.Engine assembly to log 
'                   on to the database.
'
'********************************************

Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class WebForm1
    Inherits System.Web.UI.Page
    Protected WithEvents CrystalReportViewer1 As CrystalDecisions.Web.CrystalReportViewer

    ' CR variables
    Dim crReportDocument As CrystalReport1
    Dim crDatabase As Database
    Dim crTables As Tables
    Dim crTable As Table
    Dim crTableLogOnInfo As TableLogOnInfo
    Dim crConnectionInfo As ConnectionInfo

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()

        'Create an instance of the strongly-typed report object
        crReportDocument = New CrystalReport1()

        'Create a new instance of the connectioninfo object and
        'set its properties
        'Note: This sample report connects to the sample SQL Server database, Pubs,
        'If you have access to SQL server, enter in your values for the ServerName,
        'UserID and Password properties to connect.  The report included with this
        'sample application connects using OLE-DB and SQL authentication.

        crConnectionInfo = New ConnectionInfo()
        With crConnectionInfo
            .ServerName = "EnterServerNameHere"
            .DatabaseName = "Pubs"
            .UserID = "EnterUserIDHere"
            .Password = "EnterPasswordHere"
        End With

        'Get the tables collection from the report object
        crDatabase = crReportDocument.Database
        crTables = crDatabase.Tables

        'Apply the logon information to each table in the collection
        For Each crTable In crTables
            crTableLogOnInfo = crTable.LogOnInfo
            crTableLogOnInfo.ConnectionInfo = crConnectionInfo
            crTable.ApplyLogOnInfo(crTableLogOnInfo)
        Next

        'Once the connection to the database has been established for
        'each table in the report, the report object can be bound to the viewer
        'using the reportsource property of the viewer to display the report.
        CrystalReportViewer1.ReportSource = crReportDocument

    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here
    End Sub

End Class
