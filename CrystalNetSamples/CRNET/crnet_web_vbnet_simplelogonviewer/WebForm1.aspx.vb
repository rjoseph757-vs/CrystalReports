'********************************************
'
'File Name:     	SimpleLogonViewer.sln
'Created:       	February 8, 2002
'Author ID:     	DAK
'Purpose:       	This Visual Basic .NET sample Web application demonstrates 
'                   how to log onto a secure database using the LogonInfo Property 
'                   of the CrystalReportViewer control to establish a connection. 
'
'********************************************

Imports CrystalDecisions.Shared

Public Class WebForm1
    Inherits System.Web.UI.Page
    Protected WithEvents CrystalReportViewer1 As CrystalDecisions.Web.CrystalReportViewer

    ''CR variables
    Dim crReportDocument As CrystalReport1
    Dim crTableLogonInfos As TableLogOnInfos
    Dim crTableLogonInfo As TableLogOnInfo
    Dim crConnectionInfo As ConnectionInfo

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()

        ''Create an instance of the strongly-typed report object
        crReportDocument = New CrystalReport1()

        '' Create a new instance of the tablelogoninfos class which contains
        ''the tablelogoninfo objects for each table in the report
        crTableLogonInfos = New TableLogOnInfos()

        ''Create a new instance of the TableLogonInfo class which contains
        ''the connection information for each table
        crTableLogonInfo = New TableLogOnInfo()

        ''Set the connection properties
        ''Note: this sample report connects to the SQL Server sample Pubs database
        ''If you have access to SQL server, enter in your values for the ServerName,
        ''UserID and Password properties to connect.  The report included with this
        ''sample application connects using OLE-DB and SQL authentication.
        crConnectionInfo = New ConnectionInfo()
        With crConnectionInfo
            .ServerName = "EnterServerNameHere"
            .DatabaseName = "Pubs"
            .UserID = "EnterUserIDHere"
            .Password = "EnterPasswordHere"
        End With

        ''apply the connection information to each table.  If the TableName is not specified, the 
        ''logon to the database will fail.  Once the tablename and connection information have been
        ''set for each table object, it is added to the TableLogonInfos collection using the Add method.
        crTableLogonInfo.ConnectionInfo = crConnectionInfo
        crTableLogonInfo.TableName = "authors"
        crTableLogonInfos.Add(crTableLogonInfo)

        ''Pass the table information to the logoninfo property of the viewer
        CrystalReportViewer1.LogOnInfo = crTableLogonInfos

        CrystalReportViewer1.ReportSource = crReportDocument

    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here
    End Sub

End Class
