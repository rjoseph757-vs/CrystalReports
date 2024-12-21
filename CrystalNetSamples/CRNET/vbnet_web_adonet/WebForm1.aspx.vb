'********************************************
'Created:       June 30, 2003
'Author ID:     HAN
'Purpose:       This Visual Basic .NET Web sample application demonstrates how to '               build and populate an ADO.NET dataset and pass the dataset to a 
'               report at runtime.
'********************************************


'This CR assembly is required to be able to access, load, and pass a dataset to the report
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
'These two assemblies are used to connect to a datasource and retrieve records to populate a dataset object
Imports System.Data
Imports System.Data.OleDb

Public Class WebForm1

    Inherits System.Web.UI.Page
    Protected WithEvents CrystalReportViewer1 As CrystalDecisions.Web.CrystalReportViewer

    ''CR Variable
    Dim crReportDocument As CrystalReport1

    ''ADO.NET Variables
    Dim adoOleDbConnection As OleDbConnection
    Dim adoOleDbDataAdapter As OleDbDataAdapter
    Dim dataSet As DataSet

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here

        ''Build a connection string
        Dim connectionString As String = ""
        connectionString = "Provider=SQLOLEDB;"
        connectionString += "Server=rcon1;Database=pubs;"
        connectionString += "User ID=vantech;Password=vantech"

        ''Create and open a connection using the connection string
        adoOleDbConnection = New OleDbConnection(connectionString)

        ''Build a SQL statement to query the datasource
        Dim sqlString As String = ""
        sqlString = "Select * From Authors"

        ''Retrieve the data using the SQL statement and existing connection
        adoOleDbDataAdapter = New OleDbDataAdapter(sqlString, adoOleDbConnection)

        ''Create a instance of a Dataset
        dataSet = New DataSet()

        ''Fill the dataset with the data retrieved.  The name of the table
        ''in the dataset must be the same as the table name in the report.
        adoOleDbDataAdapter.Fill(dataSet, "Authors")

        ''Create an instance of the strongly-typed report object
        crReportDocument = New CrystalReport1()

        ''Pass the populated dataset to the report
        crReportDocument.SetDataSource(dataSet)

        ''Set the viewer to the report object to be previewed.
        CrystalReportViewer1.ReportSource = crReportDocument

    End Sub

End Class
