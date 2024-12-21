'********************************************
'Created:       January 17, 2002
'Author ID:     JAS
'Purpose:       This Visual Basic .NET sample Windows application demonstrates how to '               build and populate an ADO.NET dataset and pass the dataset to a 
'               report at runtime.
'********************************************

''This CR assembly is required to be able to
''access, load, and pass a dataset to the report
Imports CrystalDecisions.CrystalReports.Engine
''These two assemblies are used to connect to a datasource
''and retrieve records to populate a dataset object
Imports System.Data
Imports System.Data.OleDb

Public Class Form1
    Inherits System.Windows.Forms.Form

    ''CR Variable
    Dim crReportDocument As CrystalReport1

    ''ADO.NET Variables
    Dim adoOleDbConnection As OleDbConnection
    Dim adoOleDbDataAdapter As OleDbDataAdapter
    Dim dataSet As DataSet


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        ''Build a connection string
        Dim connectionString As String = ""
        connectionString = "Provider=SQLOLEDB;"
        connectionString += "Server=escalade;Database=pubs;"
        connectionString += "User ID=sa;Password=admin"

        ''Create and open a connection using the connection string
        adoOleDbConnection = New OleDbConnection(connectionString)

        ''Build a SQL statement to query the datasource
        Dim sqlString As String = ""
        sqlString = "Select * From authors"

        ''Retrieve the data using the SQL statement and existing connection
        adoOleDbDataAdapter = New OleDbDataAdapter(sqlString, adoOleDbConnection)

        ''Create a instance of a Dataset
        dataSet = New DataSet()

        ''Fill the dataset with the data retrieved.  The name of the table
        ''in the dataset must be the same as the table name in the report.
        adoOleDbDataAdapter.Fill(dataSet, "authors")

        ''Create an instance of the strongly-typed report object
        crReportDocument = New CrystalReport1()

        ''Pass the populated dataset to the report
        crReportDocument.SetDataSource(dataSet)

        ''Set the viewer to the report object to be previewed.
        CrystalReportViewer1.ReportSource = crReportDocument

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If Disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.SuspendLayout()
        '
        'CrystalReportViewer1
        '
        Me.CrystalReportViewer1.ActiveViewIndex = -1
        Me.CrystalReportViewer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.ReportSource = Nothing
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(292, 273)
        Me.CrystalReportViewer1.TabIndex = 0
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 273)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.CrystalReportViewer1})
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub CrystalReportViewer1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CrystalReportViewer1.Load

    End Sub
End Class
