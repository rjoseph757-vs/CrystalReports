'********************************************
'Created:       January 14, 2004
'Author ID:     SJL
'Purpose:       This Visual Basic .NET sample Windows application demonstrates how to '               pass a same dataset to the main report and the subreport
'********************************************

''All of these assemblies are required to be able to
''access, load, and set database logon.

Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Data
Imports System.Data.OleDb

Public Class Form1
    Inherits System.Windows.Forms.Form

    Dim crReportDocument As New CrystalReport1()
    Dim crOledbConnection As OleDbConnection
    Dim crOledbDataAdapter As OleDbDataAdapter
    Dim ds As DataSet

    Dim crSections As Sections
    Dim crSection As Section
    Dim crReportObjects As ReportObjects
    Dim crReportObject As ReportObject
    Dim crSubreportObject As SubreportObject
    Dim crSubReportDoc As ReportDocument

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        'Connection information to connect to the database server to retrieve data
        'This is connecting to SQL server through OLEDB 
        Dim ConnectionString As String = ""
        ConnectionString = "Provider=SQLOLEDB;"
        ConnectionString += "Server=dbconn1;Database=Pubs;"
        ConnectionString += "User ID=vantech;Password=vantech"

        
        crOledbConnection = New OleDbConnection(ConnectionString)

        Dim sqlString As String = ""
        sqlString = "Select * From authors"
        
        crOledbDataAdapter = New OleDbDataAdapter(sqlString, crOledbConnection)

        ds = New DataSet()

        'Fill up the dataset
        crOledbDataAdapter.Fill(ds, "authors")

        'Set the datasource for the main report
        crReportDocument.Database.Tables(0).SetDataSource(ds)

        'Go through each sections in the main report and identify the subreport 
        'by name
        crSections = crReportDocument.ReportDefinition.Sections

        For Each crSection In crSections
            crReportObjects = crSection.ReportObjects
            For Each crReportObject In crReportObjects
                If crReportObject.Kind = ReportObjectKind.SubreportObject Then
                    crSubreportObject = CType(crReportObject, SubreportObject)
                    crSubReportDoc = crSubreportObject.OpenSubreport(crSubreportObject.SubreportName)

                    'Once the correct subreport has been located pass it the 
                    'appropriate dataset
                    If crSubReportDoc.Name = "FirstSub" Then
                            crSubReportDoc.Database.Tables(0).SetDataSource(ds)
                        End If
                    End If
            Next
        Next
        CrystalReportViewer1.ReportSource = crReportDocument
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
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

End Class
