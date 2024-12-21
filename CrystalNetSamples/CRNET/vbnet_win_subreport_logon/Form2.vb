'********************************************
'
'File Name:     	LogOnParameters.sln
'Created:       	April 23, 2002
'Author ID:     	901
'Purpose:       	This Visual Basic .NET sample Windows application
'       			demonstrates how to pass SQL log on information to a
'                   main report and a subreport.
'
'********************************************


Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine

Public Class Form2
    Inherits System.Windows.Forms.Form

    Dim crSections As Sections
    Dim crSection As Section
    Dim crReportObjects As ReportObjects
    Dim crReportObject As ReportObject
    Dim crSubreportObject As SubreportObject

    Dim crReportDocument As CrystalReport1
    Dim crSubreportDocument As ReportDocument

    Dim crDatabase As Database
    Dim crTables As Tables
    Dim crTable As Table
    Dim crTableLogOnInfo As TableLogOnInfo
    Dim crConnectioninfo As ConnectionInfo



#Region " Windows Form Designer generated code "

    Public Sub New(ByVal ServerName As String, ByVal UserID As String, ByVal Password As String, ByVal DatabaseName As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'declare an instance of the report and the connectionInfo object

        crReportDocument = New CrystalReport1()
        crConnectioninfo = New ConnectionInfo()

        'pass the necessary parameters to the connectionInfo object
        With crConnectioninfo
            .ServerName = ServerName
            .UserID = UserID
            .Password = Password
            .DatabaseName = DatabaseName
        End With

        'set up the database and tables objects to refer to the current report
        crDatabase = crReportDocument.Database
        crTables = crDatabase.Tables

        'loop through all the tables and pass in the connection info
        For Each crTable In crTables
            crTableLogOnInfo = crTable.LogOnInfo
            crTableLogOnInfo.ConnectionInfo = crConnectioninfo
            crTable.ApplyLogOnInfo(crTableLogOnInfo)
        Next

        'set the crSections object to the current report's sections
        crSections = crReportDocument.ReportDefinition.Sections

        'loop through all the sections to find all the report objects
        For Each crSection In crSections
            crReportObjects = crSection.ReportObjects
            'loop through all the report objects to find all the subreports
            For Each crReportObject In crReportObjects
                If crReportObject.Kind = ReportObjectKind.SubreportObject Then
                    'you will need to typecast the reportobject to a subreport 
                    'object once you find it
                    crSubreportObject = CType(crReportObject, SubreportObject)

                    'open the subreport object
                    crSubreportDocument = crSubreportObject.OpenSubreport(crSubreportObject.SubreportName)

                    'set the database and tables objects to work with the subreport
                    crDatabase = crSubreportDocument.Database
                    crTables = crDatabase.Tables

                    'loop through all the tables in the subreport and 
                    'set up the connection info and apply it to the tables
                    For Each crTable In crTables
                        With crConnectioninfo
                            .ServerName = ServerName
                            .DatabaseName = DatabaseName
                            .UserID = UserID
                            .Password = Password
                        End With
                        crTableLogOnInfo = crTable.LogOnInfo
                        crTableLogOnInfo.ConnectionInfo = crConnectioninfo
                        crTable.ApplyLogOnInfo(crTableLogOnInfo)
                    Next

                End If
            Next
        Next

        'view the report
        CrystalReportViewer1.ReportSource = crReportDocument
        Me.WindowState = FormWindowState.Maximized
        'Add any initialization after the InitializeComponent() call

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
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer.
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
        'Form2
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 273)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.CrystalReportViewer1})
        Me.Name = "Form2"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form2"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub CrystalReportViewer1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CrystalReportViewer1.Load

    End Sub
End Class
