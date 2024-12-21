'********************************************
'Created:       January 24, 2002
'Author ID:     JAS
'Purpose:       This Visual Basic .NET sample Windows application demonstrates how to '               set database logon information for a report based on a secured
'               datasource using the viewer's object model. 
'********************************************

''Both of these CR assemblies are required to be able to
''access, load, and set database logon.
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class Form1
    Inherits System.Windows.Forms.Form

    'CR Variables
    Dim crReportDocument As CrystalReport1
    Dim crTableLogOnInfos As TableLogOnInfos
    Dim crTableLogOnInfo As TableLogOnInfo
    Dim crConnectionInfo As ConnectionInfo

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        'Create an instance of the strongly-typed report object
        crReportDocument = New CrystalReport1()

        'Setup the connection information structure to be used
        'to log onto the datasource for the report.
        crConnectionInfo = New ConnectionInfo()
        With crConnectionInfo
            .ServerName = "escalade"    'physical server name
            .DatabaseName = "Pubs"
            .UserID = "sa"
            .Password = "admin"
        End With

        'Define and set the logon information for the table
        crTableLogOnInfo = New TableLogOnInfo()
        With crTableLogOnInfo
            .ConnectionInfo = crConnectionInfo
            .TableName = "authors"
        End With

        'Pass the table info to a collection of tables.
        crTableLogOnInfos = New TableLogOnInfos()
        crTableLogOnInfos.Add(crTableLogOnInfo)

        'Pass the entire collection of table logon information to the viewer.
        CrystalReportViewer1.LogOnInfo = crTableLogOnInfos

        'Set the viewer to the report object to be previewed.  This must be done
        'after all the database logon has been set.
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
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(600, 429)
        Me.CrystalReportViewer1.TabIndex = 0
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(600, 429)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.CrystalReportViewer1})
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

#End Region

End Class
