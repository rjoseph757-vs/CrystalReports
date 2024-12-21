'****************************
'
'File Name:     vbnet_win_viewerdemo
'Created:       July 16, 2003
'Author ID:     894
'Purpose:       This application demonstrates how to use the viewer to preview the report, set connection
'               information, parameters, selection formula, and drill down on groups.
'
'********************************

Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class Form1
    Inherits System.Windows.Forms.Form
    Dim crReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument()
    Dim crLogOnInfo As New TableLogOnInfo()
    Dim ConnFrm As New frmConn()
    Dim RecSelFormFrm As New frmRecSelForm()
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents MenuItem11 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu()
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.MenuItem11 = New System.Windows.Forms.MenuItem()
        Me.MenuItem6 = New System.Windows.Forms.MenuItem()
        Me.MenuItem4 = New System.Windows.Forms.MenuItem()
        Me.MenuItem5 = New System.Windows.Forms.MenuItem()
        Me.MenuItem2 = New System.Windows.Forms.MenuItem()
        Me.MenuItem8 = New System.Windows.Forms.MenuItem()
        Me.MenuItem3 = New System.Windows.Forms.MenuItem()
        Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.MenuItem7 = New System.Windows.Forms.MenuItem()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem5})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem11, Me.MenuItem6, Me.MenuItem4})
        Me.MenuItem1.Text = "File"
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 0
        Me.MenuItem11.Text = "Open Report"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 1
        Me.MenuItem6.Text = "View Report"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 2
        Me.MenuItem4.Text = "Exit"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 1
        Me.MenuItem5.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem7, Me.MenuItem2, Me.MenuItem8, Me.MenuItem3})
        Me.MenuItem5.Text = "Viewer Options"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.Text = "Set Parameters"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 2
        Me.MenuItem8.Text = "Set Selection Formula"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 3
        Me.MenuItem3.Text = "Drill Down"
        '
        'CrystalReportViewer1
        '
        Me.CrystalReportViewer1.ActiveViewIndex = -1
        Me.CrystalReportViewer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.ReportSource = Nothing
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(472, 277)
        Me.CrystalReportViewer1.TabIndex = 0
        '
        'OpenFileDialog1
        '
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 0
        Me.MenuItem7.Text = "Set Connection"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(472, 277)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.CrystalReportViewer1})
        Me.Menu = Me.MainMenu1
        Me.Name = "Form1"
        Me.Text = "vbnet_win_viewerdemo"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'unenable the menu that requires a report to be opened
        MenuItem5.Enabled = False
    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click
        'Exit application
        End
    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        'Redisplay report
        'the reportsource must be reset first to refresh the report
        'the parameters must be set through code immediatly after
        'you cannot prompt for the parameter value with a customer form,
        'as that pause gives Crystal time to prompt for parameters
        'with our standard prompt.  To make your own custom prompt use
        'the engine.
        CrystalReportViewer1.ReportSource = crReport
        Dim ParamValue As New ParameterDiscreteValue()
        ParamValue.Value = "TestValue2"
        CrystalReportViewer1.ParameterFieldInfo(0).CurrentValues.Clear()
        CrystalReportViewer1.ParameterFieldInfo(0).CurrentValues.Add(ParamValue)
    End Sub

    Private Sub MenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem8.Click
        'set the selection formula
        strRecSelForm = crReport.RecordSelectionFormula
        RecSelFormFrm.ShowDialog()
        CrystalReportViewer1.SelectionFormula = strRecSelForm
        'Redisplay report
        CrystalReportViewer1.ReportSource = crReport
    End Sub

    Private Sub MenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem6.Click
        'set the report to the viewer
        CrystalReportViewer1.ReportSource = crReport
    End Sub

    Private Sub OpenFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk

    End Sub

    Private Sub MenuItem11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem11.Click
        MenuItem5.Enabled = False
        MenuItem1.Enabled = False
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.ShowDialog()
        If OpenFileDialog1.FileName <> "" Then
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            crReport.Load(OpenFileDialog1.FileName, OpenReportMethod.OpenReportByDefault)
            MenuItem5.Enabled = True
            MenuItem1.Enabled = True
        End If
        System.Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        'drill down through code as thought you had double clicked
        'on a group.  This requires the group level, group name
        'and group path.  Group path is an integer, and requires
        'and integer variable.

        Dim groupInfo As New CrystalDecisions.Shared.TotallerNodeID()
        Dim intGP(1) As Integer
        intGP(0) = 1
        groupInfo.GroupLevel = 1
        groupInfo.GroupName = "2.00"
        groupInfo.GroupPath = intGP
        CrystalReportViewer1.DrillDownOnGroup(groupInfo)
    End Sub

    Private Sub MenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem7.Click
        'get the connection information
        crLogOnInfo = crReport.Database.Tables(0).LogOnInfo
        strServer = crLogOnInfo.ConnectionInfo.ServerName
        strDatabase = crLogOnInfo.ConnectionInfo.DatabaseName
        strUserID = crLogOnInfo.ConnectionInfo.UserID
        ConnFrm.ShowDialog()

        'set the report to the viewer
        CrystalReportViewer1.ReportSource = crReport

        'set the connection information
        CrystalReportViewer1.LogOnInfo(0).ConnectionInfo.ServerName = strServer
        CrystalReportViewer1.LogOnInfo(0).ConnectionInfo.DatabaseName = strDatabase
        CrystalReportViewer1.LogOnInfo(0).ConnectionInfo.UserID = strUserID
        CrystalReportViewer1.LogOnInfo(0).ConnectionInfo.Password = strPassword
    End Sub
End Class
