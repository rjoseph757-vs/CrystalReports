Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class Form1
    Inherits System.Windows.Forms.Form
    Dim crReportDocument As New CrystalReport1()
    Dim GroupTreeToggle As Boolean
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
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Refresh As System.Windows.Forms.Button
    Friend WithEvents Print As System.Windows.Forms.Button
    Friend WithEvents Export As System.Windows.Forms.Button
    Friend WithEvents GroupTree As System.Windows.Forms.Button
    Friend WithEvents Zoom As System.Windows.Forms.ComboBox
    Friend WithEvents FindText As System.Windows.Forms.Button
    Friend WithEvents TextToFind As System.Windows.Forms.TextBox
    Friend WithEvents FirstPage As System.Windows.Forms.Button
    Friend WithEvents PageBack As System.Windows.Forms.Button
    Friend WithEvents GoToPage As System.Windows.Forms.Button
    Friend WithEvents PageForward As System.Windows.Forms.Button
    Friend WithEvents LastPage As System.Windows.Forms.Button
    Friend WithEvents Page As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.Refresh = New System.Windows.Forms.Button()
        Me.Print = New System.Windows.Forms.Button()
        Me.Export = New System.Windows.Forms.Button()
        Me.GroupTree = New System.Windows.Forms.Button()
        Me.Zoom = New System.Windows.Forms.ComboBox()
        Me.FindText = New System.Windows.Forms.Button()
        Me.TextToFind = New System.Windows.Forms.TextBox()
        Me.FirstPage = New System.Windows.Forms.Button()
        Me.PageBack = New System.Windows.Forms.Button()
        Me.GoToPage = New System.Windows.Forms.Button()
        Me.PageForward = New System.Windows.Forms.Button()
        Me.LastPage = New System.Windows.Forms.Button()
        Me.Page = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'CrystalReportViewer1
        '
        Me.CrystalReportViewer1.ActiveViewIndex = -1
        Me.CrystalReportViewer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.ReportSource = Nothing
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(664, 301)
        Me.CrystalReportViewer1.TabIndex = 0
        '
        'Refresh
        '
        Me.Refresh.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Refresh.Image = CType(resources.GetObject("Refresh.Image"), System.Drawing.Bitmap)
        Me.Refresh.Name = "Refresh"
        Me.Refresh.Size = New System.Drawing.Size(24, 32)
        Me.Refresh.TabIndex = 1
        '
        'Print
        '
        Me.Print.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Print.Image = CType(resources.GetObject("Print.Image"), System.Drawing.Bitmap)
        Me.Print.Location = New System.Drawing.Point(24, 0)
        Me.Print.Name = "Print"
        Me.Print.Size = New System.Drawing.Size(32, 32)
        Me.Print.TabIndex = 2
        '
        'Export
        '
        Me.Export.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Export.Image = CType(resources.GetObject("Export.Image"), System.Drawing.Bitmap)
        Me.Export.Location = New System.Drawing.Point(56, 0)
        Me.Export.Name = "Export"
        Me.Export.Size = New System.Drawing.Size(40, 32)
        Me.Export.TabIndex = 3
        '
        'GroupTree
        '
        Me.GroupTree.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.GroupTree.Image = CType(resources.GetObject("GroupTree.Image"), System.Drawing.Bitmap)
        Me.GroupTree.Location = New System.Drawing.Point(96, 0)
        Me.GroupTree.Name = "GroupTree"
        Me.GroupTree.Size = New System.Drawing.Size(32, 32)
        Me.GroupTree.TabIndex = 4
        '
        'Zoom
        '
        Me.Zoom.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Zoom.Location = New System.Drawing.Point(128, 0)
        Me.Zoom.Name = "Zoom"
        Me.Zoom.Size = New System.Drawing.Size(88, 33)
        Me.Zoom.TabIndex = 5
        Me.Zoom.Text = "ComboBox1"
        '
        'FindText
        '
        Me.FindText.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.FindText.Location = New System.Drawing.Point(216, 0)
        Me.FindText.Name = "FindText"
        Me.FindText.Size = New System.Drawing.Size(40, 32)
        Me.FindText.TabIndex = 6
        Me.FindText.Text = "Find Text:"
        '
        'TextToFind
        '
        Me.TextToFind.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextToFind.Location = New System.Drawing.Point(256, 0)
        Me.TextToFind.Name = "TextToFind"
        Me.TextToFind.Size = New System.Drawing.Size(136, 31)
        Me.TextToFind.TabIndex = 7
        Me.TextToFind.Text = ""
        '
        'FirstPage
        '
        Me.FirstPage.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.FirstPage.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FirstPage.Location = New System.Drawing.Point(392, 0)
        Me.FirstPage.Name = "FirstPage"
        Me.FirstPage.Size = New System.Drawing.Size(40, 32)
        Me.FirstPage.TabIndex = 8
        Me.FirstPage.Text = "<<"
        '
        'PageBack
        '
        Me.PageBack.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.PageBack.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PageBack.Location = New System.Drawing.Point(432, 0)
        Me.PageBack.Name = "PageBack"
        Me.PageBack.Size = New System.Drawing.Size(24, 32)
        Me.PageBack.TabIndex = 9
        Me.PageBack.Text = "<"
        '
        'GoToPage
        '
        Me.GoToPage.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.GoToPage.Location = New System.Drawing.Point(456, 0)
        Me.GoToPage.Name = "GoToPage"
        Me.GoToPage.Size = New System.Drawing.Size(32, 32)
        Me.GoToPage.TabIndex = 10
        Me.GoToPage.Text = "Go"
        '
        'PageForward
        '
        Me.PageForward.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.PageForward.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PageForward.Location = New System.Drawing.Point(536, 0)
        Me.PageForward.Name = "PageForward"
        Me.PageForward.Size = New System.Drawing.Size(24, 32)
        Me.PageForward.TabIndex = 11
        Me.PageForward.Text = ">"
        '
        'LastPage
        '
        Me.LastPage.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.LastPage.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LastPage.Location = New System.Drawing.Point(560, 0)
        Me.LastPage.Name = "LastPage"
        Me.LastPage.Size = New System.Drawing.Size(40, 32)
        Me.LastPage.TabIndex = 12
        Me.LastPage.Text = ">>"
        '
        'Page
        '
        Me.Page.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Page.Location = New System.Drawing.Point(488, 0)
        Me.Page.Name = "Page"
        Me.Page.Size = New System.Drawing.Size(48, 31)
        Me.Page.TabIndex = 13
        Me.Page.Text = "1"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(664, 301)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Page, Me.LastPage, Me.PageForward, Me.GoToPage, Me.PageBack, Me.FirstPage, Me.TextToFind, Me.FindText, Me.Zoom, Me.GroupTree, Me.Export, Me.Print, Me.Refresh, Me.CrystalReportViewer1})
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextToFind.Top = FindText.Top
        TextToFind.Height = FindText.Height

        Zoom.Text = ""
        Zoom.Items.Add(25)
        Zoom.Items.Add(50)
        Zoom.Items.Add(75)
        Zoom.Items.Add(100)
        Zoom.Items.Add(150)
        Zoom.Items.Add(200)
        Zoom.SelectedIndex = 3

        CrystalReportViewer1.ReportSource = crReportDocument
        CrystalReportViewer1.DisplayGroupTree = False
        GroupTreeToggle = True

        'get rid of toolbar
        With CrystalReportViewer1
            .ShowRefreshButton = False
            .ShowPrintButton = False
            .ShowExportButton = False
            .ShowGroupTreeButton = False
            .ShowZoomButton = False
            .ShowTextSearchButton = False
            .ShowPageNavigateButtons = False
            .ShowGotoPageButton = False
            .ShowCloseButton = False
        End With

        CrystalReportViewer1.Zoom(Val(Zoom.SelectedItem))
    End Sub

    Private Sub Refresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Refresh.Click
        CrystalReportViewer1.Refresh()
    End Sub

    Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click
        CrystalReportViewer1.PrintReport()
    End Sub

    Private Sub Export_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Export.Click
        CrystalReportViewer1.ExportReport()
    End Sub

    Private Sub GroupTree_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupTree.Click
        CrystalReportViewer1.DisplayGroupTree = GroupTreeToggle
        If GroupTreeToggle = True Then
            GroupTreeToggle = False
        Else
            GroupTreeToggle = True
        End If
    End Sub

    Private Sub Zoom_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Zoom.SelectedIndexChanged
        CrystalReportViewer1.Zoom(Val(Zoom.SelectedItem))
    End Sub

    Private Sub FindText_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindText.Click
        CrystalReportViewer1.SearchForText(TextToFind.Text)
    End Sub

    Private Sub FirstPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FirstPage.Click
        CrystalReportViewer1.ShowFirstPage()
        Page.Text = CrystalReportViewer1.GetCurrentPageNumber()
    End Sub

    Private Sub PageBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PageBack.Click
        CrystalReportViewer1.ShowPreviousPage()
        Page.Text = CrystalReportViewer1.GetCurrentPageNumber()
    End Sub

    Private Sub PageForward_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PageForward.Click
        CrystalReportViewer1.ShowNextPage()
        Page.Text = CrystalReportViewer1.GetCurrentPageNumber()
    End Sub

    Private Sub LastPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LastPage.Click
        CrystalReportViewer1.ShowLastPage()
        Page.Text = CrystalReportViewer1.GetCurrentPageNumber()
    End Sub

    Private Sub GoToPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GoToPage.Click
        CrystalReportViewer1.ShowNthPage(Val(Page.Text))
        Page.Text = CrystalReportViewer1.GetCurrentPageNumber()
    End Sub

End Class
