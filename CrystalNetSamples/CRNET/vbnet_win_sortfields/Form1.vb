'********************************************************************
'
'File Name:     	SortFieldReport.sln
'Created:       	May 9, 2003
'Author ID:     	ALD
'Purpose:       	This Visual Basic .NET sample Windows application
'                   to change a Sort Field and the Sort Field 
'                   direction at runtime.
'
'********************************************************************
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class Form1
    Inherits System.Windows.Forms.Form

    'Crystal variables 
    Dim crReportDocument As SortFieldsReport
    Dim crSortField As SortField
    Dim mySortFld As String
    Dim mySortDir As String

    Dim crSortDirection As SortDirection
    Dim crDatabaseFieldDefinition As DatabaseFieldDefinition

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        crReportDocument = New SortFieldsReport()


        'Retrieve the Current Report Sort Fields and add to ComboBox1
        For Each crSortField In crReportDocument.DataDefinition.SortFields

            With ComboBox1.Items
                .Add(crSortField.Field.Name.ToString)
            End With
        Next
        'Add Sort Direction options to Combobox2
        With ComboBox2.Items
            .Add("Ascending")
            .Add("Descending")
        End With

        'Bind the report to the viewer using the reportsource property and view the report.
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'CrystalReportViewer1
        '
        Me.CrystalReportViewer1.ActiveViewIndex = -1
        Me.CrystalReportViewer1.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.CrystalReportViewer1.DisplayGroupTree = False
        Me.CrystalReportViewer1.DockPadding.Top = 10
        Me.CrystalReportViewer1.Location = New System.Drawing.Point(0, 123)
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.ReportSource = Nothing
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(552, 450)
        Me.CrystalReportViewer1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(152, 24)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Choose a Field to Sort"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(16, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(152, 24)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Choose The Sort Order"
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(376, 16)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(112, 32)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Sort Field"
        '
        'ComboBox1
        '
        Me.ComboBox1.Location = New System.Drawing.Point(176, 16)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(184, 21)
        Me.ComboBox1.TabIndex = 4
        '
        'ComboBox2
        '
        Me.ComboBox2.Location = New System.Drawing.Point(176, 48)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(184, 21)
        Me.ComboBox2.TabIndex = 5
        '
        'Form1
        '
        Me.AllowDrop = True
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(552, 573)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.ComboBox2, Me.ComboBox1, Me.Button1, Me.Label2, Me.Label1, Me.CrystalReportViewer1})
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If ComboBox1.Text <> "" And ComboBox2.Text <> "" Then

            'This routine gets the sorting fields and direction from the Comboboxes and sorts the appropriate field in the report

            mySortFld = ComboBox1.Text
            mySortDir = ComboBox2.Text

            'make the selected field the primary sortfield in the report (ie: Item(0)).  Then set the
            'sort direction for the field based on the value selected in Combobox2

            crDatabaseFieldDefinition = crReportDocument.Database.Tables(0).Fields(mySortFld.ToString)

            crSortField = crReportDocument.DataDefinition.SortFields(0)

            crSortField.Field = crDatabaseFieldDefinition

            If mySortDir = "Ascending" Then
                crSortField.SortDirection = SortDirection.AscendingOrder
            Else
                crSortField.SortDirection = SortDirection.DescendingOrder
            End If

            'Bind the report to the viewer using the reportsource property and view the report.
            CrystalReportViewer1.ReportSource = crReportDocument
        Else
            If ComboBox1.Text = "" Then
                MsgBox("Please select a report sort field")
            End If
            If ComboBox2.Text = "" Then
                MsgBox("Please select a sort direction")
            End If
        End If

    End Sub

End Class
