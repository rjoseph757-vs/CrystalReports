'****************************
'
'File Name:     FormulasAndDynamicReportCreation.sln
'Created:       Feb 19, 2002
'Author ID:     901
'Purpose:       This application demonstrates how to create a dynamic report 
'               based on user input. This application uses blank formula fields 
'               to specify which fields appear in the report and to specify a group field. 
'
'********************************

Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine

Public Class Form2
    Inherits System.Windows.Forms.Form

    Dim crReport As New CrystalReport1()

    Dim crFormulas As FormulaFieldDefinitions

    'formulas that will contain the field name
    Dim crFormulaTextField1 As FormulaFieldDefinition
    Dim crFormulaTextField2 As FormulaFieldDefinition
    Dim crFormulaTextField3 As FormulaFieldDefinition

    'formulas that will contain the field data
    Dim crFormulaDBField1 As FormulaFieldDefinition
    Dim crformulaDBField2 As FormulaFieldDefinition

    'formula that will contain the group information
    Dim crFormulaGroup1 As FormulaFieldDefinition

#Region " Windows Form Designer generated code "
    ' the constructor is being modified so it can receive values from form1
    Public Sub New(ByVal field1 As String, ByVal field2 As String, ByVal group1 As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'set the Formulas collection to the current report's formula
        'collection
        crFormulas = crReport.DataDefinition.FormulaFields

        'set the 5 formula fields in the order they appear on the
        'report.
        crFormulaTextField1 = crFormulas.Item(0)
        crFormulaTextField2 = crFormulas.Item(1)
        crFormulaTextField3 = crFormulas.Item(2)
        crFormulaDBField1 = crFormulas.Item(3)
        crformulaDBField2 = crFormulas.Item(4)
        crFormulaGroup1 = crFormulas.Item(5)

        'pass in the Field names  Chr(34) is double quotes character
        crFormulaTextField1.Text = Chr(34) & field1 & Chr(34)
        crFormulaTextField2.Text = Chr(34) & field2 & Chr(34)
        crFormulaTextField3.Text = Chr(34) & group1 & Chr(34)

        'pass in the database fields to be displayed
        'a database field needs to have a specific format like
        '  {Table.Field} ie.  {Customer.Customer ID}
        crFormulaDBField1.Text = "{Customer." & field1 & "}"
        crformulaDBField2.Text = "{Customer." & field2 & "}"

        'pass in the group that you want to group on
        crFormulaGroup1.Text = "{Customer." & group1 & "}"


        CrystalReportViewer1.ReportSource = crReport

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

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SuspendLayout()
        '
        'CrystalReportViewer1
        '
        Me.CrystalReportViewer1.ActiveViewIndex = -1
        Me.CrystalReportViewer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CrystalReportViewer1.Location = New System.Drawing.Point(0, 0)
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.ReportSource = Nothing
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(350, 314)
        Me.CrystalReportViewer1.TabIndex = 0
        '
        'Form2
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(350, 314)
        Me.Controls.Add(Me.CrystalReportViewer1)
        Me.Name = "Form2"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form2"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
