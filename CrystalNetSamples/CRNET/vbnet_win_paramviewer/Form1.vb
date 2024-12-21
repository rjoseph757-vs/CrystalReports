'********************************************
'Created:       January 30, 2002
'Author ID:     JAS
'Purpose:       This Visual Basic .NET sample Windows application demonstrates how to '               pass multiple values to a single discrete parameter field using the
'               viewer's object model.  
'********************************************

Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class Form1
    Inherits System.Windows.Forms.Form

    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer

    ''CR Variables
    Dim crReportDocument As CRParams
    Dim crParameterFields As ParameterFields
    Dim crParameterField As ParameterField
    Dim crParameterDiscreteValue As ParameterDiscreteValue

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        ''Create an instance of the strongly-typed report object
        crReportDocument = New CRParams()

        ''Create a new instance of a discrete parameter object to set the 
        ''first value for the parameter.
        crParameterDiscreteValue = New ParameterDiscreteValue()
        crParameterDiscreteValue.Value = "Canada"

        ''Define the parameter field to pass the parameter values to.
        crParameterField = New ParameterField()
        crParameterField.ParameterFieldName = "Country"
        ''Pass the first value to the discrete parameter
        crParameterField.CurrentValues.Add(crParameterDiscreteValue)


        ''Destroy the current instance of the discrete value
        crParameterDiscreteValue = Nothing
        ''Create a new instance of discrete value object to set the second 
        ''value for the parameter.
        crParameterDiscreteValue = New ParameterDiscreteValue()
        crParameterDiscreteValue.Value = "USA"

        ''Pass the second value to the discrete parameter
        crParameterField.CurrentValues.Add(crParameterDiscreteValue)

        ''Create an instance of the parameter fields collection, and 
        ''pass the discrete parameter with the two discrete values to the
        ''collection of parameter fields.
        crParameterFields = New ParameterFields()
        crParameterFields.Add(crParameterField)

        ''The collection of parameter fields must be set to the viewer
        CrystalReportViewer1.ParameterFieldInfo = crParameterFields

        ''Set the viewer to the report object to be previewed.  This
        ''must be done after the parameter information has been set.
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
