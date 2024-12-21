'********************************************
'Created:       January 17, 2002
'Author ID:     JAS
'Purpose:       This Visual Basic .NET sample Windows application demonstrates how to '               pass multiple values to a single discrete parameter field using the
'               engine object model.  
'********************************************

Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class Form1
    Inherits System.Windows.Forms.Form

    ''CR Variables
    Dim crReportDocument As CRParams
    Dim crParameterFieldDefinitions As ParameterFieldDefinitions
    Dim crParameterFieldDefinition As ParameterFieldDefinition
    Dim crParameterValues As ParameterValues
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Dim crParameterDiscreteValue As ParameterDiscreteValue

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        ''Create an instance of the strongly-typed report object
        crReportDocument = New CRParams()

        ''Get the collection of parameters from the report
        crParameterFieldDefinitions = crReportDocument.DataDefinition.ParameterFields
        ''Access the specified parameter from the collection
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("Country")

        ''Get the current values from the parameter field.  At this point
        ''there are zero values set.
        crParameterValues = crParameterFieldDefinition.CurrentValues

        ''Set the current values for the parameter field
        crParameterDiscreteValue = New ParameterDiscreteValue()
        crParameterDiscreteValue.Value = "Canada" '1st current value

        ''Add the first current value for the parameter field
        crParameterValues.Add(crParameterDiscreteValue)

        ''Since this parameter allows multiple values, the discrete value
        ''object needs to be reset.  Destroy the previous instance and create
        ''a new instance.  
        crParameterDiscreteValue = Nothing
        crParameterDiscreteValue = New ParameterDiscreteValue()
        crParameterDiscreteValue.Value = "USA" '2nd current value

        ''Add the second current value for the parameter field
        crParameterValues.Add(crParameterDiscreteValue)

        ''All current parameter values must be applied for the parameter field.
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        ''Set the viewer to the report object to be previewed.
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
