'********************************************
'Created:       January 17, 2002
'Author ID:     906
'Purpose:       This Visual Basic .NET sample Windows application demonstrates  '               how to pass in multiple values to a Currency Range Parameter.
'*********************************************

Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class Form1
    Inherits System.Windows.Forms.Form

    'Crystal Reports Variables
    Dim crReportDocument As MultiRangeParams
    Dim crParameterFieldDefinitions As ParameterFieldDefinitions
    Dim crParameterFieldDefinition As ParameterFieldDefinition
    Dim crParameterValues As ParameterValues
    Dim crParameterRangeValue As ParameterRangeValue


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Create an instance of the strongly-typed report object
        crReportDocument = New MultiRangeParams()

        'Retrieve the Parameters collection from the Report
        crParameterFieldDefinitions = crReportDocument.DataDefinition.ParameterFields

        'Access the individual parameter field "Sales Range"
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("Sales Range")

        'Cast the variable to hold the values to pass to the Report before execution
        crParameterValues = crParameterFieldDefinition.CurrentValues

        'Cast the variable to hold the range value for the parameter
        crParameterRangeValue = New ParameterRangeValue()

        'Set the 1st range and include the upper and lower bounds
        With crParameterRangeValue
            .EndValue = "1000"
            .LowerBoundType = RangeBoundType.BoundInclusive
            .StartValue = "100"
            .UpperBoundType = RangeBoundType.BoundInclusive
        End With


        'Apply the 1st range to the values to be passed to the Report and reset variables in preparation 
        'for setting the next range value
        crParameterValues.Add(crParameterRangeValue)
        crParameterRangeValue = Nothing
        crParameterRangeValue = New ParameterRangeValue()

        'Set the 2nd range and include the upper and lower bounds
        With crParameterRangeValue
            .EndValue = "1000000"
            .LowerBoundType = RangeBoundType.BoundInclusive
            .StartValue = "300000"
            .UpperBoundType = RangeBoundType.BoundInclusive
        End With


        'Apply the 2nd range to the values to be passed to the Report and reset variables in preparation 
        'for setting the next range value
        crParameterValues.Add(crParameterRangeValue)
        crParameterRangeValue = Nothing
        crParameterRangeValue = New ParameterRangeValue()

        'Set the 3rd range and include the upper and lower bounds
        With crParameterRangeValue
            .EndValue = "6000"
            .LowerBoundType = RangeBoundType.BoundInclusive
            .StartValue = "4000"
            .UpperBoundType = RangeBoundType.BoundInclusive
        End With

        'Apply the 3rd range to the values to be passed to the Report
        crParameterValues.Add(crParameterRangeValue)

        'Pass the parameter values back to the report
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        'Set the ReportSource of the CrystalReportViewer to the strongly typed Report included in the project
        CrystalReportViewer1.ReportSource() = crReportDocument


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
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(920, 581)
        Me.CrystalReportViewer1.TabIndex = 0
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(920, 581)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.CrystalReportViewer1})
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
