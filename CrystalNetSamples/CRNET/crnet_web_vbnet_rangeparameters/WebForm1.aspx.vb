'********************************************
'
'File Name:     WebRangeParameter.sln
'Created:       January 3, 2002
'Author ID:     DAK
'Purpose:       This sample Web application demonstrates how to pass 
'               multiple values to a single discrete parameter field.  
'               In addition, this application passes a start and end value 
'               to a range parameter field.
'
'********************************************


Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class WebForm1
    Inherits System.Web.UI.Page
    Protected WithEvents CrystalReportViewer1 As CrystalDecisions.Web.CrystalReportViewer

    '' CR variables
    Dim crReportDocument As WebRangeParameter
    Dim crParameterFields As ParameterFields
    Dim crParameterField As ParameterField
    Dim crParameterValues As ParameterValues
    Dim crParameterDiscreteValue As ParameterDiscreteValue
    Dim crParameterRangeValue As ParameterRangeValue


#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()

        '' This sample application demonstrates the use of Range parameters in a Report and
        '' passing the parameters to the report via the Web Forms Viewer.
        '' The report file included with this application has two parameter fields.  The
        '' First is a Discrete parameter called "Country" and will accept multiple discrete
        '' values.  The Second parameter is called "SalesRange" which accepts Range values.



        ''Create an instance of the strongly-typed report object
        crReportDocument = New WebRangeParameter()

        ''The viewer's reportsource must be set to a report before any
        ''parameter fields can be accessed.
        CrystalReportViewer1.ReportSource = crReportDocument

        ''Get the collection of parameters from the report
        crParameterFields = CrystalReportViewer1.ParameterFieldInfo

        ''Access the specified parameter from the collection
        crParameterField = crParameterFields.Item("Country")

        ''Get the current values from the parameter field.  At this point
        ''there are zero values set.
        crParameterValues = crParameterField.CurrentValues

        ''Set the current values for the parameter field
        crParameterDiscreteValue = New ParameterDiscreteValue()
        crParameterDiscreteValue.Value = "Canada"

        ''Add the first current value for the parameter field
        crParameterValues.Add(crParameterDiscreteValue)

        ''Since this parameter allows multiple values, the discrete value
        ''object needs to be reset.  Destroy the previous instance and create
        ''a new instance.
        crParameterDiscreteValue = Nothing

        crParameterDiscreteValue = New ParameterDiscreteValue()
        crParameterDiscreteValue.Value = "USA"

        ''Add the second current value for the parameter field
        crParameterValues.Add(crParameterDiscreteValue)

        '' access the specified parameter from the collection.  This parameter is a Range parameter
        crParameterField = crParameterFields.Item("SalesRange")

        ''Get the current values from the parameter field.  At this point
        ''there are zero values set.
        crParameterValues = crParameterField.CurrentValues

        crParameterRangeValue = New ParameterRangeValue()
        ''Set the current values for the Start and End valuse of the parameter field

        With crParameterRangeValue
            .EndValue = "50000"
            .LowerBoundType = RangeBoundType.BoundInclusive
            .StartValue = "35000"
            .UpperBoundType = RangeBoundType.BoundInclusive
        End With

        ''Add the range current value for the parameter field
        crParameterValues.Add(crParameterRangeValue)

        ''Set the modified parameters collection back to the viewer so that
        ''the new parameter information can be used for the report.
        CrystalReportViewer1.ParameterFieldInfo = crParameterFields

    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here
    End Sub

End Class
