Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'This will populate the drop-down boxes to allow
        'the user to choose what fields and the grouping
        'that they want on the report

        With comboField2
            .Items.Add("Customer ID")
            .Items.Add("Customer Name")
            .Items.Add("Address1")
            .Items.Add("City")
            .Items.Add("Region")
            .Items.Add("Country")
            .Items.Add("Phone")
            .Items.Add("Last Year's Sales")
        End With

        With comboField1
            .Items.Add("Customer ID")
            .Items.Add("Customer Name")
            .Items.Add("Address1")
            .Items.Add("City")
            .Items.Add("Region")
            .Items.Add("Country")
            .Items.Add("Phone")
            .Items.Add("Last Year's Sales")
        End With

        With comboGroup1
            .Items.Add("City")
            .Items.Add("Region")
            .Items.Add("Country")
        End With

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
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents comboField1 As System.Windows.Forms.ComboBox
    Friend WithEvents comboField2 As System.Windows.Forms.ComboBox
    Friend WithEvents comboGroup1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.comboField1 = New System.Windows.Forms.ComboBox
        Me.comboField2 = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.comboGroup1 = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(19, 286)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(192, 37)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Run Report"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(250, 286)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(192, 37)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Exit"
        '
        'comboField1
        '
        Me.comboField1.Location = New System.Drawing.Point(134, 92)
        Me.comboField1.Name = "comboField1"
        Me.comboField1.Size = New System.Drawing.Size(192, 24)
        Me.comboField1.TabIndex = 2
        '
        'comboField2
        '
        Me.comboField2.Location = New System.Drawing.Point(134, 157)
        Me.comboField2.Name = "comboField2"
        Me.comboField2.Size = New System.Drawing.Size(192, 24)
        Me.comboField2.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(29, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(413, 46)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Select 2 database fields and a group field  you want to appear on the report from" & _
        " the drop-down boxes"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(134, 74)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(192, 18)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Field 1"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(134, 138)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(173, 19)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Field 2"
        '
        'comboGroup1
        '
        Me.comboGroup1.Location = New System.Drawing.Point(134, 222)
        Me.comboGroup1.Name = "comboGroup1"
        Me.comboGroup1.Size = New System.Drawing.Size(192, 24)
        Me.comboGroup1.TabIndex = 7
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(134, 203)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(192, 19)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Group by:"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(460, 356)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.comboGroup1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.comboField2)
        Me.Controls.Add(Me.comboField1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Using Formulas for Dynamic Report Creation"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'FORM2's Constructor has been modified to receive 3 parameters
        '(the 2 fields and the group)
        Dim Form2 As New Form2(comboField1.Text, comboField2.Text, comboGroup1.Text)

        Form2.Show()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        End
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
