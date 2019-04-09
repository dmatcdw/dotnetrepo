<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnRun = New System.Windows.Forms.Button
        Me.lblFunction = New System.Windows.Forms.Label
        Me.cbxFunctions = New System.Windows.Forms.ComboBox
        Me.txtXmlRequest = New System.Windows.Forms.TextBox
        Me.txtXmlType = New System.Windows.Forms.TextBox
        Me.lblXmlType = New System.Windows.Forms.Label
        Me.txtResult = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtElement = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'btnRun
        '
        Me.btnRun.Location = New System.Drawing.Point(350, 20)
        Me.btnRun.Name = "btnRun"
        Me.btnRun.Size = New System.Drawing.Size(75, 23)
        Me.btnRun.TabIndex = 0
        Me.btnRun.Text = "Run Test"
        Me.btnRun.UseVisualStyleBackColor = True
        '
        'lblFunction
        '
        Me.lblFunction.AutoSize = True
        Me.lblFunction.Location = New System.Drawing.Point(12, 20)
        Me.lblFunction.Name = "lblFunction"
        Me.lblFunction.Size = New System.Drawing.Size(48, 13)
        Me.lblFunction.TabIndex = 1
        Me.lblFunction.Text = "Function"
        '
        'cbxFunctions
        '
        Me.cbxFunctions.FormattingEnabled = True
        Me.cbxFunctions.Items.AddRange(New Object() {"ExtractFromXml", "ProcessMTA", "CreatePolicy"})
        Me.cbxFunctions.Location = New System.Drawing.Point(92, 20)
        Me.cbxFunctions.Name = "cbxFunctions"
        Me.cbxFunctions.Size = New System.Drawing.Size(242, 21)
        Me.cbxFunctions.TabIndex = 2
        '
        'txtXmlRequest
        '
        Me.txtXmlRequest.Location = New System.Drawing.Point(12, 67)
        Me.txtXmlRequest.Multiline = True
        Me.txtXmlRequest.Name = "txtXmlRequest"
        Me.txtXmlRequest.Size = New System.Drawing.Size(845, 192)
        Me.txtXmlRequest.TabIndex = 3
        '
        'txtXmlType
        '
        Me.txtXmlType.Location = New System.Drawing.Point(336, 290)
        Me.txtXmlType.Name = "txtXmlType"
        Me.txtXmlType.Size = New System.Drawing.Size(100, 20)
        Me.txtXmlType.TabIndex = 4
        '
        'lblXmlType
        '
        Me.lblXmlType.AutoSize = True
        Me.lblXmlType.Location = New System.Drawing.Point(263, 290)
        Me.lblXmlType.Name = "lblXmlType"
        Me.lblXmlType.Size = New System.Drawing.Size(56, 13)
        Me.lblXmlType.TabIndex = 5
        Me.lblXmlType.Text = "XML Type"
        '
        'txtResult
        '
        Me.txtResult.Location = New System.Drawing.Point(12, 328)
        Me.txtResult.Multiline = True
        Me.txtResult.Name = "txtResult"
        Me.txtResult.Size = New System.Drawing.Size(847, 130)
        Me.txtResult.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 290)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(45, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Element"
        '
        'txtElement
        '
        Me.txtElement.Location = New System.Drawing.Point(78, 290)
        Me.txtElement.Name = "txtElement"
        Me.txtElement.Size = New System.Drawing.Size(127, 20)
        Me.txtElement.TabIndex = 8
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(994, 477)
        Me.Controls.Add(Me.txtElement)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtResult)
        Me.Controls.Add(Me.lblXmlType)
        Me.Controls.Add(Me.txtXmlType)
        Me.Controls.Add(Me.txtXmlRequest)
        Me.Controls.Add(Me.cbxFunctions)
        Me.Controls.Add(Me.lblFunction)
        Me.Controls.Add(Me.btnRun)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnRun As System.Windows.Forms.Button
    Friend WithEvents lblFunction As System.Windows.Forms.Label
    Friend WithEvents cbxFunctions As System.Windows.Forms.ComboBox
    Friend WithEvents txtXmlRequest As System.Windows.Forms.TextBox
    Friend WithEvents txtXmlType As System.Windows.Forms.TextBox
    Friend WithEvents lblXmlType As System.Windows.Forms.Label
    Friend WithEvents txtResult As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtElement As System.Windows.Forms.TextBox

End Class
