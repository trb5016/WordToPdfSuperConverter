<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SuperConverterForm
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
        Me.tbDocPath = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.butPickWordDoc = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.FontDialog1 = New System.Windows.Forms.FontDialog()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.butConvert = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'tbDocPath
        '
        Me.tbDocPath.Location = New System.Drawing.Point(83, 35)
        Me.tbDocPath.Name = "tbDocPath"
        Me.tbDocPath.Size = New System.Drawing.Size(610, 20)
        Me.tbDocPath.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(45, 38)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(32, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Path:"
        '
        'butPickWordDoc
        '
        Me.butPickWordDoc.Location = New System.Drawing.Point(699, 33)
        Me.butPickWordDoc.Name = "butPickWordDoc"
        Me.butPickWordDoc.Size = New System.Drawing.Size(27, 23)
        Me.butPickWordDoc.TabIndex = 2
        Me.butPickWordDoc.Text = "..."
        Me.butPickWordDoc.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(12, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(131, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Document to Convert:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(13, 94)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Advanced"
        '
        'butConvert
        '
        Me.butConvert.Location = New System.Drawing.Point(618, 61)
        Me.butConvert.Name = "butConvert"
        Me.butConvert.Size = New System.Drawing.Size(75, 23)
        Me.butConvert.TabIndex = 5
        Me.butConvert.Text = "Convert"
        Me.butConvert.UseVisualStyleBackColor = True
        '
        'SuperConverterForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(738, 274)
        Me.Controls.Add(Me.butConvert)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.butPickWordDoc)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.tbDocPath)
        Me.Name = "SuperConverterForm"
        Me.Text = "Word to PDF Super Converter"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents tbDocPath As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents butPickWordDoc As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents FontDialog1 As FontDialog
    Friend WithEvents Label2 As Label
    Friend WithEvents butConvert As Button
End Class
