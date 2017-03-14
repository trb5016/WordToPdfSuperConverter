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
        Me.Label2 = New System.Windows.Forms.Label()
        Me.butConvert = New System.Windows.Forms.Button()
        Me.cbIncludePageNumbers = New System.Windows.Forms.CheckBox()
        Me.cbIncludeHeaders = New System.Windows.Forms.CheckBox()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.HeaderOptionsPanel = New System.Windows.Forms.Panel()
        Me.tbReplaceRegExWith = New System.Windows.Forms.TextBox()
        Me.tbReplaceRegExFind = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.pageNumberOptionsPanel = New System.Windows.Forms.Panel()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.tbPageNumberPrefixText = New System.Windows.Forms.TextBox()
        Me.cbAddReturnLinks = New System.Windows.Forms.CheckBox()
        Me.MarginOffsetPanel = New System.Windows.Forms.Panel()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.tbMarginOffset = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.HeaderOptionsPanel.SuspendLayout()
        Me.pageNumberOptionsPanel.SuspendLayout()
        Me.MarginOffsetPanel.SuspendLayout()
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
        Me.Label2.Size = New System.Drawing.Size(54, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Options:"
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
        'cbIncludePageNumbers
        '
        Me.cbIncludePageNumbers.AutoSize = True
        Me.cbIncludePageNumbers.Checked = True
        Me.cbIncludePageNumbers.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbIncludePageNumbers.Location = New System.Drawing.Point(3, 3)
        Me.cbIncludePageNumbers.Name = "cbIncludePageNumbers"
        Me.cbIncludePageNumbers.Size = New System.Drawing.Size(154, 17)
        Me.cbIncludePageNumbers.TabIndex = 6
        Me.cbIncludePageNumbers.Text = "Add Overall Page Numbers"
        Me.cbIncludePageNumbers.UseVisualStyleBackColor = True
        '
        'cbIncludeHeaders
        '
        Me.cbIncludeHeaders.AutoSize = True
        Me.cbIncludeHeaders.Checked = True
        Me.cbIncludeHeaders.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbIncludeHeaders.Location = New System.Drawing.Point(3, 32)
        Me.cbIncludeHeaders.Name = "cbIncludeHeaders"
        Me.cbIncludeHeaders.Size = New System.Drawing.Size(240, 17)
        Me.cbIncludeHeaders.TabIndex = 9
        Me.cbIncludeHeaders.Text = "Add Headers (Based on attachment filename)"
        Me.cbIncludeHeaders.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 44.89796!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 55.10204!))
        Me.TableLayoutPanel1.Controls.Add(Me.cbAddReturnLinks, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.HeaderOptionsPanel, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.cbIncludePageNumbers, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.cbIncludeHeaders, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.pageNumberOptionsPanel, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.MarginOffsetPanel, 0, 4)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(48, 110)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 5
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(588, 186)
        Me.TableLayoutPanel1.TabIndex = 10
        '
        'HeaderOptionsPanel
        '
        Me.HeaderOptionsPanel.AutoSize = True
        Me.HeaderOptionsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.HeaderOptionsPanel.Controls.Add(Me.tbReplaceRegExWith)
        Me.HeaderOptionsPanel.Controls.Add(Me.tbReplaceRegExFind)
        Me.HeaderOptionsPanel.Controls.Add(Me.Label7)
        Me.HeaderOptionsPanel.Controls.Add(Me.Label6)
        Me.HeaderOptionsPanel.Controls.Add(Me.Label5)
        Me.HeaderOptionsPanel.Location = New System.Drawing.Point(267, 32)
        Me.HeaderOptionsPanel.Name = "HeaderOptionsPanel"
        Me.HeaderOptionsPanel.Size = New System.Drawing.Size(189, 71)
        Me.HeaderOptionsPanel.TabIndex = 11
        '
        'tbReplaceRegExWith
        '
        Me.tbReplaceRegExWith.Location = New System.Drawing.Point(78, 48)
        Me.tbReplaceRegExWith.Name = "tbReplaceRegExWith"
        Me.tbReplaceRegExWith.Size = New System.Drawing.Size(108, 20)
        Me.tbReplaceRegExWith.TabIndex = 4
        '
        'tbReplaceRegExFind
        '
        Me.tbReplaceRegExFind.Location = New System.Drawing.Point(78, 22)
        Me.tbReplaceRegExFind.Name = "tbReplaceRegExFind"
        Me.tbReplaceRegExFind.Size = New System.Drawing.Size(108, 20)
        Me.tbReplaceRegExFind.TabIndex = 3
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(22, 51)
        Me.Label7.Margin = New System.Windows.Forms.Padding(3)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(32, 13)
        Me.Label7.TabIndex = 2
        Me.Label7.Text = "With:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(22, 25)
        Me.Label6.Margin = New System.Windows.Forms.Padding(3)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(50, 13)
        Me.Label6.TabIndex = 1
        Me.Label6.Text = "Replace:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 3)
        Me.Label5.Margin = New System.Windows.Forms.Padding(3)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(183, 13)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Filename Text Replacement (RegEx):"
        '
        'pageNumberOptionsPanel
        '
        Me.pageNumberOptionsPanel.AutoSize = True
        Me.pageNumberOptionsPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.pageNumberOptionsPanel.Controls.Add(Me.Label4)
        Me.pageNumberOptionsPanel.Controls.Add(Me.tbPageNumberPrefixText)
        Me.pageNumberOptionsPanel.Location = New System.Drawing.Point(267, 3)
        Me.pageNumberOptionsPanel.Name = "pageNumberOptionsPanel"
        Me.pageNumberOptionsPanel.Size = New System.Drawing.Size(314, 23)
        Me.pageNumberOptionsPanel.TabIndex = 13
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 3)
        Me.Label4.Margin = New System.Windows.Forms.Padding(3)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(128, 13)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "Page Number Prefix Text:"
        '
        'tbPageNumberPrefixText
        '
        Me.tbPageNumberPrefixText.Location = New System.Drawing.Point(137, 0)
        Me.tbPageNumberPrefixText.Name = "tbPageNumberPrefixText"
        Me.tbPageNumberPrefixText.Size = New System.Drawing.Size(174, 20)
        Me.tbPageNumberPrefixText.TabIndex = 13
        Me.tbPageNumberPrefixText.Text = "Page "
        '
        'cbAddReturnLinks
        '
        Me.cbAddReturnLinks.AutoSize = True
        Me.cbAddReturnLinks.Checked = True
        Me.cbAddReturnLinks.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbAddReturnLinks.Location = New System.Drawing.Point(3, 109)
        Me.cbAddReturnLinks.Name = "cbAddReturnLinks"
        Me.cbAddReturnLinks.Size = New System.Drawing.Size(172, 17)
        Me.cbAddReturnLinks.TabIndex = 11
        Me.cbAddReturnLinks.Text = "Add return links to attachments"
        Me.cbAddReturnLinks.UseVisualStyleBackColor = True
        '
        'MarginOffsetPanel
        '
        Me.MarginOffsetPanel.AutoSize = True
        Me.MarginOffsetPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.TableLayoutPanel1.SetColumnSpan(Me.MarginOffsetPanel, 2)
        Me.MarginOffsetPanel.Controls.Add(Me.tbMarginOffset)
        Me.MarginOffsetPanel.Controls.Add(Me.Label8)
        Me.MarginOffsetPanel.Location = New System.Drawing.Point(3, 152)
        Me.MarginOffsetPanel.Name = "MarginOffsetPanel"
        Me.MarginOffsetPanel.Size = New System.Drawing.Size(196, 26)
        Me.MarginOffsetPanel.TabIndex = 11
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(3, 6)
        Me.Label8.Margin = New System.Windows.Forms.Padding(3)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(116, 13)
        Me.Label8.TabIndex = 0
        Me.Label8.Text = "Margin offset of above:"
        '
        'tbMarginOffset
        '
        Me.tbMarginOffset.Location = New System.Drawing.Point(125, 3)
        Me.tbMarginOffset.Name = "tbMarginOffset"
        Me.tbMarginOffset.Size = New System.Drawing.Size(68, 20)
        Me.tbMarginOffset.TabIndex = 1
        Me.tbMarginOffset.Text = "10"
        '
        'SuperConverterForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(738, 314)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.butConvert)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.butPickWordDoc)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.tbDocPath)
        Me.Name = "SuperConverterForm"
        Me.Text = "Word to PDF Super Converter"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.HeaderOptionsPanel.ResumeLayout(False)
        Me.HeaderOptionsPanel.PerformLayout()
        Me.pageNumberOptionsPanel.ResumeLayout(False)
        Me.pageNumberOptionsPanel.PerformLayout()
        Me.MarginOffsetPanel.ResumeLayout(False)
        Me.MarginOffsetPanel.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents tbDocPath As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents butPickWordDoc As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents butConvert As Button
    Friend WithEvents cbIncludePageNumbers As CheckBox
    Friend WithEvents cbIncludeHeaders As CheckBox
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents pageNumberOptionsPanel As Panel
    Friend WithEvents Label4 As Label
    Friend WithEvents tbPageNumberPrefixText As TextBox
    Friend WithEvents HeaderOptionsPanel As Panel
    Friend WithEvents tbReplaceRegExWith As TextBox
    Friend WithEvents tbReplaceRegExFind As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents cbAddReturnLinks As CheckBox
    Friend WithEvents MarginOffsetPanel As Panel
    Friend WithEvents tbMarginOffset As TextBox
    Friend WithEvents Label8 As Label
End Class
