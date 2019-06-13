Imports System.ComponentModel
Imports System.IO

Public Class SuperConverterForm

    Private Sub butPickWordDoc_Click(sender As Object, e As EventArgs) Handles butPickWordDoc.Click
        Dim fd As New OpenFileDialog

        fd.Filter = "Word Documents|*.doc;*.docx"
        fd.RestoreDirectory = True
        fd.Multiselect = False

        If fd.ShowDialog() = DialogResult.OK Then
            tbDocPath.Text = fd.FileName
        End If

    End Sub

    Private Sub SuperConverterForm_Load(sender As Object, e As EventArgs) Handles Me.Load

        If HelperFunctions.OpenArgSourceDocument <> "" Then
            tbDocPath.Text = HelperFunctions.OpenArgSourceDocument
        Else
            Call LoadParameters()
        End If

    End Sub

    Private Sub ConvertToPdf()

        Call SaveParameters()

        Dim docPath As String = tbDocPath.Text

        If File.Exists(docPath) Then

            If Path.GetExtension(docPath) = ".doc" Or Path.GetExtension(docPath) = ".docx" Then

                Call SuperConverter.ConvertWordLinksToPdfandMergeAll(docPath, GetFormParameters())
            Else
                MsgBox("You must select a word document")

            End If

        Else
            MsgBox("No file exists for the document path provided")
        End If

        MsgBox("Conversion Complete")

    End Sub

    Private Sub butConvert_Click(sender As Object, e As EventArgs) Handles butConvert.Click
        Call ConvertToPdf()
    End Sub

    Private Sub tbMarginOffset_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles tbMarginOffset.Validating

        If IsNumeric(tbMarginOffset.Text) AndAlso CInt(tbMarginOffset.Text) >= 0 AndAlso CInt(tbMarginOffset.Text) <= 200 Then
            'do nothing
        Else
            MsgBox("You must enter an integer between 0 and 200")
            e.Cancel = True
        End If

    End Sub

    Private Function GetFormParameters() As FormParameters

        Dim p As New FormParameters

        p.AddPages = cbIncludePageNumbers.Checked
        p.PagePrefix = tbPageNumberPrefixText.Text

        p.AddHeaders = cbIncludeHeaders.Checked
        p.HeaderRegExFind = tbReplaceRegExFind.Text
        p.HeaderRegExReplace = tbReplaceRegExWith.Text

        p.AddReturnLinks = cbAddReturnLinks.Checked

        p.MarginOffset = CSng(tbMarginOffset.Text)

        Return p

    End Function

    Private Sub SaveParameters()

        My.Settings.LastPath = tbDocPath.Text
        My.Settings.LastPageNum = cbIncludePageNumbers.Checked
        My.Settings.LastPrefix = tbPageNumberPrefixText.Text
        My.Settings.LastHeaders = cbIncludeHeaders.Checked
        My.Settings.LastRegexFind = tbReplaceRegExFind.Text
        My.Settings.LastRegexReplace = tbReplaceRegExWith.Text
        My.Settings.LastReturnLinks = cbAddReturnLinks.Checked
        My.Settings.LastMargin = tbMarginOffset.Text

    End Sub

    Private Sub LoadParameters()

        tbDocPath.Text = My.Settings.LastPath
        cbIncludePageNumbers.Checked = My.Settings.LastPageNum
        tbPageNumberPrefixText.Text = My.Settings.LastPrefix
        cbIncludeHeaders.Checked = My.Settings.LastHeaders
        tbReplaceRegExFind.Text = My.Settings.LastRegexFind
        tbReplaceRegExWith.Text = My.Settings.LastRegexReplace
        cbAddReturnLinks.Checked = My.Settings.LastReturnLinks
        tbMarginOffset.Text = My.Settings.LastMargin

    End Sub

    Private Sub SuperConverterForm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Call SaveParameters()
    End Sub

End Class