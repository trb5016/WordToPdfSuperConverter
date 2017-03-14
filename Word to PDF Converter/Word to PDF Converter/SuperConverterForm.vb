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
        End If

    End Sub

    Private Sub ConvertToPdf()

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


    Public Class FormParameters

        Public Property AddPages As Boolean
        Public Property PagePrefix As String

        Public Property AddHeaders As Boolean
        Public Property HeaderRegExFind As String
        Public Property HeaderRegExReplace As String

        Public Property AddReturnLinks As Boolean

        Public Property MarginOffset As Single

    End Class

End Class