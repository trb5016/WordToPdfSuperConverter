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
            Call LoadParameters(False)
        End If

    End Sub

    Private Sub ConvertToPdf()

        Call SaveParameters(False)

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

    Private Sub SaveParameters(defaults As Boolean)

        Dim settingPrefix = If(defaults, "Default", "Last")

        My.Settings(settingPrefix & "Path") = tbDocPath.Text
        My.Settings(settingPrefix & "PageNum") = cbIncludePageNumbers.Checked
        My.Settings(settingPrefix & "Prefix") = tbPageNumberPrefixText.Text
        My.Settings(settingPrefix & "Headers") = cbIncludeHeaders.Checked
        My.Settings(settingPrefix & "RegexFind") = tbReplaceRegExFind.Text
        My.Settings(settingPrefix & "RegexReplace") = tbReplaceRegExWith.Text
        My.Settings(settingPrefix & "ReturnLinks") = cbAddReturnLinks.Checked
        My.Settings(settingPrefix & "Margin") = CInt(tbMarginOffset.Text)

        My.Settings.Save()

    End Sub

    Private Sub LoadParameters(defaults As Boolean)

        Dim settingPrefix = If(defaults, "Default", "Last")

        tbDocPath.Text = My.Settings(settingPrefix & "Path")
        cbIncludePageNumbers.Checked = My.Settings(settingPrefix & "PageNum")
        tbPageNumberPrefixText.Text = My.Settings(settingPrefix & "Prefix")
        cbIncludeHeaders.Checked = My.Settings(settingPrefix & "Headers")
        tbReplaceRegExFind.Text = My.Settings(settingPrefix & "RegexFind")
        tbReplaceRegExWith.Text = My.Settings(settingPrefix & "RegexReplace")
        cbAddReturnLinks.Checked = My.Settings(settingPrefix & "ReturnLinks")
        tbMarginOffset.Text = My.Settings(settingPrefix & "Margin")

    End Sub

    Private Sub SuperConverterForm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Call SaveParameters(False)
    End Sub

    Private Sub ButtonResetDefaults_Click(sender As Object, e As EventArgs) Handles buttonResetDefaults.Click

        If MsgBox("Are you sure you want to clear these settings and reset to defaults?", vbYesNo) = vbYes Then

            Call LoadParameters(True)

        End If

    End Sub

    Private Sub ButtonSaveDefaults_Click(sender As Object, e As EventArgs) Handles buttonSaveDefaults.Click

        If MsgBox("Are you sure you want to overwrite the current defaults with these options?", vbYesNo) = vbYes Then
            Call SaveParameters(True)

            MsgBox("These settings are now the default settings.")

        End If

    End Sub


End Class