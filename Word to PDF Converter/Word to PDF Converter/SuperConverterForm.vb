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
            'Call ConvertToPdf()
        End If

    End Sub

    Private Sub ConvertToPdf()

        Dim docPath As String = tbDocPath.Text

        If File.Exists(docPath) Then

            If Path.GetExtension(docPath) = ".doc" Or Path.GetExtension(docPath) = ".docx" Then
                Call SuperConverter.ConvertWordLinksToPdfandMergeAll(docPath)
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
End Class