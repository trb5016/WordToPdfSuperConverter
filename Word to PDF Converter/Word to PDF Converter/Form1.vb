
Public Class Form1


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        Dim sourceDocPath As String = HelperFunctions.SourceDocument

        Call SuperConverter.ConvertWordLinksToPdfandMergeAll(sourceDocPath)

    End Sub


End Class
