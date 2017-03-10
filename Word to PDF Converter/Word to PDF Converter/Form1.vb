Imports System.IO

Imports iTextSharp.text
Imports iTextSharp.text.pdf


Public Class Form1


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        Dim sourceDocPath As String = HelperFunctions.SourceDocument

        Call Conversion.ConvertWordLinksToPdfandMergeAll(sourceDocPath)

    End Sub


End Class
