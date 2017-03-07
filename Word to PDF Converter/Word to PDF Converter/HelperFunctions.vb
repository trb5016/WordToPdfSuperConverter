Imports Word = Microsoft.Office.Interop.Word
Imports System.IO

Public Class HelperFunctions

    Public Shared Function SourceDocument()

        Dim args As String() = Environment.GetCommandLineArgs

        If args.Count > 1 Then

            Return args(1)
        Else
            Return ""
            MsgBox("Need to pass a word file as an open argument")
        End If

    End Function

    Public Shared Function ConvertWordToPdf(ByVal SourcePath As String, ByVal OutputPath As String) As Boolean

        Dim wordApp As New Word.Application
        Dim wordDoc As Word.Document
        Dim bFailed As Boolean = False

        Try
            wordDoc = wordApp.Documents.Open(SourcePath)    'Opens the source document in word
            File.Delete(OutputPath) 'deletes existing file on that path
            wordDoc.ExportAsFixedFormat(OutputPath, Word.WdExportFormat.wdExportFormatPDF, False) 'performs the conversion

        Catch ex As Exception
            MsgBox(ex.Message)
            bFailed = True
        Finally

            'Clean up objects
#Disable Warning BC42104 ' Variable is used before it has been assigned a value
            If IsNothing(wordDoc) = False Then
#Enable Warning BC42104 ' Variable is used before it has been assigned a value
                wordDoc.Close(False)
                wordDoc = Nothing
            End If

            If IsNothing(wordApp) = False Then
                wordApp.Quit()
                wordApp = Nothing
            End If

        End Try

        If bFailed Then
            Return False
        Else
            Return True
        End If

    End Function

End Class
