Imports Word = Microsoft.Office.Interop.Word
Imports System.IO

Public Class Converters

    Public Shared Function ConvertToPDF(ByVal FilePath As String) As String

        Select Case Path.GetExtension(FilePath)
            Case ".pdf"
                Return FilePath

            Case ".doc", ".docx", ".docm"

                Dim tempPath As String = Path.GetTempFileName

                If ConvertWordToPdf(FilePath, tempPath) = True Then
                    Return tempPath
                Else
                    MsgBox("Failed to convert file to PDF, file will not be appended: " & vbCrLf & vbCrLf & FilePath)
                    Return ""

                End If

            Case Else
                MsgBox("Could not convert " & FilePath & " to a pdf file.  (No method written)" & vbCrLf & vbCrLf & "File will not be appended.")
                Return ""

        End Select

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
