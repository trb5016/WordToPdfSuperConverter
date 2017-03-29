Imports System.IO

Public Class SuperConverter2

    Public Shared Function ConvertWordLinksToPdfandMergeAll(ByVal WordDocPath As String, ByVal options As SuperConverterForm.FormParameters) As Boolean

        Dim appendList As New List(Of PdfManipulation2.AttachmentFile)
        Dim pdfPath As String = Path.ChangeExtension(WordDocPath, "pdf")

        'Set current directory to directory of source word document (for later use in resolving relative paths)
        Directory.SetCurrentDirectory(Path.GetDirectoryName(WordDocPath))

        'Get list of actions associated with links
        appendList = PdfManipulation2.GetPdfAppendList(WordDocPath)

        'Create output pdf
        Return PdfManipulation2.CreateOuputDocument(pdfPath, appendList, options)

        Return True

    End Function




End Class
