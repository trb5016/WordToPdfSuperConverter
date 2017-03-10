Imports System.IO

Public Class Conversion

    Public Shared Function ConvertWordLinksToPdfandMergeAll(ByVal WordDocPath As String) As Boolean

        Dim appendList As New List(Of HelperFunctions.AttachmentFile)
        Dim appendFileList As List(Of String)

        Dim tempPdf As String = Path.GetTempFileName
        Dim pdfPath As String = Path.ChangeExtension(WordDocPath, "pdf")
        Call HelperFunctions.ConvertWordToPdf(WordDocPath, tempPdf)

        'Set current directory to directory of source word document (for later use in resolving relative paths)
        Directory.SetCurrentDirectory(Path.GetDirectoryName(WordDocPath))

        'Get list of actions associated with links
        appendList = HelperFunctions.GetPdfAppendList(tempPdf)

        'Adds source file pdf to start of list of documents to be merged
        Dim mainDoc As New HelperFunctions.AttachmentFile
        mainDoc.SourcePath = tempPdf
        mainDoc.IsMainDocument = True
        appendList.Insert(0, mainDoc)

        'make list of just the source filenames
        appendFileList = appendList.Select(Function(a) a.SourcePath).ToList

        'Merge source file and attachments into one pdf
        File.Delete(pdfPath)
        Call HelperFunctions.MergePdfsWithLinks(appendFileList, pdfPath)

        Return True

    End Function




End Class
