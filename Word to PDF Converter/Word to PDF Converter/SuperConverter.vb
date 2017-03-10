Imports System.IO

Public Class SuperConverter

    Public Shared Function ConvertWordLinksToPdfandMergeAll(ByVal WordDocPath As String) As Boolean

        Dim appendList As New List(Of PdfManipulation.AttachmentFile)
        Dim appendFileList As List(Of String)

        Dim tempSourcePdf As String = Path.GetTempFileName
        Dim tempIntermediatePdf As String = Path.GetTempFileName
        Dim pdfPath As String = Path.ChangeExtension(WordDocPath, "pdf")
        Call Converters.ConvertWordToPdf(WordDocPath, tempSourcePdf)

        'Set current directory to directory of source word document (for later use in resolving relative paths)
        Directory.SetCurrentDirectory(Path.GetDirectoryName(WordDocPath))

        'Get list of actions associated with links
        appendList = PdfManipulation.GetPdfAppendList(tempSourcePdf)

        'Adds source file pdf to start of list of documents to be merged
        Dim mainDoc As New PdfManipulation.AttachmentFile
        mainDoc.SourcePath = tempSourcePdf
        mainDoc.IsMainDocument = True
        appendList.Insert(0, mainDoc)

        'make list of just the source filenames
        appendFileList = appendList.Select(Function(a) a.SourcePath).ToList

        'Merge source file and attachments into one pdf

        Call PdfManipulation.MergePdfsWithLinks(appendFileList, tempIntermediatePdf)

        'Create final PDF with updated links
        File.Delete(pdfPath)
        Call PdfManipulation.UpdateLinks(tempIntermediatePdf, appendList, pdfPath)

        'Delete temp files
        'File.Delete(tempSourcePdf)
        File.Delete(tempIntermediatePdf)

        Return True

    End Function




End Class
