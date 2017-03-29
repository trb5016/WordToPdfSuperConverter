Imports System.IO

Public Class SuperConverter

    Public Shared Function ConvertWordLinksToPdfandMergeAll(ByVal WordDocPath As String, ByVal options As SuperConverterForm.FormParameters) As Boolean

        Dim appendList As New List(Of PdfManipulation.AttachmentFile)
        Dim appendFileList As List(Of String)

        Dim tempSourcePdf As String = Path.GetTempFileName
        Dim tempIntermediatePdf As String = Path.GetTempFileName
        Dim tempIntermediatePdf2 As String = Path.GetTempFileName

        Dim pdfPath As String = Path.ChangeExtension(WordDocPath, "pdf")
        Call Converters.ConvertWordToPdf(WordDocPath, tempSourcePdf)

        'Set current directory to directory of source word document (for later use in resolving relative paths)
        Directory.SetCurrentDirectory(Path.GetDirectoryName(WordDocPath))

        'Get list of actions associated with links
        appendList = PdfManipulation.GetPdfAppendList(tempSourcePdf)

        'Adds source file pdf to start of list of documents to be merged
        Dim mainDoc As New PdfManipulation.AttachmentFile
        mainDoc.OriginalSourcePath = WordDocPath
        mainDoc.CurrentSourcePath = tempSourcePdf
        mainDoc.IsMainDocument = True
        appendList.Insert(0, mainDoc)

        'make list of just the source filenames
        appendFileList = appendList.Select(Function(a) a.CurrentSourcePath).ToList

        MergePdfFiles(appendFileList.ToArray, "C:\Users\u001tb7\Desktop\test.pdf")

        MsgBox("Halt")
        Return True

        'Merge source file and attachments into one pdf
        Call PdfManipulation.MergePdfsWithLinks(appendFileList, tempIntermediatePdf)

        'Create PDF with updated links
        Call PdfManipulation.UpdateLinks(tempIntermediatePdf, appendList, tempIntermediatePdf2)

        'Create final PDF with return links
        File.Delete(pdfPath)
        Call PdfManipulation.AddExtraText(tempIntermediatePdf2, appendList, pdfPath, options)

        'Delete temp files
        File.Delete(tempSourcePdf)
        File.Delete(tempIntermediatePdf)
        File.Delete(tempIntermediatePdf2)

        'Delete temp attachment files
        For Each tempFile As PdfManipulation.AttachmentFile In appendList.Where(Function(a) a.IsTempFile = True).ToList
            File.Delete(tempFile.CurrentSourcePath)
        Next

        Return True

    End Function




End Class
