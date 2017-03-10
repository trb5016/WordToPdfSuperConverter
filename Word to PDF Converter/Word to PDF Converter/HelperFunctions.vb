Imports Word = Microsoft.Office.Interop.Word
Imports System.IO

Imports iTextSharp.text
Imports iTextSharp.text.pdf

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

    Public Shared Function GetPdfNumPages(ByVal FilePath As String) As Integer
        Dim R As New PdfReader(FilePath)
        Return R.NumberOfPages
    End Function

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

    Public Shared Function GetPdfAppendList(ByVal SourceDocPath As String) As List(Of AttachmentFile)

        Dim R As New PdfReader(SourceDocPath)
        Dim pageDictionary As PdfDictionary
        Dim annots As PdfArray
        Dim attachmentPath As String
        Dim attachmentFile As AttachmentFile
        Dim tempList As New List(Of AttachmentFile)
        Dim currentInsertPage As Integer = R.NumberOfPages + 1

        'Loop through each page
        For i As Integer = 1 To R.NumberOfPages

            'Gets current page
            pageDictionary = R.GetPageN(i)

            'Get all of the annotations for the current page
            annots = pageDictionary.GetAsArray(PdfName.ANNOTS)

            'Make sure we have something
            If IsNothing(annots) OrElse annots.Length = 0 Then
                Continue For
            End If

            'Loop through each annotation
            For Each A As PdfObject In annots.ArrayList

                'Convert the iText-specific object as a generic pdf object
                Dim annotationDictionary As PdfDictionary = DirectCast(PdfReader.GetPdfObject(A), PdfDictionary)

                'Make sure this annotation has a LINK
                If Not annotationDictionary.Get(PdfName.SUBTYPE).Equals(PdfName.LINK) Then
                    Continue For
                End If

                'Make sure this annotation has an ACTION
                If annotationDictionary.Get(PdfName.A) Is Nothing Then
                    Continue For
                End If

                'Get the ACTION for the current annotation
                Dim annotationAction As PdfDictionary = DirectCast(annotationDictionary.Get(PdfName.A), PdfDictionary)


                'Test if it is a URI action
                If annotationAction.Get(PdfName.S).Equals(PdfName.URI) Then

                    'prep object for list insertion
                    attachmentFile = New AttachmentFile
                    attachmentFile.IsMainDocument = False
                    attachmentFile.LinkPage = i

                    'Get path of attachment to be appended
                    attachmentPath = annotationAction.GetAsString(PdfName.URI).ToString
                    attachmentPath = Uri.UnescapeDataString(Path.GetFullPath(attachmentPath)) 'This should take care of cases where the link is a relative path            

                    'Convert to pdf if needed (also store if we're making a temp pdf file)
                    If Path.GetExtension(attachmentPath) <> ".pdf" Then
                        attachmentPath = ConvertToPDF(attachmentPath)
                        attachmentFile.IsTempFile = True
                    Else
                        attachmentFile.IsTempFile = False
                    End If


                    'we were able to come up with a pdf file - add it to list
                    If attachmentPath <> "" Then

                        attachmentFile.SourcePath = attachmentPath 'set the path

                        'set the eventual insert page (and increment for next one)
                        attachmentFile.AttachmentStartPage = currentInsertPage
                        currentInsertPage = currentInsertPage + GetPdfNumPages(attachmentPath)

                        'add to list
                        tempList.Add(attachmentFile)

                    Else
                        'Do nothing (?)

                    End If

                End If

            Next

        Next

        Return tempList

    End Function

    Public Shared Function MergePdfFiles(ByVal pdfFiles() As String, ByVal outputPath As String) As Boolean
        Dim result As Boolean = False
        Dim pdfCount As Integer = 0     'total input pdf file count
        Dim f As Integer = 0    'pointer to current input pdf file
        Dim fileName As String
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim pageCount As Integer = 0
        Dim pdfDoc As iTextSharp.text.Document = Nothing    'the output pdf document
        Dim writer As PdfWriter = Nothing
        Dim cb As PdfContentByte = Nothing

        Dim page As PdfImportedPage = Nothing
        Dim rotation As Integer = 0

        Try
            pdfCount = pdfFiles.Length
            If pdfCount > 1 Then
                'Open the 1st item in the array PDFFiles
                fileName = pdfFiles(f)
                reader = New iTextSharp.text.pdf.PdfReader(fileName)
                'Get page count
                pageCount = reader.NumberOfPages

                pdfDoc = New iTextSharp.text.Document(reader.GetPageSizeWithRotation(1), 18, 18, 18, 18)

                writer = PdfWriter.GetInstance(pdfDoc, New FileStream(outputPath, FileMode.OpenOrCreate))


                With pdfDoc
                    .Open()
                End With
                'Instantiate a PdfContentByte object
                cb = writer.DirectContent
                'Now loop thru the input pdfs
                While f < pdfCount
                    'Declare a page counter variable
                    Dim i As Integer = 0
                    'Loop thru the current input pdf's pages starting at page 1
                    While i < pageCount
                        i += 1
                        'Get the input page size
                        pdfDoc.SetPageSize(reader.GetPageSizeWithRotation(i))
                        'Create a new page on the output document
                        pdfDoc.NewPage()
                        'If it is the 1st page, we add bookmarks to the page
                        'Now we get the imported page
                        page = writer.GetImportedPage(reader, i)
                        'Read the imported page's rotation
                        rotation = reader.GetPageRotation(i)
                        'Then add the imported page to the PdfContentByte object as a template based on the page's rotation
                        If rotation = 90 Then
                            cb.AddTemplate(page, 0, -1.0F, 1.0F, 0, 0, reader.GetPageSizeWithRotation(i).Height)
                        ElseIf rotation = 270 Then
                            cb.AddTemplate(page, 0, 1.0F, -1.0F, 0, reader.GetPageSizeWithRotation(i).Width + 60, -30)
                        Else
                            cb.AddTemplate(page, 1.0F, 0, 0, 1.0F, 0, 0)
                        End If
                    End While
                    'Increment f and read the next input pdf file
                    f += 1
                    If f < pdfCount Then
                        fileName = pdfFiles(f)
                        reader = New iTextSharp.text.pdf.PdfReader(fileName)
                        pageCount = reader.NumberOfPages
                    End If
                End While
                'When all done, we close the document so that the pdfwriter object can write it to the output file
                pdfDoc.Close()
                result = True
            End If
        Catch ex As Exception
            Return False
        End Try
        Return result
    End Function

    Public Shared Function MergePdfsWithLinks(ByVal pdfPaths As List(Of String), ByVal OutputPath As String) As Boolean

        Dim R As PdfReader

        'Next we create a new document add import each page from the reader above
        Using FS As New FileStream(OutputPath, FileMode.Create, FileAccess.Write, FileShare.None)
            Using Doc As New Document()
                Using writer As New PdfCopy(Doc, FS)
                    Doc.Open()


                    For Each file As String In pdfPaths

                        R = New PdfReader(file)

                        For i As Integer = 1 To R.NumberOfPages
                            writer.AddPage(writer.GetImportedPage(R, i))
                        Next

                    Next

                    Doc.Close()
                End Using
            End Using
        End Using

        Return True

    End Function

    Public Class AttachmentFile

        Public SourcePath As String
        Public IsTempFile As Boolean
        Public IsMainDocument As Boolean
        Public LinkPage As Integer
        Public AttachmentStartPage As Integer

    End Class

End Class
