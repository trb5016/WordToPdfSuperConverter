Imports System.IO

Imports iTextSharp.text
Imports iTextSharp.text.pdf

Public Class PdfManipulation

    Public Shared Function GetPdfNumPages(ByVal FilePath As String) As Integer
        Using R As New PdfReader(FilePath)
            Dim numPages As Integer = R.NumberOfPages
            R.Close()
            Return numPages
        End Using
    End Function

    Public Shared Function GetPdfAppendList(ByVal SourceDocPath As String) As List(Of AttachmentFile)

        Dim R As New PdfReader(SourceDocPath)
        Dim pageDictionary As PdfDictionary
        Dim annots As PdfArray
        Dim attachmentFile As AttachmentFile
        Dim tempList As New List(Of AttachmentFile)
        Dim currentInsertPage As Integer = R.NumberOfPages + 1
        Dim attachmentCounter As Integer = 1

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
                    attachmentFile.OriginalSourcePath = annotationAction.GetAsString(PdfName.URI).ToString
                    attachmentFile.OriginalSourcePath = Uri.UnescapeDataString(Path.GetFullPath(attachmentFile.OriginalSourcePath)) 'This should take care of cases where the link is a relative path            

                    'Convert to pdf if needed (also store if we're making a temp pdf file)
                    If Path.GetExtension(attachmentFile.OriginalSourcePath) <> ".pdf" Then

                        attachmentFile.CurrentSourcePath = Converters.ConvertToPDF(attachmentFile.OriginalSourcePath)
                        attachmentFile.IsTempFile = True

                    Else
                        attachmentFile.IsTempFile = False
                        attachmentFile.CurrentSourcePath = attachmentFile.OriginalSourcePath
                    End If


                    'we were able to come up with a pdf file - add it to list
                    If attachmentFile.CurrentSourcePath <> "" Then

                        'set the eventual insert page (and increment for next one)
                        attachmentFile.AttachmentStartPage = currentInsertPage
                        attachmentFile.AttachmentEndPage = currentInsertPage + GetPdfNumPages(attachmentFile.CurrentSourcePath) - 1
                        currentInsertPage = attachmentFile.AttachmentEndPage + 1

                        'Set the attachment number
                        attachmentFile.AttachmentNumber = attachmentCounter
                        attachmentCounter = attachmentCounter + 1

                        'add to list
                        tempList.Add(attachmentFile)

                    Else
                        'Do nothing (?)

                    End If



                End If

            Next

        Next

        R.Close()
        R = Nothing

        Return tempList

    End Function

    Public Shared Function MergePdfsWithLinks(ByVal pdfPaths As List(Of String), ByVal OutputPath As String) As Boolean

        'Next we create a new document add import each page from the reader above
        Using FS As New FileStream(OutputPath, FileMode.Create, FileAccess.Write, FileShare.None)
            Using Doc As New Document()
                Using writer As New PdfCopy(Doc, FS)
                    Doc.Open()

                    For Each file As String In pdfPaths

                        Dim R As New PdfReader(file)

                        For i As Integer = 1 To R.NumberOfPages
                            writer.AddPage(writer.GetImportedPage(R, i))
                        Next

                        R.Close()
                        R = Nothing
                    Next

                    Doc.Close()
                End Using
            End Using
        End Using

        Return True

    End Function

    Public Shared Function UpdateLinks(ByVal pdfSourcePath As String, ByVal appendList As List(Of AttachmentFile), ByVal outputPath As String) As Boolean


        Dim pageDictionary As PdfDictionary
        Dim annots As PdfArray
        Dim curLink As Integer 'used for tracking which link we're working on in the appendList

        Dim mainDocNumPages As Integer = GetPdfNumPages(appendList(0).CurrentSourcePath) 'The first item in the append list is always the main document

        Using R As New PdfReader(pdfSourcePath)
            Using FS As New FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None)
                Using Doc As New Document()
                    Using writer As New PdfCopy(Doc, FS)
                        Doc.Open()

                        curLink = 1 'this also skips the first entry in appendList which is the main doc

                        'Loop through each page
                        For i As Integer = 1 To mainDocNumPages

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

                                If annotationAction.Get(PdfName.S).Equals(PdfName.URI) Then

                                    'Remove the old action, I don't think this is actually necessary but I do it anyways
                                    annotationAction.Remove(PdfName.S)

                                    'Add a new action that is a GOTO action
                                    annotationAction.Put(PdfName.S, PdfName.GOTO)

                                    'The destination is an array containing an indirect reference to the page as well as a fitting option
                                    Dim NewLocalDestination As New PdfArray()

                                    'Link it to appropriate page
                                    NewLocalDestination.Add(R.GetPageOrigRef(appendList(curLink).AttachmentStartPage))

                                    'Set it to fit page
                                    NewLocalDestination.Add(PdfName.FIT)

                                    'Add the array to the annotation's destination (/D)
                                    annotationAction.Put(PdfName.D, NewLocalDestination)

                                    curLink = curLink + 1

                                End If

                            Next

                        Next

                        'Writes out final PDF
                        For i As Integer = 1 To R.NumberOfPages

                            writer.AddPage(writer.GetImportedPage(R, i))

                        Next

                        Doc.Close()

                    End Using
                End Using
            End Using
            R.Close()
        End Using




        Return True

    End Function

    Public Shared Function AddExtraText(ByVal pdfSourcePath As String, ByVal appendList As List(Of AttachmentFile), ByVal outputPath As String, ByVal options As SuperConverterForm.FormParameters) As Boolean

        Dim margin As Single = options.MarginOffset
        Dim bytes As Byte() = File.ReadAllBytes(pdfSourcePath)

        Using stream As New MemoryStream()

            Dim reader As New PdfReader(bytes)

            Using stamper As New PdfStamper(reader, stream)

                Dim numPages As Integer = reader.NumberOfPages

                For i As Integer = 1 To numPages

                    Dim pageSize As Rectangle = reader.GetPageSizeWithRotation(i)
                    Dim pageCanvas As PdfContentByte = stamper.GetOverContent(i)
                    Dim curAttachment As AttachmentFile = GetCurrentAttachment(appendList, i)

                    'Add page Number (main doc and attachments)
                    If options.AddPages = True Then


                        'TODO: The addition of the backing box is hacky and doesn't actually adjust based on the text size...

                        Dim pageFont As Font = FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.GRAY)
                        Dim fontHeight As Single = pageFont.BaseFont.GetFontDescriptor(BaseFont.ASCENT, pageFont.Size) - pageFont.BaseFont.GetFontDescriptor(BaseFont.DESCENT, pageFont.Size)

                        Dim pageNumRectangle As New Rectangle(pageSize.Width / 2, pageSize.GetBottom(margin) - 2, pageSize.GetRight(margin), pageSize.GetBottom(margin) + fontHeight)
                        pageNumRectangle.BackgroundColor = BaseColor.WHITE


                        pageCanvas.Rectangle(pageNumRectangle)
                        pageCanvas.Stroke()


                        Dim pageNumber As New Phrase(options.PagePrefix & i & " of " & numPages, pageFont)
                        ColumnText.ShowTextAligned(pageCanvas, Element.ALIGN_RIGHT, pageNumber, pageSize.GetRight(margin), pageSize.GetBottom(margin), 0)

                    End If

                    'Working with an attachment page
                    If Not IsNothing(curAttachment) Then

                        'Add Headers
                        If options.AddHeaders = True Then

                            Dim headerFont As Font = FontFactory.GetFont(FontFactory.HELVETICA, 14, Font.BOLD, BaseColor.GRAY)
                            Dim headerText As New Phrase("Attachment " & curAttachment.AttachmentNumber & ": " & curAttachment.GetAttachmentName(options.HeaderRegExFind, options.HeaderRegExReplace), headerFont)
                            ColumnText.ShowTextAligned(pageCanvas, Element.ALIGN_LEFT, headerText, pageSize.GetLeft(margin), pageSize.GetTop(margin) - margin, 0)

                        End If

                        'Add return links
                        If options.AddReturnLinks = True Then

                            Dim returnFont As Font = FontFactory.GetFont(FontFactory.HELVETICA, 12, Font.UNDERLINE, BaseColor.BLUE)
                            Dim returnTextChunk As New Chunk("RETURN", returnFont)
                            Dim action As PdfAction = PdfAction.GotoLocalPage(curAttachment.LinkPage, New PdfDestination(PdfDestination.FIT), stamper.Writer)

                            returnTextChunk.SetAction(action)

                            Dim returnTextPhrase As New Phrase(returnTextChunk)

                            ColumnText.ShowTextAligned(pageCanvas, Element.ALIGN_LEFT, returnTextPhrase, pageSize.GetLeft(margin), pageSize.GetBottom(margin), 0)

                        End If

                    End If

                Next

            End Using

            bytes = stream.ToArray()
        End Using

        File.WriteAllBytes(outputPath, bytes)

        Return True

    End Function

    Private Shared Function GetCurrentAttachment(ByVal appendList As List(Of AttachmentFile), ByVal currentPage As Integer) As AttachmentFile

        Dim releventAttachment As List(Of AttachmentFile) = appendList.Where(Function(a) currentPage >= a.AttachmentStartPage And currentPage <= a.AttachmentEndPage).ToList


        Select Case releventAttachment.Count
            Case 0
                'do nothing
                Return Nothing
            Case 1
                Return releventAttachment.First
            Case Else
                Throw New Exception("Something went wrong in trying to add the return links for Page " & currentPage)

        End Select

    End Function

    Public Class AttachmentFile

        Public OriginalSourcePath As String
        Public CurrentSourcePath As String
        Public IsTempFile As Boolean
        Public IsMainDocument As Boolean

        Public LinkPage As Integer

        Public AttachmentStartPage As Integer
        Public AttachmentEndPage As Integer
        Public AttachmentNumber As Integer

        Public Function GetAttachmentName(ByVal RegExFind As String, ByVal RegExReplace As String) As String

            Dim filename As String = Path.GetFileNameWithoutExtension(OriginalSourcePath)
            Dim regEx As New Text.RegularExpressions.Regex(RegExFind)

            Return regEx.Replace(filename, RegExReplace)

        End Function

    End Class

End Class
