Imports System.IO

Imports iTextSharp.text
Imports iTextSharp.text.pdf

Public Class PdfManipulation2

    Public Shared Function GetPdfNumPages(ByVal FilePath As String) As Integer
        Using R As New PdfReader(FilePath)
            Dim numPages As Integer = R.NumberOfPages
            R.Close()
            Return numPages
        End Using
    End Function

    Public Shared Function GetPdfAppendList(ByVal SourceWordDocPath As String) As List(Of AttachmentFile)

        Dim R As PdfReader
        Dim pageDictionary As PdfDictionary
        Dim annots As PdfArray
        Dim attachmentFile As AttachmentFile
        Dim tempList As New List(Of AttachmentFile)
        Dim currentInsertPage As Integer
        Dim attachmentCounter As Integer = 1

        'Convert source word document to pdf and add to append list as first entry
        attachmentFile = New AttachmentFile
        With attachmentFile
            .OriginalSourcePath = SourceWordDocPath
            .CurrentSourcePath = Converters.ConvertToPDF(attachmentFile.OriginalSourcePath)
            .IsMainDocument = True
            .IsTempFile = True
        End With
        tempList.Add(attachmentFile)

        'Open reader for the pdf of the source document (and set currentInsertPage)
        R = New PdfReader(attachmentFile.CurrentSourcePath)
        currentInsertPage = R.NumberOfPages + 1

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

    Public Shared Function CreateOuputDocument(ByVal outputPath As String, ByVal appendList As List(Of AttachmentFile), ByVal options As SuperConverterForm.FormParameters)

        Dim pdfDoc As Document = Nothing
        Dim reader As PdfReader = Nothing
        Dim writer As PdfWriter = Nothing
        Dim cb As PdfContentByte = Nothing

        If appendList.Count > 1 Then

            'Prep the pdf document
            reader = New PdfReader(appendList(0).CurrentSourcePath) 'Open the first file (which is the primary document) for use
            pdfDoc = New Document()
            writer = PdfWriter.GetInstance(pdfDoc, New FileStream(outputPath, FileMode.Create))

            'Set metadata and open document
            With pdfDoc
                .AddAuthor(Environment.UserName)
                .AddCreationDate()
                .AddCreator("Created with Agenda Creator")
                .AddTitle(Path.GetFileNameWithoutExtension(appendList(0).OriginalSourcePath))
                .Open()
            End With

            'Instantiate PdfContentByte object
            cb = writer.DirectContent





        End If


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
