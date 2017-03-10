Imports System.IO

Imports iTextSharp.text
Imports iTextSharp.text.pdf

Public Class PdfManipulation

    Public Shared Function GetPdfNumPages(ByVal FilePath As String) As Integer
        Dim R As New PdfReader(FilePath)
        Return R.NumberOfPages
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
                        attachmentPath = Converters.ConvertToPDF(attachmentPath)
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

        R.Close()

        Return tempList

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

        R.Close()
        Return True

    End Function

    Public Shared Function UpdateLinks(ByVal pdfSourcePath As String, ByVal appendList As List(Of AttachmentFile), ByVal outputPath As String) As Boolean

        Dim R As New PdfReader(pdfSourcePath)
        Dim pageDictionary As PdfDictionary
        Dim annots As PdfArray

        Dim mainDocNumPages As Integer = GetPdfNumPages(appendList(0).SourcePath) 'The first item in the append list is always the main document


        Using FS As New FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None)
            Using Doc As New Document()
                Using writer As New PdfCopy(Doc, FS)
                    Doc.Open()

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

                            Dim l As New Chunk()

                            'Get the ACTION for the current annotation
                            Dim annotationAction As PdfDictionary = DirectCast(annotationDictionary.Get(PdfName.A), PdfDictionary)

                            If annotationAction.Get(PdfName.S).Equals(PdfName.URI) Then

                                'Remove the old action, I don't think this is actually necessary but I do it anyways
                                annotationAction.Remove(PdfName.S)

                                'Add a new action that is a GOTO action
                                annotationAction.Put(PdfName.S, PdfName.GOTO)

                                'The destination is an array containing an indirect reference to the page as well as a fitting option
                                Dim NewLocalDestination As New PdfArray()

                                'Link it to page 2
                                NewLocalDestination.Add(R.GetPageOrigRef(2))

                                'Set it to fit page
                                NewLocalDestination.Add(PdfName.FIT)

                                'Add the array to the annotation's destination (/D)
                                annotationAction.Put(PdfName.D, NewLocalDestination)

                            End If

                            Exit For

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
