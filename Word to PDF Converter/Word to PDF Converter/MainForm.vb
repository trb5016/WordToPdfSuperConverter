Imports System.IO

Imports iTextSharp.text
Imports iTextSharp.text.pdf


Public Class Form1


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        Dim sourceDocPath As String = HelperFunctions.SourceDocument
        Dim tempPdf As String = Path.GetTempFileName
        Dim pdfPath As String = Path.ChangeExtension(sourceDocPath, "pdf")

        HelperFunctions.ConvertWordToPdf(sourceDocPath, tempPdf)

        Dim R As PdfReader
        Dim pageCount As Integer = 0
        Dim pageDictionary As PdfDictionary
        Dim annots As PdfArray

        'Open reader
        R = New PdfReader(tempPdf)

        'Get page count
        pageCount = R.NumberOfPages

        'Loop through each page
        For i As Integer = 1 To pageCount

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

                    'show the link
                    Dim link As PdfString = annotationAction.GetAsString(PdfName.URI)

                    'change the uri to something else
                    annotationAction.Put(PdfName.URI, New PdfString("http://www.google.com"))

                    link = annotationAction.GetAsString(PdfName.URI)

                End If

            Next

        Next

        'Create new document add import each page from the reader above
        File.Delete(pdfPath)
        'Next we create a new document add import each page from the reader above
        Using FS As New FileStream(pdfPath, FileMode.Create, FileAccess.Write, FileShare.None)
            Using Doc As New Document()
                Using writer As New PdfCopy(Doc, FS)
                    Doc.Open()
                    For i As Integer = 1 To R.NumberOfPages
                        writer.AddPage(writer.GetImportedPage(R, i))
                    Next
                    Doc.Close()
                End Using
            End Using
        End Using



    End Sub


End Class
