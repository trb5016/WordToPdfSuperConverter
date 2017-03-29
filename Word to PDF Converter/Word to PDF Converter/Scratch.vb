﻿Imports iTextSharp.text.pdf

Module Scratch

    Public Function MergePdfFiles(ByVal pdfFiles() As String, ByVal outputPath As String) As Boolean
        Dim result As Boolean = False
        Dim pdfCount As Integer = 0     'total input pdf file count
        Dim f As Integer = 0            'pointer to current input pdf file
        Dim fileName As String = String.Empty   'current input pdf filename
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim pageCount As Integer = 0    'cureent input pdf page count
        Dim pdfDoc As iTextSharp.text.Document = Nothing    'the output pdf document
        Dim writer As PdfWriter = Nothing
        Dim cb As PdfContentByte = Nothing
        'Declare a variable to hold the imported pages
        Dim page As PdfImportedPage = Nothing
        Dim rotation As Integer = 0
        'Declare a font to used for the bookmarks
        Dim bookmarkFont As iTextSharp.text.Font = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.HELVETICA,
                                                                  12, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLUE)
        Try
            pdfCount = pdfFiles.Length
            If pdfCount > 1 Then
                'Open the 1st pad using PdfReader object
                fileName = pdfFiles(f)
                reader = New iTextSharp.text.pdf.PdfReader(fileName)
                'Get page count
                pageCount = reader.NumberOfPages
                'Instantiate an new instance of pdf document and set its margins. This will be the output pdf.
                'NOTE: bookmarks will be added at the 1st page of very original pdf file using its filename. The location
                'of this bookmark will be placed at the upper left hand corner of the document. So you'll need to adjust
                'the margin left and margin top values such that the bookmark won't overlay on the merged pdf page. The 
                'unit used is "points" (72 points = 1 inch), thus in this example, the bookmarks' location is at 1/4 inch from
                'left and 1/4 inch from top of the page.
                pdfDoc = New iTextSharp.text.Document(reader.GetPageSizeWithRotation(1), 18, 18, 18, 18)
                'Instantiate a PdfWriter that listens to the pdf document
                writer = PdfWriter.GetInstance(pdfDoc, New IO.FileStream(outputPath, IO.FileMode.Create))
                'Set metadata and open the document
                With pdfDoc
                    .AddAuthor("Your name here")
                    .AddCreationDate()
                    .AddCreator("Your program name here")
                    .AddSubject("Whatever subject you want to give it")
                    'Use the filename as the title... You can give it any title of course.
                    .AddTitle(IO.Path.GetFileNameWithoutExtension(outputPath))
                    'Add keywords, whatever keywords you want to attach to it
                    .AddKeywords("Report, Merged PDF, " & IO.Path.GetFileName(outputPath))
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
                        If i = 1 Then
                            'First create a paragraph using the filename as the heading
                            Dim para As New iTextSharp.text.Paragraph(IO.Path.GetFileName(fileName).ToUpper(), bookmarkFont)
                            'Then create a chapter from the above paragraph
                            Dim chpter As New iTextSharp.text.Chapter(para, f + 1)
                            'Finally add the chapter to the document
                            pdfDoc.Add(chpter)
                        End If
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
                'When all done, we close the documwent so that the pdfwriter object can write it to the output file
                pdfDoc.Close()
                result = True
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return result
    End Function

End Module
