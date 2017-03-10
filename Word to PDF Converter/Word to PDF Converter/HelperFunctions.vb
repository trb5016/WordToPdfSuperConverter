

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

End Class
