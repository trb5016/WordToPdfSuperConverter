

Public Class HelperFunctions

    Public Shared Function OpenArgSourceDocument()

        Dim args As String() = Environment.GetCommandLineArgs

        If args.Count > 1 Then

            Return args(1)
        Else
            Return ""
        End If


    End Function

End Class
