<Serializable()>
Public Class FormParameters

    Public Property AddPages As Boolean
    Public Property PagePrefix As String

    Public Property AddHeaders As Boolean
    Public Property HeaderRegExFind As String
    Public Property HeaderRegExReplace As String

    Public Property AddReturnLinks As Boolean

    Public Property MarginOffset As Single

    Public Sub New()
    End Sub

End Class