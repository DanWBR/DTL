Module ControlExtensions

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToArrayString(vector As Double()) As String

        Dim retstr As String = "{ "
        For Each d As Double In vector
            retstr += d.ToString + ", "
        Next
        retstr.TrimEnd(",")
        retstr += "}"

        Return retstr

    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToArrayString(vector As String()) As String

        Dim retstr As String = "{ "
        For Each s As String In vector
            retstr += s + ", "
        Next
        retstr.TrimEnd(",")
        retstr += "}"

        Return retstr

    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToArrayString(vector As Object()) As String

        Dim retstr As String = "{ "
        For Each d As Object In vector
            retstr += d.ToString + ", "
        Next
        retstr.TrimEnd(",")
        retstr += "}"

        Return retstr

    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToArrayString(vector As Array) As String

        Dim retstr As String = "{ "
        For Each d As Object In vector
            retstr += d.ToString + ", "
        Next
        retstr.TrimEnd(",")
        retstr += "}"

        Return retstr

    End Function

End Module
