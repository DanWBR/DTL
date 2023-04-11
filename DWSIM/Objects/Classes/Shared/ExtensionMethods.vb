Public Module Extensions

    <Runtime.CompilerServices.Extension()>
    Public Function ToArrayString(vector As Double()) As String

        Dim retstr As String = "{ "
        For Each d In vector
            retstr += d.ToString + ", "
        Next
        retstr.TrimEnd(",")
        retstr += "}"

        Return retstr

    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function ToArrayString(vector As String()) As String

        Dim retstr As String = "{ "
        For Each s In vector
            retstr += s + ", "
        Next
        retstr.TrimEnd(",")
        retstr += "}"

        Return retstr

    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function ToArrayString(vector As Object()) As String

        Dim retstr As String = "{ "
        For Each d In vector
            retstr += d.ToString + ", "
        Next
        retstr.TrimEnd(",")
        retstr += "}"

        Return retstr

    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function ToArrayString(vector As Array) As String

        Dim retstr As String = "{ "
        For Each d In vector
            retstr += d.ToString + ", "
        Next
        retstr.TrimEnd(",")
        retstr += "}"

        Return retstr

    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function IsValid(d As Double) As Boolean
        If Double.IsNaN(d) Or Double.IsInfinity(d) Then Return False Else Return True
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function IsValid(d As Double?) As Boolean
        If Double.IsNaN(d.GetValueOrDefault) Or Double.IsInfinity(d.GetValueOrDefault) Then Return False Else Return True
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function IsPositive(d As Double) As Boolean
        If d.IsValid() Then
            If d > 0.0# Then Return True Else Return False
        Else
            Throw New ArgumentException("invalid double")
        End If
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function IsPositive(d As Double?) As Boolean
        If d.GetValueOrDefault.IsValid() Then
            If d.GetValueOrDefault > 0.0# Then Return True Else Return False
        Else
            Throw New ArgumentException("invalid double")
        End If
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function IsNegative(d As Double) As Boolean
        If d.IsValid() Then
            If d < 0.0# Then Return True Else Return False
        Else
            Throw New ArgumentException("invalid double")
        End If
    End Function

    <Runtime.CompilerServices.Extension()>
    Public Function IsNegative(d As Double?) As Boolean
        If d.GetValueOrDefault.IsValid() Then
            If d.GetValueOrDefault < 0.0# Then Return True Else Return False
        Else
            Throw New ArgumentException("invalid double")
        End If
    End Function

    ''' <summary>
    ''' Alternative implementation for the Exponential (Exp) function.
    ''' </summary>
    ''' <param name="val"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()> Public Function ExpY(val As Double) As Double
        Dim tmp As Long = 1512775 * val + 1072632447
        Return BitConverter.Int64BitsToDouble(tmp << 32)
    End Function

#Region "SIMD Functions"

    ''' <summary>
    ''' Computes the exponential of each vector element.
    ''' </summary>
    ''' <param name="vector"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()> Public Function ExpY(vector As Double()) As Double()

        Dim vector2(vector.Length - 1) As Double

        If My.MyApplication.UseSIMDExtensions Then
            'Yeppp.Math.Exp_V64f_V64f(vector, 0, vector2, 0, vector.Length)
        Else
            For i As Integer = 0 To vector.Length - 1
                vector2(i) = Math.Exp(vector(i))
            Next
        End If

        Return vector2

    End Function

    ''' <summary>
    ''' Computes the natural logarithm of each vector element.
    ''' </summary>
    ''' <param name="vector"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()> Public Function LogY(vector As Double()) As Double()

        Dim vector2(vector.Length - 1) As Double

        If My.MyApplication.UseSIMDExtensions Then
            'Yeppp.Math.Log_V64f_V64f(vector, 0, vector2, 0, vector.Length)
        Else
            For i As Integer = 0 To vector.Length - 1
                vector2(i) = Math.Log(vector(i))
            Next
        End If

        Return vector2

    End Function

    ''' <summary>
    ''' Returns the smallest element in the vector.
    ''' </summary>
    ''' <param name="vector"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()> Public Function MinY(vector As Double()) As Double

        If My.MyApplication.UseSIMDExtensions Then
            'Return Yeppp.Core.Min_V64f_S64f(vector, 0, vector.Length)
        Else
            Return Double.Parse(DTL.MathEx.Common.Min(vector))
        End If

    End Function

    ''' <summary>
    ''' Returns the biggest element in the vector.
    ''' </summary>
    ''' <param name="vector"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()> Public Function MaxY(vector As Double()) As Double

        If My.MyApplication.UseSIMDExtensions Then
            'Return Yeppp.Core.Max_V64f_S64f(vector, 0, vector.Length)
        Else
            Return Double.Parse(DTL.MathEx.Common.Max(vector))
        End If

    End Function

    ''' <summary>
    ''' Sum of the vector elements.
    ''' </summary>
    ''' <param name="vector"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()> Public Function SumY(vector As Double()) As Double

        If My.MyApplication.UseSIMDExtensions Then
            'Return Yeppp.Core.Sum_V64f_S64f(vector, 0, vector.Length)
        Else
            Return DTL.MathEx.Common.Sum(vector)
        End If

    End Function

    ''' <summary>
    ''' Absolute sum of the vector elements
    ''' </summary>
    ''' <param name="vector"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()> Public Function AbsSumY(vector As Double()) As Double

        If My.MyApplication.UseSIMDExtensions Then
            'Return Yeppp.Core.SumAbs_V64f_S64f(vector, 0, vector.Length)
        Else
            Return DTL.MathEx.Common.AbsSum(vector)
        End If

    End Function

    ''' <summary>
    ''' Absolute square sum of vector elements.
    ''' </summary>
    ''' <param name="vector"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()> Public Function AbsSqrSumY(vector As Double()) As Double

        If My.MyApplication.UseSIMDExtensions Then
            'Return Yeppp.Core.SumSquares_V64f_S64f(vector, 0, vector.Length)
        Else
            Return DTL.MathEx.Common.SumSqr(vector)
        End If

    End Function

    ''' <summary>
    ''' Negates the elements of a vector.
    ''' </summary>
    ''' <param name="vector"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()> Public Function NegateY(vector As Double()) As Double()

        Dim vector0(vector.Length - 1) As Double

        If My.MyApplication.UseSIMDExtensions Then
            'Yeppp.Core.Negate_V64f_V64f(vector, 0, vector0, 0, vector.Length)
        Else
            For i As Integer = 0 To vector.Length - 1
                vector0(i) = -vector(i)
            Next
        End If

        Return vector0

    End Function

    ''' <summary>
    ''' Multiplies vector elements.
    ''' </summary>
    ''' <param name="vector"></param>
    ''' <param name="vector2"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()> Public Function MultiplyY(vector As Double(), vector2 As Double()) As Double()

        Dim vector0(vector.Length - 1) As Double

        If My.MyApplication.UseSIMDExtensions Then
            'Yeppp.Core.Multiply_V64fV64f_V64f(vector, 0, vector2, 0, vector0, 0, vector.Length)
        Else
            For i As Integer = 0 To vector.Length - 1
                vector0(i) = vector(i) * vector2(i)
            Next
        End If

        Return vector0

    End Function

    ''' <summary>
    ''' Divides vector elements.
    ''' </summary>
    ''' <param name="vector"></param>
    ''' <param name="vector2"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()> Public Function DivideY(vector As Double(), vector2 As Double()) As Double()

        Dim vector0(vector.Length - 1) As Double

        If My.MyApplication.UseSIMDExtensions Then
            Dim invvector2(vector.Length - 1) As Double
            For i As Integer = 0 To vector2.Length - 1
                invvector2(i) = 1 / vector2(i)
            Next
            'Yeppp.Core.Multiply_V64fV64f_V64f(vector, 0, invvector2, 0, vector0, 0, vector.Length)
        Else
            For i As Integer = 0 To vector.Length - 1
                vector0(i) = vector(i) / vector2(i)
            Next
        End If


        Return vector0

    End Function

    ''' <summary>
    ''' Subtracts vector elements.
    ''' </summary>
    ''' <param name="vector"></param>
    ''' <param name="vector2"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()> Public Function SubtractY(vector As Double(), vector2 As Double()) As Double()

        Dim vector0(vector.Length - 1) As Double

        If My.MyApplication.UseSIMDExtensions Then
            'Yeppp.Core.Subtract_V64fV64f_V64f(vector, 0, vector2, 0, vector0, 0, vector.Length)
        Else
            For i As Integer = 0 To vector.Length - 1
                vector0(i) = vector(i) - vector2(i)
            Next
        End If

        Return vector0

    End Function

    <Runtime.CompilerServices.Extension()> Public Function SubtractInversesY(vector As Double(), vector2 As Double()) As Double()

        Dim vector0(vector.Length - 1) As Double
        Dim invvector1(vector.Length - 1), invvector2(vector.Length - 1) As Double

        For i As Integer = 0 To vector.Length - 1
            invvector1(i) = 1 / vector(i)
            invvector2(i) = 1 / vector2(i)
        Next

        If My.MyApplication.UseSIMDExtensions Then
            'Yeppp.Core.Subtract_V64fV64f_V64f(invvector1, 0, invvector2, 0, vector0, 0, vector.Length)
        Else
            For i As Integer = 0 To vector.Length - 1
                vector0(i) = invvector1(i) - invvector2(i)
            Next
        End If

        Return vector0

    End Function

    <Runtime.CompilerServices.Extension()> Public Function SubtractInverseY(vector As Double(), vector2 As Double()) As Double()

        Dim vector0(vector.Length - 1) As Double

        If My.MyApplication.UseSIMDExtensions Then
            Dim invvector2(vector.Length - 1) As Double
            For i As Integer = 0 To vector.Length - 1
                invvector2(i) = 1 / vector2(i)
            Next
            'Yeppp.Core.Subtract_V64fV64f_V64f(vector, 0, invvector2, 0, vector0, 0, vector.Length)
        Else
            For i As Integer = 0 To vector.Length - 1
                vector0(i) = vector(i) - 1 / vector2(i)
            Next
        End If

        Return vector0

    End Function

    ''' <summary>
    ''' Multiplies vector elements by a constant.
    ''' </summary>
    ''' <param name="vector"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()> Public Function MultiplyConstY(vector As Double(), constant As Double) As Double()

        Dim vector0(vector.Length - 1) As Double

        If My.MyApplication.UseSIMDExtensions Then
            'Yeppp.Core.Multiply_V64fS64f_V64f(vector, 0, constant, vector0, 0, vector.Length)
        Else
            For i As Integer = 0 To vector.Length - 1
                vector0(i) = vector(i) * constant
            Next
        End If


        Return vector0

    End Function

    ''' <summary>
    ''' Multiplies vector elements by a constant.
    ''' </summary>
    ''' <param name="vector"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()> Public Function NormalizeY(vector As Double()) As Double()

        Dim vector0(vector.Length - 1) As Double
        Dim sum As Double = vector.SumY

        If My.MyApplication.UseSIMDExtensions Then
            'Yeppp.Core.Multiply_V64fS64f_V64f(vector, 0, 1 / sum, vector0, 0, vector.Length)
        Else
            For i As Integer = 0 To vector.Length - 1
                vector0(i) = vector(i) / sum
            Next
        End If

        Return vector0

    End Function

    ''' <summary>
    ''' Adds vector elements.
    ''' </summary>
    ''' <param name="vector"></param>
    ''' <param name="vector2"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()> Public Function AddY(vector As Double(), vector2 As Double()) As Double()

        Dim vector0 As Double() = vector.Clone()

        If My.MyApplication.UseSIMDExtensions Then
            'Yeppp.Core.Add_IV64fV64f_IV64f(vector0, 0, vector2, 0, vector.Length)
        Else
            For i As Integer = 0 To vector.Length - 1
                vector0(i) = vector(i) + vector2(i)
            Next
        End If

        Return vector0

    End Function

    ''' <summary>
    ''' Adds a constant value to vector elements.
    ''' </summary>
    ''' <param name="vector"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()> Public Function AddConstY(vector As Double(), constant As Double) As Double()

        Dim vector0 As Double() = vector.Clone()

        If My.MyApplication.UseSIMDExtensions Then
            'Yeppp.Core.Add_V64fS64f_V64f(vector, 0, constant, vector0, 0, vector.Length)
        Else
            For i As Integer = 0 To vector.Length - 1
                vector0(i) = vector(i) + constant
            Next
        End If

        Return vector0

    End Function

#End Region

    ''' <summary>
    ''' Converts a two-dimensional array to a jagged array.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="twoDimensionalArray"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension> Public Function ToJaggedArray(Of T)(twoDimensionalArray As T(,)) As T()()

        Dim rowsFirstIndex As Integer = twoDimensionalArray.GetLowerBound(0)
        Dim rowsLastIndex As Integer = twoDimensionalArray.GetUpperBound(0)
        Dim numberOfRows As Integer = rowsLastIndex + 1

        Dim columnsFirstIndex As Integer = twoDimensionalArray.GetLowerBound(1)
        Dim columnsLastIndex As Integer = twoDimensionalArray.GetUpperBound(1)
        Dim numberOfColumns As Integer = columnsLastIndex + 1

        Dim jaggedArray As T()() = New T(numberOfRows - 1)() {}
        For i As Integer = rowsFirstIndex To rowsLastIndex
            jaggedArray(i) = New T(numberOfColumns - 1) {}

            For j As Integer = columnsFirstIndex To columnsLastIndex
                jaggedArray(i)(j) = twoDimensionalArray(i, j)
            Next
        Next
        Return jaggedArray

    End Function

    ''' <summary>
    ''' Converts a jagged array to a two-dimensional array.
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="jaggedArray"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension> Public Function FromJaggedArray(Of T)(jaggedArray As T()()) As T(,)

        Dim rowsFirstIndex As Integer = jaggedArray.GetLowerBound(0)
        Dim rowsLastIndex As Integer = jaggedArray.GetUpperBound(0)
        Dim numberOfRows As Integer = rowsLastIndex + 1

        Dim columnsFirstIndex As Integer = jaggedArray(0).GetLowerBound(0)
        Dim columnsLastIndex As Integer = jaggedArray(0).GetUpperBound(0)
        Dim numberOfColumns As Integer = columnsLastIndex + 1

        Dim twoDimensionalArray As T(,) = New T(numberOfRows - 1, numberOfColumns - 1) {}
        For i As Integer = rowsFirstIndex To rowsLastIndex
            For j As Integer = columnsFirstIndex To columnsLastIndex
                twoDimensionalArray(i, j) = jaggedArray(i)(j)
            Next
        Next
        Return twoDimensionalArray

    End Function

End Module
