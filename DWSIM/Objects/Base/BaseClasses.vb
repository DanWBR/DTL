'    Calculation Engine Base Classes 
'    Copyright 2008 Daniel Wagner O. de Medeiros
'
'    This file is part of DTL.
'
'    DWSIM is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    DWSIM is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with DTL.  If not, see <http://www.gnu.org/licenses/>.
'

Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Runtime.Serialization
Imports System.IO
Imports CAPEOPEN110
Imports System.Runtime.Serialization.Formatters
Imports System.Runtime.InteropServices.Marshal
Imports System.Runtime.InteropServices

<System.Serializable()> <ComVisible(True)> Friend MustInherit Class SimulationObjects_BaseClass

    Implements ICloneable, IDisposable

    Protected m_ComponentDescription As String
    Protected m_ComponentName As String

    Public Const ClassId As String = ""

    Protected m_showqtable As Boolean = True

    Public Enum PropertyType
        RO = 0
        RW = 1
        WR = 2
        ALL = 3
    End Enum

    Public MustOverride Function GetProperties(ByVal proptype As PropertyType) As String()

    Public MustOverride Function GetPropertyValue(ByVal prop As String, Optional ByVal su As DTL.SistemasDeUnidades.Unidades = Nothing)
    Public MustOverride Function GetPropertyUnit(ByVal prop As String, Optional ByVal su As DTL.SistemasDeUnidades.Unidades = Nothing)
    Public MustOverride Function SetPropertyValue(ByVal prop As String, ByVal propval As Object, Optional ByVal su As DTL.SistemasDeUnidades.Unidades = Nothing)

    Public Function FT(ByRef prop As String, ByVal unit As String)
        Return prop & " (" & unit & ")"
    End Function

    Public Property Descricao() As String
        Get
            Return m_ComponentDescription
        End Get
        Set(ByVal value As String)
            m_ComponentDescription = value
        End Set
    End Property

    Public Property Nome() As String
        Get
            Return m_ComponentName
        End Get
        Set(ByVal value As String)
            m_ComponentName = value
        End Set
    End Property

    Sub CreateNew()

    End Sub

    Public Function Clone() As Object Implements System.ICloneable.Clone

        Return ObjectCopy(Me)

    End Function

    Function ObjectCopy(ByVal obj As SimulationObjects_BaseClass) As SimulationObjects_BaseClass

        Dim objMemStream As New MemoryStream(50000)
        Dim objBinaryFormatter As New BinaryFormatter(Nothing, New StreamingContext(StreamingContextStates.Clone))

        objBinaryFormatter.Serialize(objMemStream, obj)

        objMemStream.Seek(0, SeekOrigin.Begin)

        ObjectCopy = objBinaryFormatter.Deserialize(objMemStream)

        objMemStream.Close()

    End Function

    Public disposedValue As Boolean = False        ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: free other state (managed objects).
            End If

            ' TODO: free your own state (unmanaged objects).
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

#Region "   IDisposable Support "
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class