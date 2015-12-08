Imports System.IO

'    Shared Functions
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

Namespace DTL

    <System.Serializable()> Public Class App

        Public Shared Sub WriteToConsole(text As String, minlevel As Integer)

            If My.MyApplication._DebugLevel >= minlevel Then Console.WriteLine(text)

        End Sub

        Public Shared Function GetLocalString(ByVal id As String) As String
            Return id
        End Function

        Public Shared Function GetPropertyName(ByVal PropID As String) As String
            Return PropID
        End Function

        Public Shared Function GetComponentName(ByVal UniqueName As String) As String
            Return UniqueName
        End Function

        Public Shared Function GetComponentType(ByRef comp As DTL.ClassesBasicasTermodinamica.ConstantProperties) As String
            If comp.IsHYPO Then
                Return GetLocalString("CompHypo")
            ElseIf comp.IsPF Then
                Return GetLocalString("CompPseudo")
            Else
                Return GetLocalString("CompNormal")
            End If
        End Function

        Public Shared Function IsRunningOnMono() As Boolean
            Return Not Type.GetType("Mono.Runtime") Is Nothing
        End Function

    End Class

End Namespace
