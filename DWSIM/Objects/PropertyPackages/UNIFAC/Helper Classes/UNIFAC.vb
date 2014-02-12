﻿'    UNIFAC Property Package 
'    Copyright 2008-2011 Daniel Wagner O. de Medeiros
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

Imports Microsoft.VisualBasic.FileIO
Namespace DTL.SimulationObjects.PropertyPackages.Auxiliary

    <System.Serializable()> Public Class Unifac

        Public UnifGroups As UnifacGroups

        Sub New()

            UnifGroups = New UnifacGroups

        End Sub

        Function ID2Group(ByVal id As Integer) As String

            For Each group As UnifacGroup In Me.UnifGroups.Groups.Values
                If group.Secondary_Group = id Then Return group.GroupName
            Next

            Return ""

        End Function

        Function Group2ID(ByVal groupname As String) As String

            For Each group As UnifacGroup In Me.UnifGroups.Groups.Values
                If group.GroupName = groupname Then
                    Return group.Secondary_Group
                End If
            Next

            Return 0

        End Function

        Function GAMMA(ByVal T, ByVal Vx, ByVal VQ, ByVal VR, ByVal VEKI, ByVal index)

            Dim Q(), R(), j(), L()
            Dim i, k, m As Integer

            Dim n = UBound(Vx)
            Dim n2 = UBound(VEKI, 2)

            Dim beta(,), teta(n2), s(n2), Vgammac(), Vgammar(), Vgamma(), b(,)
            ReDim Preserve beta(n, n2), Vgammac(n), Vgammar(n), Vgamma(n), b(n, n2)
            ReDim Q(n), R(n), j(n), L(n)

            i = 0
            Do
                k = 0
                Do
                    beta(i, k) = 0
                    m = 0
                    Do
                        beta(i, k) = beta(i, k) + VEKI(i, m) * TAU(m, k, T)
                        m = m + 1
                    Loop Until m = n2 + 1
                    k = k + 1
                Loop Until k = n2 + 1
                i = i + 1
            Loop Until i = n + 1

            Dim soma_xq = 0
            i = 0
            Do
                Q(i) = VQ(i)
                soma_xq = soma_xq + Vx(i) * Q(i)
                i = i + 1
            Loop Until i = n + 1

            k = 0
            Do
                i = 0
                Do
                    teta(k) = teta(k) + Vx(i) * Q(i) * VEKI(i, k)
                    i = i + 1
                Loop Until i = n + 1
                teta(k) = teta(k) / soma_xq
                k = k + 1
            Loop Until k = n2 + 1

            k = 0
            Do
                m = 0
                Do
                    s(k) = s(k) + teta(m) * TAU(m, k, T)
                    m = m + 1
                Loop Until m = n2 + 1
                k = k + 1
            Loop Until k = n2 + 1

            Dim soma_xr = 0
            i = 0
            Do
                R(i) = VR(i)
                soma_xr = soma_xr + Vx(i) * R(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                j(i) = R(i) / soma_xr
                L(i) = Q(i) / soma_xq
                Vgammac(i) = 1 - j(i) + Math.Log(j(i)) - 5 * Q(i) * (1 - j(i) / L(i) + Math.Log(j(i) / L(i)))
                k = 0
                Dim tmpsum = 0
                Do
                    tmpsum = tmpsum + teta(k) * beta(i, k) / s(k) - VEKI(i, k) * Math.Log(beta(i, k) / s(k))
                    k = k + 1
                Loop Until k = n2 + 1
                Vgammar(i) = Q(i) * (1 - tmpsum)
                Vgamma(i) = Math.Exp(Vgammac(i) + Vgammar(i))
                If Vgamma(i) = 0 Then Vgamma(i) = 0.000001
                i = i + 1
            Loop Until i = n + 1

            i = 1
            k = 1
            GAMMA = Vgamma(index)

        End Function

        Function GAMMA_DINF(ByVal T, ByVal Vx, ByVal VQ, ByVal VR, ByVal VEKI, ByVal index)

            Dim Q(), R(), j(), L()
            Dim i, k, m As Integer

            Dim n = UBound(Vx)
            Dim n2 = UBound(VEKI, 2)

            Dim beta(,), teta(n2), s(n2), Vgammac(), Vgammar(), Vgamma(), b(,)
            ReDim beta(n, n2), Vgammac(n), Vgammar(n), Vgamma(n), b(n, n2)
            ReDim Q(n), R(n), j(n), L(n)

            i = 0
            Do
                k = 0
                Do
                    m = 0
                    Do
                        beta(i, k) = beta(i, k) + VEKI(i, m) * TAU(m, k, T)
                        m = m + 1
                    Loop Until m = n2 + 1
                    k = k + 1
                Loop Until k = n2 + 1
                i = i + 1
            Loop Until i = n + 1

            Dim soma_xq = 0
            i = 0
            Dim Vxid_ant = Vx(index)
            Do
                Q(i) = VQ(i)
                If i <> index Then
                    Vx(i) = Vx(i) + Vxid_ant / (n - 1)
                End If
                Vx(index) = 0.0000000001
                soma_xq = soma_xq + Vx(i) * Q(i)
                i = i + 1
            Loop Until i = n + 1

            k = 0
            Do
                i = 0
                Do
                    teta(k) = teta(k) + Vx(i) * Q(i) * VEKI(i, k)
                    i = i + 1
                Loop Until i = n + 1
                teta(k) = teta(k) / soma_xq
                k = k + 1
            Loop Until k = n2 + 1

            k = 0
            Do
                m = 0
                Do
                    s(k) = s(k) + teta(m) * TAU(m, k, T)
                    m = m + 1
                Loop Until m = n2 + 1
                k = k + 1
            Loop Until k = n2 + 1

            Dim soma_xr = 0
            i = 0
            Do
                R(i) = VR(i)
                soma_xr = soma_xr + Vx(i) * R(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                j(i) = R(i) / soma_xr
                L(i) = Q(i) / soma_xq
                Vgammac(i) = 1 - j(i) + Math.Log(j(i)) - 5 * Q(i) * (1 - j(i) / L(i) + Math.Log(j(i) / L(i)))
                k = 0
                Dim tmpsum = 0
                Do
                    tmpsum = tmpsum + teta(k) * beta(i, k) / s(k) - VEKI(i, k) * Math.Log(beta(i, k) / s(k))
                    k = k + 1
                Loop Until k = n2 + 1
                Vgammar(i) = Q(i) * (1 - tmpsum)
                Vgamma(i) = Math.Exp(Vgammac(i) + Vgammar(i))
                If Vgamma(i) = 0 Then Vgamma(i) = 0.000001
                i = i + 1
            Loop Until i = n + 1

            GAMMA_DINF = Vgamma(index)

        End Function

        Function GAMMA_MR(ByVal T, ByVal Vx, ByVal VQ, ByVal VR, ByVal VEKI)
            Dim i, k, m As Integer

            Dim n = UBound(Vx)
            Dim n2 = UBound(VEKI, 2)

            Dim teta(n2), s(n2) As Double
            Dim beta(n, n2), Vgammac(n), Vgammar(n), Vgamma(n), b(n, n2) As Double
            Dim Q(n), R(n), j(n), L(n) As Double

            i = 0
            Do
                k = 0
                Do
                    beta(i, k) = 0
                    m = 0
                    Do
                        beta(i, k) = beta(i, k) + VEKI(i, m) * TAU(m, k, T)
                        m = m + 1
                    Loop Until m = n2 + 1
                    k = k + 1
                Loop Until k = n2 + 1
                i = i + 1
            Loop Until i = n + 1

            Dim soma_xq = 0
            i = 0
            Do
                Q(i) = VQ(i)
                soma_xq = soma_xq + Vx(i) * Q(i)
                i = i + 1
            Loop Until i = n + 1

            k = 0
            Do
                i = 0
                Do
                    teta(k) = teta(k) + Vx(i) * Q(i) * VEKI(i, k)
                    i = i + 1
                Loop Until i = n + 1
                teta(k) = teta(k) / soma_xq
                k = k + 1
            Loop Until k = n2 + 1

            k = 0
            Do
                m = 0
                Do
                    s(k) = s(k) + teta(m) * TAU(m, k, T)
                    m = m + 1
                Loop Until m = n2 + 1
                k = k + 1
            Loop Until k = n2 + 1

            Dim soma_xr = 0
            i = 0
            Do
                R(i) = VR(i)
                soma_xr = soma_xr + Vx(i) * R(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                j(i) = R(i) / soma_xr
                L(i) = Q(i) / soma_xq
                Vgammac(i) = 1 - j(i) + Math.Log(j(i)) - 5 * Q(i) * (1 - j(i) / L(i) + Math.Log(j(i) / L(i)))
                k = 0
                Dim tmpsum = 0
                Do
                    tmpsum = tmpsum + teta(k) * beta(i, k) / s(k) - VEKI(i, k) * Math.Log(beta(i, k) / s(k))
                    k = k + 1
                Loop Until k = n2 + 1
                Vgammar(i) = Q(i) * (1 - tmpsum)
                Vgamma(i) = Math.Exp(Vgammac(i) + Vgammar(i))
                If Vgamma(i) = 0 Then Vgamma(i) = 0.000001
                i = i + 1
            Loop Until i = n + 1

            Return Vgamma

        End Function

        Function GAMMA_DINF_MR(ByVal T, ByVal Vx, ByVal VQ, ByVal VR, ByVal VEKI, ByVal index)

            Dim i, k, m As Integer

            Dim n = UBound(Vx)
            Dim n2 = UBound(VEKI, 2)

            Dim teta(n2), s(n2) As Double
            Dim beta(n, n2), Vgammac(n), Vgammar(n), Vgamma(n), b(n, n2) As Double
            Dim Q(n), R(n), j(n), L(n) As Double

            i = 0
            Do
                k = 0
                Do
                    m = 0
                    Do
                        beta(i, k) = beta(i, k) + VEKI(i, m) * TAU(m, k, T)
                        m = m + 1
                    Loop Until m = n2 + 1
                    k = k + 1
                Loop Until k = n2 + 1
                i = i + 1
            Loop Until i = n + 1

            Dim soma_xq = 0
            i = 0
            Dim Vxid_ant = Vx(index)
            Do
                Q(i) = VQ(i)
                If i <> index Then
                    Vx(i) = Vx(i) + Vxid_ant / (n - 1)
                End If
                Vx(index) = 0.0000000001
                soma_xq = soma_xq + Vx(i) * Q(i)
                i = i + 1
            Loop Until i = n + 1

            k = 0
            Do
                i = 0
                Do
                    teta(k) = teta(k) + Vx(i) * Q(i) * VEKI(i, k)
                    i = i + 1
                Loop Until i = n + 1
                teta(k) = teta(k) / soma_xq
                k = k + 1
            Loop Until k = n2 + 1

            k = 0
            Do
                m = 0
                Do
                    s(k) = s(k) + teta(m) * TAU(m, k, T)
                    m = m + 1
                Loop Until m = n2 + 1
                k = k + 1
            Loop Until k = n2 + 1

            Dim soma_xr = 0
            i = 0
            Do
                R(i) = VR(i)
                soma_xr = soma_xr + Vx(i) * R(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                j(i) = R(i) / soma_xr
                L(i) = Q(i) / soma_xq
                Vgammac(i) = 1 - j(i) + Math.Log(j(i)) - 5 * Q(i) * (1 - j(i) / L(i) + Math.Log(j(i) / L(i)))
                k = 0
                Dim tmpsum = 0
                Do
                    tmpsum = tmpsum + teta(k) * beta(i, k) / s(k) - VEKI(i, k) * Math.Log(beta(i, k) / s(k))
                    k = k + 1
                Loop Until k = n2 + 1
                Vgammar(i) = Q(i) * (1 - tmpsum)
                Vgamma(i) = Math.Exp(Vgammac(i) + Vgammar(i))
                If Vgamma(i) = 0 Then Vgamma(i) = 0.000001
                i = i + 1
            Loop Until i = n + 1

            Return Vgamma

        End Function

        Function TAU(ByVal group_1, ByVal group_2, ByVal T)

            Dim g1, g2 As Integer
            Dim res As Double
            g1 = Me.UnifGroups.Groups(group_1 + 1).PrimaryGroup
            g2 = Me.UnifGroups.Groups(group_2 + 1).PrimaryGroup

            If Me.UnifGroups.InteracParam.ContainsKey(g1) Then
                If Me.UnifGroups.InteracParam(g1).ContainsKey(g2) Then
                    res = Me.UnifGroups.InteracParam(g1)(g2)
                Else
                    res = 0
                End If
            Else
                res = 0
            End If

            Return Math.Exp(-res / T)

        End Function

        Function RET_Ri(ByVal VN As Object) As Double

            Dim i As Integer = 0
            Dim res As Double

            For Each group As UnifacGroup In Me.UnifGroups.Groups.Values
                res += group.R * VN(i)
                i += 1
            Next

            Return res

        End Function

        Function RET_Qi(ByVal VN As Object) As Double

            Dim i As Integer = 0
            Dim res As Double

            For Each group As UnifacGroup In Me.UnifGroups.Groups.Values
                res += group.Q * VN(i)
                i += 1
            Next

            Return res

        End Function

        Function RET_EKI(ByVal VN As Object, ByVal Q As Double) As Object

            Dim i As Integer = 0
            Dim res As New ArrayList

            For Each group As UnifacGroup In Me.UnifGroups.Groups.Values
                If Q <> 0 Then res.Add(group.Q * VN(i) / Q) Else res.Add(0.0#)
                i += 1
            Next

            Return res.ToArray(Type.GetType("System.Double"))

        End Function

        Function RET_VN(ByVal cp As DTL.ClassesBasicasTermodinamica.ConstantProperties)

            Dim i As Integer = 0
            Dim res As New ArrayList
            Dim added As Boolean = False

            res.Clear()

            For Each group As UnifacGroup In Me.UnifGroups.Groups.Values
                For Each s As String In cp.UNIFACGroups.Collection.Keys
                    If s = group.GroupName Then
                        res.Add(cp.UNIFACGroups.Collection(s))
                        added = True
                        Exit For
                    End If
                Next
                If Not added Then res.Add(0)
                added = False
            Next

            Return res.ToArray

        End Function

    End Class

    <System.Serializable()> Public Class UnifacLL

        Public UnifGroups As UnifacGroupsLL

        Sub New()

            UnifGroups = New UnifacGroupsLL

        End Sub

        Function ID2Group(ByVal id As Integer) As String

            For Each group As UnifacGroup In Me.UnifGroups.Groups.Values
                If group.Secondary_Group = id Then Return group.GroupName
            Next

            Return ""

        End Function

        Function Group2ID(ByVal groupname As String) As String

            For Each group As UnifacGroup In Me.UnifGroups.Groups.Values
                If group.GroupName = groupname Then
                    Return group.Secondary_Group
                End If
            Next

            Return 0

        End Function

        Function GAMMA(ByVal T, ByVal Vx, ByVal VQ, ByVal VR, ByVal VEKI, ByVal index)

            Dim Q(), R(), j(), L()
            Dim i, k, m As Integer

            Dim n = UBound(Vx)
            Dim n2 = UBound(VEKI, 2)

            Dim beta(,), teta(n2), s(n2), Vgammac(), Vgammar(), Vgamma(), b(,)
            ReDim Preserve beta(n, n2), Vgammac(n), Vgammar(n), Vgamma(n), b(n, n2)
            ReDim Q(n), R(n), j(n), L(n)

            i = 0
            Do
                k = 0
                Do
                    beta(i, k) = 0
                    m = 0
                    Do
                        beta(i, k) = beta(i, k) + VEKI(i, m) * TAU(m, k, T)
                        m = m + 1
                    Loop Until m = n2 + 1
                    k = k + 1
                Loop Until k = n2 + 1
                i = i + 1
            Loop Until i = n + 1

            Dim soma_xq = 0
            i = 0
            Do
                Q(i) = VQ(i)
                soma_xq = soma_xq + Vx(i) * Q(i)
                i = i + 1
            Loop Until i = n + 1

            k = 0
            Do
                i = 0
                Do
                    teta(k) = teta(k) + Vx(i) * Q(i) * VEKI(i, k)
                    i = i + 1
                Loop Until i = n + 1
                teta(k) = teta(k) / soma_xq
                k = k + 1
            Loop Until k = n2 + 1

            k = 0
            Do
                m = 0
                Do
                    s(k) = s(k) + teta(m) * TAU(m, k, T)
                    m = m + 1
                Loop Until m = n2 + 1
                k = k + 1
            Loop Until k = n2 + 1

            Dim soma_xr = 0
            i = 0
            Do
                R(i) = VR(i)
                soma_xr = soma_xr + Vx(i) * R(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                j(i) = R(i) / soma_xr
                L(i) = Q(i) / soma_xq
                Vgammac(i) = 1 - j(i) + Math.Log(j(i)) - 5 * Q(i) * (1 - j(i) / L(i) + Math.Log(j(i) / L(i)))
                k = 0
                Dim tmpsum = 0
                Do
                    tmpsum = tmpsum + teta(k) * beta(i, k) / s(k) - VEKI(i, k) * Math.Log(beta(i, k) / s(k))
                    k = k + 1
                Loop Until k = n2 + 1
                Vgammar(i) = Q(i) * (1 - tmpsum)
                Vgamma(i) = Math.Exp(Vgammac(i) + Vgammar(i))
                If Vgamma(i) = 0 Then Vgamma(i) = 0.000001
                i = i + 1
            Loop Until i = n + 1

            i = 1
            k = 1
            GAMMA = Vgamma(index)

        End Function

        Function GAMMA_DINF(ByVal T, ByVal Vx, ByVal VQ, ByVal VR, ByVal VEKI, ByVal index)

            Dim Q(), R(), j(), L()
            Dim i, k, m As Integer

            Dim n = UBound(Vx)
            Dim n2 = UBound(VEKI, 2)

            Dim beta(,), teta(n2), s(n2), Vgammac(), Vgammar(), Vgamma(), b(,)
            ReDim beta(n, n2), Vgammac(n), Vgammar(n), Vgamma(n), b(n, n2)
            ReDim Q(n), R(n), j(n), L(n)

            i = 0
            Do
                k = 0
                Do
                    m = 0
                    Do
                        beta(i, k) = beta(i, k) + VEKI(i, m) * TAU(m, k, T)
                        m = m + 1
                    Loop Until m = n2 + 1
                    k = k + 1
                Loop Until k = n2 + 1
                i = i + 1
            Loop Until i = n + 1

            Dim soma_xq = 0
            i = 0
            Dim Vxid_ant = Vx(index)
            Do
                Q(i) = VQ(i)
                If i <> index Then
                    Vx(i) = Vx(i) + Vxid_ant / (n - 1)
                End If
                Vx(index) = 0.0000000001
                soma_xq = soma_xq + Vx(i) * Q(i)
                i = i + 1
            Loop Until i = n + 1

            k = 0
            Do
                i = 0
                Do
                    teta(k) = teta(k) + Vx(i) * Q(i) * VEKI(i, k)
                    i = i + 1
                Loop Until i = n + 1
                teta(k) = teta(k) / soma_xq
                k = k + 1
            Loop Until k = n2 + 1

            k = 0
            Do
                m = 0
                Do
                    s(k) = s(k) + teta(m) * TAU(m, k, T)
                    m = m + 1
                Loop Until m = n2 + 1
                k = k + 1
            Loop Until k = n2 + 1

            Dim soma_xr = 0
            i = 0
            Do
                R(i) = VR(i)
                soma_xr = soma_xr + Vx(i) * R(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                j(i) = R(i) / soma_xr
                L(i) = Q(i) / soma_xq
                Vgammac(i) = 1 - j(i) + Math.Log(j(i)) - 5 * Q(i) * (1 - j(i) / L(i) + Math.Log(j(i) / L(i)))
                k = 0
                Dim tmpsum = 0
                Do
                    tmpsum = tmpsum + teta(k) * beta(i, k) / s(k) - VEKI(i, k) * Math.Log(beta(i, k) / s(k))
                    k = k + 1
                Loop Until k = n2 + 1
                Vgammar(i) = Q(i) * (1 - tmpsum)
                Vgamma(i) = Math.Exp(Vgammac(i) + Vgammar(i))
                If Vgamma(i) = 0 Then Vgamma(i) = 0.000001
                i = i + 1
            Loop Until i = n + 1

            GAMMA_DINF = Vgamma(index)

        End Function

        Function GAMMA_MR(ByVal T, ByVal Vx, ByVal VQ, ByVal VR, ByVal VEKI)
            Dim i, k, m As Integer

            Dim n = UBound(Vx)
            Dim n2 = UBound(VEKI, 2)

            Dim teta(n2), s(n2) As Double
            Dim beta(n, n2), Vgammac(n), Vgammar(n), Vgamma(n), b(n, n2) As Double
            Dim Q(n), R(n), j(n), L(n) As Double

            i = 0
            Do
                k = 0
                Do
                    beta(i, k) = 0
                    m = 0
                    Do
                        beta(i, k) = beta(i, k) + VEKI(i, m) * TAU(m, k, T)
                        m = m + 1
                    Loop Until m = n2 + 1
                    k = k + 1
                Loop Until k = n2 + 1
                i = i + 1
            Loop Until i = n + 1

            Dim soma_xq = 0
            i = 0
            Do
                Q(i) = VQ(i)
                soma_xq = soma_xq + Vx(i) * Q(i)
                i = i + 1
            Loop Until i = n + 1

            k = 0
            Do
                i = 0
                Do
                    teta(k) = teta(k) + Vx(i) * Q(i) * VEKI(i, k)
                    i = i + 1
                Loop Until i = n + 1
                teta(k) = teta(k) / soma_xq
                k = k + 1
            Loop Until k = n2 + 1

            k = 0
            Do
                m = 0
                Do
                    s(k) = s(k) + teta(m) * TAU(m, k, T)
                    m = m + 1
                Loop Until m = n2 + 1
                k = k + 1
            Loop Until k = n2 + 1

            Dim soma_xr = 0
            i = 0
            Do
                R(i) = VR(i)
                soma_xr = soma_xr + Vx(i) * R(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                j(i) = R(i) / soma_xr
                L(i) = Q(i) / soma_xq
                Vgammac(i) = 1 - j(i) + Math.Log(j(i)) - 5 * Q(i) * (1 - j(i) / L(i) + Math.Log(j(i) / L(i)))
                k = 0
                Dim tmpsum = 0
                Do
                    tmpsum = tmpsum + teta(k) * beta(i, k) / s(k) - VEKI(i, k) * Math.Log(beta(i, k) / s(k))
                    k = k + 1
                Loop Until k = n2 + 1
                Vgammar(i) = Q(i) * (1 - tmpsum)
                Vgamma(i) = Math.Exp(Vgammac(i) + Vgammar(i))
                If Vgamma(i) = 0 Then Vgamma(i) = 0.000001
                i = i + 1
            Loop Until i = n + 1

            Return Vgamma

        End Function

        Function GAMMA_DINF_MR(ByVal T, ByVal Vx, ByVal VQ, ByVal VR, ByVal VEKI, ByVal index)

            Dim i, k, m As Integer

            Dim n = UBound(Vx)
            Dim n2 = UBound(VEKI, 2)

            Dim teta(n2), s(n2) As Double
            Dim beta(n, n2), Vgammac(n), Vgammar(n), Vgamma(n), b(n, n2) As Double
            Dim Q(n), R(n), j(n), L(n) As Double

            i = 0
            Do
                k = 0
                Do
                    m = 0
                    Do
                        beta(i, k) = beta(i, k) + VEKI(i, m) * TAU(m, k, T)
                        m = m + 1
                    Loop Until m = n2 + 1
                    k = k + 1
                Loop Until k = n2 + 1
                i = i + 1
            Loop Until i = n + 1

            Dim soma_xq = 0
            i = 0
            Dim Vxid_ant = Vx(index)
            Do
                Q(i) = VQ(i)
                If i <> index Then
                    Vx(i) = Vx(i) + Vxid_ant / (n - 1)
                End If
                Vx(index) = 0.0000000001
                soma_xq = soma_xq + Vx(i) * Q(i)
                i = i + 1
            Loop Until i = n + 1

            k = 0
            Do
                i = 0
                Do
                    teta(k) = teta(k) + Vx(i) * Q(i) * VEKI(i, k)
                    i = i + 1
                Loop Until i = n + 1
                teta(k) = teta(k) / soma_xq
                k = k + 1
            Loop Until k = n2 + 1

            k = 0
            Do
                m = 0
                Do
                    s(k) = s(k) + teta(m) * TAU(m, k, T)
                    m = m + 1
                Loop Until m = n2 + 1
                k = k + 1
            Loop Until k = n2 + 1

            Dim soma_xr = 0
            i = 0
            Do
                R(i) = VR(i)
                soma_xr = soma_xr + Vx(i) * R(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                j(i) = R(i) / soma_xr
                L(i) = Q(i) / soma_xq
                Vgammac(i) = 1 - j(i) + Math.Log(j(i)) - 5 * Q(i) * (1 - j(i) / L(i) + Math.Log(j(i) / L(i)))
                k = 0
                Dim tmpsum = 0
                Do
                    tmpsum = tmpsum + teta(k) * beta(i, k) / s(k) - VEKI(i, k) * Math.Log(beta(i, k) / s(k))
                    k = k + 1
                Loop Until k = n2 + 1
                Vgammar(i) = Q(i) * (1 - tmpsum)
                Vgamma(i) = Math.Exp(Vgammac(i) + Vgammar(i))
                If Vgamma(i) = 0 Then Vgamma(i) = 0.000001
                i = i + 1
            Loop Until i = n + 1

            Return Vgamma

        End Function

        Function TAU(ByVal group_1, ByVal group_2, ByVal T)

            Dim g1, g2 As Integer
            Dim res As Double
            g1 = Me.UnifGroups.Groups(group_1 + 1).PrimaryGroup
            g2 = Me.UnifGroups.Groups(group_2 + 1).PrimaryGroup

            If Me.UnifGroups.InteracParam.ContainsKey(g1) Then
                If Me.UnifGroups.InteracParam(g1).ContainsKey(g2) Then
                    res = Me.UnifGroups.InteracParam(g1)(g2)
                Else
                    res = 0
                End If
            Else
                res = 0
            End If

            Return Math.Exp(-res / T)

        End Function

        Function RET_Ri(ByVal VN As Object) As Double

            Dim i As Integer = 0
            Dim res As Double

            For Each group As UnifacGroup In Me.UnifGroups.Groups.Values
                res += group.R * VN(i)
                i += 1
            Next

            Return res

        End Function

        Function RET_Qi(ByVal VN As Object) As Double

            Dim i As Integer = 0
            Dim res As Double

            For Each group As UnifacGroup In Me.UnifGroups.Groups.Values
                res += group.Q * VN(i)
                i += 1
            Next

            Return res

        End Function

        Function RET_EKI(ByVal VN As Object, ByVal Q As Double) As Object

            Dim i As Integer = 0
            Dim res As New ArrayList

            For Each group As UnifacGroup In Me.UnifGroups.Groups.Values
                If Q <> 0 Then res.Add(group.Q * VN(i) / Q) Else res.Add(0.0#)
                i += 1
            Next

            Return res.ToArray(Type.GetType("System.Double"))

        End Function

        Function RET_VN(ByVal cp As DTL.ClassesBasicasTermodinamica.ConstantProperties)

            Dim i As Integer = 0
            Dim res As New ArrayList
            Dim added As Boolean = False

            res.Clear()

            For Each group As UnifacGroup In Me.UnifGroups.Groups.Values
                For Each s As String In cp.UNIFACGroups.Collection.Keys
                    If s = group.GroupName Then
                        res.Add(cp.UNIFACGroups.Collection(s))
                        added = True
                        Exit For
                    End If
                Next
                If Not added Then res.Add(0)
                added = False
            Next

            Return res.ToArray

        End Function

    End Class

    <System.Serializable()> Public Class UnifacGroups

        Public InteracParam As System.Collections.Generic.Dictionary(Of Integer, System.Collections.Generic.Dictionary(Of Integer, Double))
        Protected m_groups As System.Collections.Generic.Dictionary(Of Integer, UnifacGroup)

        Sub New()

            Dim pathsep = System.IO.Path.DirectorySeparatorChar

            m_groups = New System.Collections.Generic.Dictionary(Of Integer, UnifacGroup)
            InteracParam = New System.Collections.Generic.Dictionary(Of Integer, System.Collections.Generic.Dictionary(Of Integer, Double))

            Dim cult As Globalization.CultureInfo = New Globalization.CultureInfo("en-US")

            Dim fields As String()
            Dim delimiter As String = ","
            Using stream As IO.Stream = New IO.MemoryStream(My.Resources.unifac)
                Using reader As New IO.StreamReader(stream)
                    Using parser As New TextFieldParser(reader)
                        parser.SetDelimiters(delimiter)
                        fields = parser.ReadFields()
                        fields = parser.ReadFields()
                        While Not parser.EndOfData
                            fields = parser.ReadFields()
                            Me.Groups.Add(fields(1), New UnifacGroup(fields(3), fields(0), fields(1), Double.Parse(fields(4), cult), Double.Parse(fields(5), cult)))
                        End While
                    End Using
                End Using
            End Using

            Using stream As IO.Stream = New IO.MemoryStream(My.Resources.unifac_ip)
                Using reader As New IO.StreamReader(stream)
                    Using parser As New TextFieldParser(reader)
                        delimiter = ","
                        parser.SetDelimiters(delimiter)
                        While Not parser.EndOfData
                            fields = parser.ReadFields()
                            If Not Me.InteracParam.ContainsKey(fields(0)) Then
                                Me.InteracParam.Add(fields(0), New System.Collections.Generic.Dictionary(Of Integer, Double))
                                Me.InteracParam(fields(0)).Add(fields(2), Double.Parse(fields(4), cult))
                            Else
                                If Not Me.InteracParam(fields(0)).ContainsKey(fields(2)) Then
                                    Me.InteracParam(fields(0)).Add(fields(2), Double.Parse(fields(4), cult))
                                Else
                                    Me.InteracParam(fields(0))(fields(2)) = Double.Parse(fields(4), cult)
                                End If
                            End If
                            If Not Me.InteracParam.ContainsKey(fields(2)) Then
                                Me.InteracParam.Add(fields(2), New System.Collections.Generic.Dictionary(Of Integer, Double))
                                Me.InteracParam(fields(2)).Add(fields(0), Double.Parse(fields(5), cult))
                            Else
                                If Not Me.InteracParam(fields(2)).ContainsKey(fields(0)) Then
                                    Me.InteracParam(fields(2)).Add(fields(0), Double.Parse(fields(5), cult))
                                Else
                                    Me.InteracParam(fields(2))(fields(0)) = Double.Parse(fields(5), cult)
                                End If
                            End If
                        End While
                    End Using
                End Using
            End Using




        End Sub

        Public ReadOnly Property Groups() As System.Collections.Generic.Dictionary(Of Integer, UnifacGroup)
            Get
                Return m_groups
            End Get
        End Property

    End Class

    <System.Serializable()> Public Class UnifacGroupsLL

        Public InteracParam As System.Collections.Generic.Dictionary(Of Integer, System.Collections.Generic.Dictionary(Of Integer, Double))
        Protected m_groups As System.Collections.Generic.Dictionary(Of Integer, UnifacGroup)

        Sub New()

            Dim pathsep = System.IO.Path.DirectorySeparatorChar

            m_groups = New System.Collections.Generic.Dictionary(Of Integer, UnifacGroup)
            InteracParam = New System.Collections.Generic.Dictionary(Of Integer, System.Collections.Generic.Dictionary(Of Integer, Double))

            Dim cult As Globalization.CultureInfo = New Globalization.CultureInfo("en-US")

            Dim fields As String()
            Dim delimiter As String = ","
            Using stream As IO.Stream = New IO.MemoryStream(My.Resources.unifac)
                Using reader As New IO.StreamReader(stream)
                    Using parser As New TextFieldParser(reader)
                        parser.SetDelimiters(delimiter)
                        fields = parser.ReadFields()
                        fields = parser.ReadFields()
                        While Not parser.EndOfData
                            fields = parser.ReadFields()
                            Me.Groups.Add(fields(1), New UnifacGroup(fields(3), fields(0), fields(1), Double.Parse(fields(4), cult), Double.Parse(fields(5), cult)))
                        End While
                    End Using
                End Using
            End Using

            Using stream As IO.Stream = New IO.MemoryStream(My.Resources.unifac_ll_ip)
                Using reader As New IO.StreamReader(stream)
                    Using parser As New TextFieldParser(reader)
                        delimiter = ","
                        parser.SetDelimiters(delimiter)
                        While Not parser.EndOfData
                            fields = parser.ReadFields()
                            If Not Me.InteracParam.ContainsKey(fields(0)) Then
                                Me.InteracParam.Add(fields(0), New System.Collections.Generic.Dictionary(Of Integer, Double))
                                Me.InteracParam(fields(0)).Add(fields(2), Double.Parse(fields(4), cult))
                            Else
                                If Not Me.InteracParam(fields(0)).ContainsKey(fields(2)) Then
                                    Me.InteracParam(fields(0)).Add(fields(2), Double.Parse(fields(4), cult))
                                Else
                                    Me.InteracParam(fields(0))(fields(2)) = Double.Parse(fields(4), cult)
                                End If
                            End If
                            If Not Me.InteracParam.ContainsKey(fields(2)) Then
                                Me.InteracParam.Add(fields(2), New System.Collections.Generic.Dictionary(Of Integer, Double))
                                Me.InteracParam(fields(2)).Add(fields(0), Double.Parse(fields(5), cult))
                            Else
                                If Not Me.InteracParam(fields(2)).ContainsKey(fields(0)) Then
                                    Me.InteracParam(fields(2)).Add(fields(0), Double.Parse(fields(5), cult))
                                Else
                                    Me.InteracParam(fields(2))(fields(0)) = Double.Parse(fields(5), cult)
                                End If
                            End If
                        End While
                    End Using
                End Using
            End Using

        End Sub

        Public ReadOnly Property Groups() As System.Collections.Generic.Dictionary(Of Integer, UnifacGroup)
            Get
                Return m_groups
            End Get
        End Property

    End Class

    <System.Serializable()> Public Class UnifacGroup

        Protected m_groupname As String

        Protected m_main_group As Integer
        Protected m_secondary_group As Integer

        Protected m_r As Double
        Protected m_q As Double

        Public Property GroupName() As String
            Get
                Return m_groupname
            End Get
            Set(ByVal value As String)
                m_groupname = value
            End Set
        End Property

        Public Property PrimaryGroup() As String
            Get
                Return m_main_group
            End Get
            Set(ByVal value As String)
                m_main_group = value
            End Set
        End Property

        Public Property Secondary_Group() As String
            Get
                Return m_secondary_group
            End Get
            Set(ByVal value As String)
                m_secondary_group = value
            End Set
        End Property

        Public Property R() As Double
            Get
                Return m_r
            End Get
            Set(ByVal value As Double)
                m_r = value
            End Set
        End Property

        Public Property Q() As Double
            Get
                Return m_q
            End Get
            Set(ByVal value As Double)
                m_q = value
            End Set
        End Property

        Sub New(ByVal Nome As String, ByVal PrimGroup As String, ByVal SecGroup As String, ByVal R As Double, ByVal Q As Double)
            Me.GroupName = Nome
            Me.PrimaryGroup = PrimGroup
            Me.Secondary_Group = SecGroup
            Me.R = R
            Me.Q = Q
        End Sub

        Sub New()
        End Sub

    End Class

End Namespace