﻿'    Peng-Robinson Property Package 
'    Copyright 2008 Daniel Wagner O. de Medeiros
'
'    This file is part of DWSIM.
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
'    along with DWSIM.  If not, see <http://www.gnu.org/licenses/>.
Imports FileHelpers
Imports System.Linq

Namespace DTL.SimulationObjects.PropertyPackages.Auxiliary

    <DelimitedRecord(";")> <IgnoreFirst()> <Serializable()> _
    Public Class PR_IPData

        Implements ICloneable

        Public ID1 As Integer = -1
        Public ID2 As Integer = -1
        Public kij As Double = 0.0#
        Public comment As String = ""

        Public Function Clone() As Object Implements ICloneable.Clone

            Dim newclass As New PR_IPData
            With newclass
                .ID1 = Me.ID1
                .ID2 = Me.ID2
                .kij = Me.kij
                .comment = Me.comment
            End With
            Return newclass
        End Function

    End Class

    <Serializable()> Public Class PengRobinson

        Dim m_pr As New PROPS
        Private _ip As Dictionary(Of String, Dictionary(Of String, PR_IPData))

        Public ReadOnly Property InteractionParameters() As Dictionary(Of String, Dictionary(Of String, PR_IPData))
            Get
                Return _ip
            End Get
        End Property

        Sub New()
            _ip = New Dictionary(Of String, Dictionary(Of String, PR_IPData))

            Dim pathsep As Char = IO.Path.DirectorySeparatorChar

            Dim prip As PR_IPData
            Dim pripc() As PR_IPData
            Dim fh1 As New FileHelperEngine(Of PR_IPData)
            Using stream As IO.Stream = New IO.MemoryStream(My.Resources.pr_ip)
                Using reader As New IO.StreamReader(stream)
                    pripc = fh1.ReadStream(reader)
                End Using
            End Using

            Dim csdb As New Databases.ChemSep
            For Each prip In pripc
                If Me.InteractionParameters.ContainsKey(csdb.GetCSName(prip.ID1)) Then
                    If Me.InteractionParameters(csdb.GetCSName(prip.ID1)).ContainsKey(csdb.GetCSName(prip.ID2)) Then
                    Else
                        Me.InteractionParameters(csdb.GetCSName(prip.ID1)).Add(csdb.GetCSName(prip.ID2), prip.Clone)
                    End If
                Else
                    Me.InteractionParameters.Add(csdb.GetCSName(prip.ID1), New Dictionary(Of String, PR_IPData))
                    Me.InteractionParameters(csdb.GetCSName(prip.ID1)).Add(csdb.GetCSName(prip.ID2), prip.Clone)
                End If
            Next
            For Each prip In pripc
                If Me.InteractionParameters.ContainsKey(csdb.GetDWSIMName(prip.ID1)) Then
                    If Me.InteractionParameters(csdb.GetDWSIMName(prip.ID1)).ContainsKey(csdb.GetDWSIMName(prip.ID2)) Then
                    Else
                        Me.InteractionParameters(csdb.GetDWSIMName(prip.ID1)).Add(csdb.GetDWSIMName(prip.ID2), prip.Clone)
                    End If
                Else
                    Me.InteractionParameters.Add(csdb.GetDWSIMName(prip.ID1), New Dictionary(Of String, PR_IPData))
                    Me.InteractionParameters(csdb.GetDWSIMName(prip.ID1)).Add(csdb.GetDWSIMName(prip.ID2), prip.Clone)
                End If
            Next
            prip = Nothing
            pripc = Nothing
            fh1 = Nothing
        End Sub

        Function JT_PR(ByVal TYPE, ByVal T, ByVal P, ByVal Vz, ByVal VKij, ByVal Vzmass, ByVal VTc, ByVal VPc, ByVal VCpig, ByVal VMM, ByVal Vw, ByVal VZRa)

            Dim n, R, Cpm_ig As Double
            Dim vetor(8) As Double
            Dim Tc(), Pc(), Vc(), W(), Zc(), a, b, c, Tr() As Double

            n = Vz.Length - 1

            ReDim Zc(n), Tc(n), Pc(n), Vc(n), W(n), Tr(n)

            R = 8.314

            Dim i, j As Integer
            i = 0
            Do
                Tc(i) = VTc(i)
                Tr(i) = T / Tc(i)
                Pc(i) = VPc(i)
                W(i) = Vw(i)
                Vc(i) = Zc1(W(i)) * R * Tc(i) / Pc(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Dim Vcm = 0.0#
            Dim wm = 0.0#
            Dim Zcm = 0.0#
            Dim MMm = 0.0#
            Dim ZRam = 0.0#
            Do
                If Vz(i) <> 0 Then
                    Vcm += Vz(i) * Vc(i)
                    wm += Vz(i) * W(i)
                    Zcm += Vz(i) * Zc1(W(i))
                    MMm += Vz(i) * VMM(i)
                    ZRam = Vz(i) * VZRa(i)
                End If
                i += 1
            Loop Until i = n + 1

            i = 0
            Dim Tcm = 0.0#
            Do
                j = 0
                Do
                    If Vz(i) <> 0 And Vz(j) <> 0 Then Tcm += Vz(i) * Vz(j) * (Tc(i) * Tc(j)) ^ 0.5
                    j += 1
                Loop Until j = n + 1
                i += 1
            Loop Until i = n + 1

            Dim Pcm = Zcm * R * Tcm / (Vcm)

            Dim V = 0.0#
            If TYPE = "L" Then

                V = (Z_PR(T, P, Vz, VKij, VTc, VPc, Vw, "L") * R * T / P) * 1000 ' m3/kgmol

            ElseIf TYPE = "V" Then

                V = (Z_PR(T, P, Vz, VKij, VTc, VPc, Vw, "V") * R * T / P) * 1000 ' m3/kgmol

            End If

            a = 0.45724 * R ^ 2 * Tcm ^ 2 / Pcm
            b = 0.0778 * R * Tcm / Pcm
            c = 0.37464 + 1.54226 * wm - 0.26992 * wm ^ 2

            Dim Trm = T / Tcm
            Dim AG = -b * R * T
            Dim BG = -2 * a * T * (1 + c - c * Trm ^ 0.5) * (-0.5 * c * Trm ^ 0.5)
            Dim CG = a * (1 + c - c * Trm ^ 0.5) ^ 2

            Dim dP_dT_V = R / (V - b) - (a * (1 + c) * c * Tcm ^ -0.5 * T ^ -0.5 + a * c ^ 2 * Tcm ^ -1) / (V ^ 2 + 2 * b * V - b ^ 2)

            Dim dP_dV_T = -R * T / (V - b) ^ 2 + (2 * b + 2 * V) * (a * (1 + c) ^ 2 + 2 * a * (1 + c) * c * Tcm ^ -0.5 * T ^ 0.5 + a * c ^ 2 * Tcm ^ -1 * T) / (V ^ 2 + 2 * b * V - b ^ 2) ^ 2

            Cpm_ig = 0
            i = 0
            Do
                Cpm_ig += Vzmass(i) * VCpig(i)
                i += 1
            Loop Until i = n + 1

            Dim JT = -(T * dP_dT_V / dP_dV_T + V) / (Cpm_ig * MMm)

            JT_PR = JT

        End Function

        Function Zc1(ByVal w As Double) As Double

            Zc1 = 0.291 - 0.08 * w

        End Function

        Function bi(ByVal omega As Double, ByVal Tc As Double, ByVal Pc As Double) As Double

            Return omega * 8.314 * Tc / Pc

        End Function

        Function Z_PR(ByVal T, ByVal P, ByVal Vx, ByVal VKij, ByVal VTc, ByVal VPc, ByVal Vw, ByVal TYPE)

            App.WriteToConsole("PR cubic equation root finder (Z) for T = " & T & " K, P = " & P & " Pa and Phase = " & TYPE, 3)
            App.WriteToConsole("Mole fractions: " & DirectCast(Vx, Double()).ToArrayString, 3)

            Dim ai(), bi(), aml2(), amv2() As Double
            Dim n, R, coeff(3), tmp() As Double
            Dim Tc(), Pc(), W(), alpha(), Vant(0, 4), m(), a(,), b(,), Tr() As Double

            n = Vx.Length - 1

            ReDim ai(n), bi(n), tmp(n + 1), a(n, n), b(n, n)
            ReDim aml2(n), amv2(n)
            ReDim Tc(n), Pc(n), W(n), alpha(n), m(n), Tr(n)

            R = 8.314

            Dim i, j As Integer
            i = 0
            Do
                Tc(i) = VTc(i)
                Tr(i) = T / Tc(i)
                Pc(i) = VPc(i)
                W(i) = Vw(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                alpha(i) = (1 + (0.37464 + 1.54226 * W(i) - 0.26992 * W(i) ^ 2) * (1 - (T / Tc(i)) ^ 0.5)) ^ 2
                ai(i) = 0.45724 * alpha(i) * R ^ 2 * Tc(i) ^ 2 / Pc(i)
                bi(i) = 0.0778 * R * Tc(i) / Pc(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                j = 0
                Do
                    a(i, j) = (ai(i) * ai(j)) ^ 0.5 * (1 - VKij(i, j))
                    j = j + 1
                Loop Until j = n + 1
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                aml2(i) = 0
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Dim aml = 0.0#
            Do
                j = 0
                Do
                    aml = aml + Vx(i) * Vx(j) * a(i, j)
                    aml2(i) = aml2(i) + Vx(j) * a(j, i)
                    j = j + 1
                Loop Until j = n + 1
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Dim bml = 0.0#
            Do
                bml = bml + Vx(i) * bi(i)
                i = i + 1
            Loop Until i = n + 1

            Dim AG = aml * P / (R * T) ^ 2
            Dim BG = bml * P / (R * T)

            coeff(0) = -AG * BG + BG ^ 2 + BG ^ 3
            coeff(1) = AG - 3 * BG ^ 2 - 2 * BG
            coeff(2) = BG - 1
            coeff(3) = 1

            Dim temp1 = Poly_Roots(coeff)
            Dim tv = 0.0#
            Dim ZV, tv2 As Double

            If Not IsNumeric(temp1) Then

                If temp1(0, 0) > temp1(1, 0) Then
                    tv = temp1(1, 0)
                    temp1(1, 0) = temp1(0, 0)
                    temp1(0, 0) = tv
                    tv2 = temp1(1, 1)
                    temp1(1, 1) = temp1(0, 1)
                    temp1(0, 1) = tv2
                End If
                If temp1(0, 0) > temp1(2, 0) Then
                    tv = temp1(2, 0)
                    temp1(2, 0) = temp1(0, 0)
                    temp1(0, 0) = tv
                    tv2 = temp1(2, 1)
                    temp1(2, 1) = temp1(0, 1)
                    temp1(0, 1) = tv2
                End If
                If temp1(1, 0) > temp1(2, 0) Then
                    tv = temp1(2, 0)
                    temp1(2, 0) = temp1(1, 0)
                    temp1(1, 0) = tv
                    tv2 = temp1(2, 1)
                    temp1(2, 1) = temp1(1, 1)
                    temp1(1, 1) = tv2
                End If

                ZV = temp1(2, 0)
                If temp1(2, 1) <> 0 Then
                    ZV = temp1(1, 0)
                    If temp1(1, 1) <> 0 Then
                        ZV = temp1(0, 0)
                    End If
                End If

            Else

                Dim findZV, dfdz, zant As Double
                If TYPE = "V" Then ZV = 1 Else ZV = 0.05
                Do
                    findZV = coeff(3) * ZV ^ 3 + coeff(2) * ZV ^ 2 + coeff(1) * ZV + coeff(0)
                    dfdz = 3 * coeff(3) * ZV ^ 2 + 2 * coeff(2) * ZV + coeff(1)
                    zant = ZV
                    ZV = ZV - findZV / dfdz
                    If ZV < 0 Then ZV = 1
                Loop Until Math.Abs(findZV) < 0.0001 Or Double.IsNaN(ZV)

                Return ZV

            End If

            Z_PR = 0
            If TYPE = "L" Then
                Z_PR = temp1(0, 0)
            ElseIf TYPE = "V" Then
                Z_PR = temp1(2, 0)
            End If

            App.WriteToConsole("Result: Z = " & Z_PR, 3)

        End Function

        Function H_PR(ByVal TYPE As String, ByVal T As Double, ByVal P As Double, ByVal Tc As Double, ByVal Pc As Double, ByVal w As Double, ByVal MM As Double, Optional ByVal ZRa As Double = 0) As Double

            Dim R As Double
            Dim a, c As Double

            R = 8.314

            a = 0.45724 * R ^ 2 * Tc ^ 2 / Pc
            c = 0.37464 + 1.54226 * w - 0.26992 * w ^ 2

            Dim alpha, ai, bi

            alpha = (1 + c * (1 - (T / Tc) ^ 0.5)) ^ 2
            ai = a * alpha
            bi = 0.0778 * R * Tc / Pc

            Dim dadT = -a / T * (1 + c * (1 - (T / Tc) ^ 0.5)) * (c * (T / Tc) ^ 0.5)

            Dim AG1 = ai * P / (R * T) ^ 2
            Dim BG1 = bi * P / (R * T)

            Dim coeff(3) As Double

            coeff(0) = -AG1 * BG1 + BG1 ^ 2 + BG1 ^ 3
            coeff(1) = AG1 - 3 * BG1 ^ 2 - 2 * BG1
            coeff(2) = BG1 - 1
            coeff(3) = 1

            Dim temp1 = Poly_Roots(coeff)
            Dim tv = 0.0#

            If temp1(0, 0) > temp1(1, 0) Then
                tv = temp1(1, 0)
                temp1(1, 0) = temp1(0, 0)
                temp1(0, 0) = tv
            End If
            If temp1(0, 0) > temp1(2, 0) Then
                tv = temp1(2, 0)
                temp1(2, 0) = temp1(0, 0)
                temp1(0, 0) = tv
            End If
            If temp1(1, 0) > temp1(2, 0) Then
                tv = temp1(2, 0)
                temp1(2, 0) = temp1(1, 0)
                temp1(1, 0) = tv
            End If

            Dim Z = 0.0#

            If TYPE = "L" Then
                Z = temp1(0, 0)
            ElseIf TYPE = "V" Then
                Z = temp1(2, 0)
            End If

            Dim V = 0.0#
            If TYPE = "L" Then

                V = (Z * R * T / P) ' m3/mol

            ElseIf TYPE = "V" Then

                V = (Z * R * T / P) ' m3/mol

            End If

            Dim tmp1 = MM / V / 1000

            Dim DHres = R * T * (Z - 1) + (T * dadT - ai) / (2 ^ 1.5 * bi) * Math.Log((Z + 2.44 * BG1) / (Z - 0.414 * BG1))

            H_PR = DHres / MM

        End Function

        Function H_PR_MIX(ByVal TYPE As String, ByVal T As Double, ByVal P As Double, ByVal Vz As Object, ByVal VKij As Object, ByVal VTc As Object, ByVal VPc As Object, ByVal Vw As Object, ByVal VMM As Object, ByVal Hid As Double) As Double

            Dim H As Double = 0.0#
            H = H_PR_MIX_CPU(TYPE, T, P, Vz, VKij, VTc, VPc, Vw, VMM, Hid)

            Return H

        End Function

        Function H_PR_MIX_CPU(ByVal TYPE As String, ByVal T As Double, ByVal P As Double, ByVal Vz As Object, ByVal VKij As Object, ByVal VTc As Object, ByVal VPc As Object, ByVal Vw As Object, ByVal VMM As Object, ByVal Hid As Double) As Double

            Dim ai(), bi(), ci() As Double
            Dim n, R As Double
            Dim Tc(), Pc(), Vc(), w(), Zc(), alpha(), m(), a(,), b(,), Z, Tr() As Double
            Dim i, j, dadT

            n = Vz.Length - 1

            ReDim ai(n), bi(n), ci(n), a(n, n), b(n, n)
            ReDim Tc(n), Pc(n), Vc(n), Zc(n), w(n), alpha(n), m(n), Tr(n)

            R = 8.314

            i = 0
            Do
                Tc(i) = VTc(i)
                Tr(i) = T / Tc(i)
                Pc(i) = VPc(i)
                w(i) = Vw(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Dim MMm = 0.0#
            Do
                MMm += Vz(i) * VMM(i)
                i += 1
            Loop Until i = n + 1

            i = 0
            Do
                alpha(i) = (1 + (0.37464 + 1.54226 * w(i) - 0.26992 * w(i) ^ 2) * (1 - (T / Tc(i)) ^ 0.5)) ^ 2
                ai(i) = 0.45724 * alpha(i) * R ^ 2 * Tc(i) ^ 2 / Pc(i)
                bi(i) = 0.0778 * R * Tc(i) / Pc(i)
                ci(i) = 0.37464 + 1.54226 * w(i) - 0.26992 * w(i) ^ 2
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                j = 0
                Do
                    a(i, j) = (ai(i) * ai(j)) ^ 0.5 * (1 - VKij(i, j))
                    j = j + 1
                Loop Until j = n + 1
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Dim am = 0.0#
            Do
                j = 0
                Do
                    am = am + Vz(i) * Vz(j) * a(i, j)
                    j = j + 1
                Loop Until j = n + 1
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Dim bm = 0.0#
            Do
                bm = bm + Vz(i) * bi(i)
                i = i + 1
            Loop Until i = n + 1

            Dim AG1 = am * P / (R * T) ^ 2
            Dim BG1 = bm * P / (R * T)

            Dim coeff(3) As Double

            coeff(0) = -AG1 * BG1 + BG1 ^ 2 + BG1 ^ 3
            coeff(1) = AG1 - 3 * BG1 ^ 2 - 2 * BG1
            coeff(2) = BG1 - 1
            coeff(3) = 1

            Dim temp1 = Poly_Roots(coeff)
            Dim tv
            Dim tv2

            If Not IsNumeric(temp1) Then

                If temp1(0, 0) > temp1(1, 0) Then
                    tv = temp1(1, 0)
                    tv2 = temp1(1, 1)
                    temp1(1, 0) = temp1(0, 0)
                    temp1(0, 0) = tv
                    temp1(1, 1) = temp1(0, 1)
                    temp1(0, 1) = tv2
                End If
                If temp1(0, 0) > temp1(2, 0) Then
                    tv = temp1(2, 0)
                    temp1(2, 0) = temp1(0, 0)
                    temp1(0, 0) = tv
                    tv2 = temp1(2, 1)
                    temp1(2, 1) = temp1(0, 1)
                    temp1(0, 1) = tv2
                End If
                If temp1(1, 0) > temp1(2, 0) Then
                    tv = temp1(2, 0)
                    temp1(2, 0) = temp1(1, 0)
                    temp1(1, 0) = tv
                    tv2 = temp1(2, 1)
                    temp1(2, 1) = temp1(1, 1)
                    temp1(1, 1) = tv2
                End If

                If TYPE = "L" Then
                    Z = temp1(0, 0)
                    If temp1(0, 1) <> 0 Then
                        Z = temp1(1, 0)
                        If temp1(1, 1) <> 0 Then
                            Z = temp1(2, 0)
                        End If
                    End If
                    If Z < 0 Then Z = temp1(1, 0)
                ElseIf TYPE = "V" Then
                    Z = temp1(2, 0)
                    If temp1(2, 1) <> 0 Then
                        Z = temp1(1, 0)
                        If temp1(1, 1) <> 0 Then
                            Z = temp1(0, 0)
                        End If
                    End If
                End If

            Else

                Dim findZV, dfdz, zant As Double
                If TYPE = "V" Then Z = 1 Else Z = 0.05
                Do
                    findZV = coeff(3) * Z ^ 3 + coeff(2) * Z ^ 2 + coeff(1) * Z + coeff(0)
                    dfdz = 3 * coeff(3) * Z ^ 2 + 2 * coeff(2) * Z + coeff(1)
                    zant = Z
                    Z = Z - findZV / dfdz
                    If Z < 0 Then Z = 1
                Loop Until Math.Abs(findZV) < 0.0001 Or Double.IsNaN(Z)

            End If

            Dim V = (Z * R * T / P) ' m3/mol

            Dim tmp1 = MMm / V / 1000

            Dim aux1 = -R / 2 * (0.45724 / T) ^ 0.5
            i = 0
            Dim aux2 = 0.0#
            Do
                j = 0
                Do
                    aux2 += Vz(i) * Vz(j) * (1 - VKij(i, j)) * (ci(j) * (ai(i) * Tc(j) / Pc(j)) ^ 0.5 + ci(i) * (ai(j) * Tc(i) / Pc(i)) ^ 0.5)
                    j = j + 1
                Loop Until j = n + 1
                i = i + 1
            Loop Until i = n + 1

            dadT = aux1 * aux2

            Dim uu, ww As Double
            uu = 2
            ww = -1

            Dim DAres = am / (bm * (uu ^ 2 - 4 * ww) ^ 0.5) * Math.Log((2 * Z + BG1 * (uu - (uu ^ 2 - 4 * ww) ^ 0.5)) / (2 * Z + BG1 * (uu + (uu ^ 2 - 4 * ww) ^ 0.5))) - R * T * Math.Log((Z - BG1) / Z) - R * T * Math.Log(Z)
            Dim V0 As Double = R * 298.15 / 101325
            Dim DSres = R * Math.Log((Z - BG1) / Z) + R * Math.Log(Z) - 1 / (8 ^ 0.5 * bm) * dadT * Math.Log((2 * Z + BG1 * (2 - 8 ^ 0.5)) / (2 * Z + BG1 * (2 + 8 ^ 0.5)))
            Dim DHres = DAres + T * (DSres) + R * T * (Z - 1)

            If MathEx.Common.Sum(Vz) = 0.0# Then
                Return 0.0#
            Else
                Return Hid + DHres / MMm '/ 1000
            End If

        End Function

        Function S_PR_MIX(ByVal TYPE As String, ByVal T As Double, ByVal P As Double, ByVal Vz As Array, ByVal VKij As Object, ByVal VTc As Array, ByVal VPc As Array, ByVal Vw As Array, ByVal VMM As Array, ByVal Sid As Double) As Double

            Dim ai(), bi(), ci() As Double
            Dim n, R As Double
            Dim Tc(), Pc(), Vc(), w(), Zc(), alpha(), m(), a(,), b(,), Z, Tr() As Double
            Dim i, j, dadT

            n = Vz.Length - 1

            ReDim ai(n), bi(n), ci(n), a(n, n), b(n, n)
            ReDim Tc(n), Pc(n), Vc(n), Zc(n), w(n), alpha(n), m(n), Tr(n)

            R = 8.314

            i = 0
            Do
                Tc(i) = VTc(i)
                Tr(i) = T / Tc(i)
                Pc(i) = VPc(i)
                w(i) = Vw(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Dim MMm = 0.0#
            Do
                MMm += Vz(i) * VMM(i)
                i += 1
            Loop Until i = n + 1

            i = 0
            Do
                alpha(i) = (1 + (0.37464 + 1.54226 * w(i) - 0.26992 * w(i) ^ 2) * (1 - (T / Tc(i)) ^ 0.5)) ^ 2
                ai(i) = 0.45724 * alpha(i) * R ^ 2 * Tc(i) ^ 2 / Pc(i)
                bi(i) = 0.0778 * R * Tc(i) / Pc(i)
                ci(i) = 0.37464 + 1.54226 * w(i) - 0.26992 * w(i) ^ 2
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                j = 0
                Do
                    a(i, j) = (ai(i) * ai(j)) ^ 0.5 * (1 - VKij(i, j))
                    j = j + 1
                Loop Until j = n + 1
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Dim am = 0.0#
            Do
                j = 0
                Do
                    am = am + Vz(i) * Vz(j) * a(i, j)
                    j = j + 1
                Loop Until j = n + 1
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Dim bm = 0.0#
            Do
                bm = bm + Vz(i) * bi(i)
                i = i + 1
            Loop Until i = n + 1

            Dim AG1 = am * P / (R * T) ^ 2
            Dim BG1 = bm * P / (R * T)

            Dim coeff(3) As Double

            coeff(0) = -AG1 * BG1 + BG1 ^ 2 + BG1 ^ 3
            coeff(1) = AG1 - 3 * BG1 ^ 2 - 2 * BG1
            coeff(2) = BG1 - 1
            coeff(3) = 1

            Dim temp1 = Poly_Roots(coeff)
            Dim tv = 0.0#
            Dim tv2 = 0.0#
            If Not IsNumeric(temp1) Then

                If temp1(0, 0) > temp1(1, 0) Then
                    tv = temp1(1, 0)
                    tv2 = temp1(1, 1)
                    temp1(1, 0) = temp1(0, 0)
                    temp1(0, 0) = tv
                    temp1(1, 1) = temp1(0, 1)
                    temp1(0, 1) = tv2
                End If
                If temp1(0, 0) > temp1(2, 0) Then
                    tv = temp1(2, 0)
                    temp1(2, 0) = temp1(0, 0)
                    temp1(0, 0) = tv
                    tv2 = temp1(2, 1)
                    temp1(2, 1) = temp1(0, 1)
                    temp1(0, 1) = tv2
                End If
                If temp1(1, 0) > temp1(2, 0) Then
                    tv = temp1(2, 0)
                    temp1(2, 0) = temp1(1, 0)
                    temp1(1, 0) = tv
                    tv2 = temp1(2, 1)
                    temp1(2, 1) = temp1(1, 1)
                    temp1(1, 1) = tv2
                End If

                If TYPE = "L" Then
                    Z = temp1(0, 0)
                    If temp1(0, 1) <> 0 Then
                        Z = temp1(1, 0)
                        If temp1(1, 1) <> 0 Then
                            Z = temp1(2, 0)
                        End If
                    End If
                    If Z < 0 Then Z = temp1(1, 0)
                ElseIf TYPE = "V" Then
                    Z = temp1(2, 0)
                    If temp1(2, 1) <> 0 Then
                        Z = temp1(1, 0)
                        If temp1(1, 1) <> 0 Then
                            Z = temp1(0, 0)
                        End If
                    End If
                End If
            Else

                Dim findZV, dfdz, zant As Double
                If TYPE = "V" Then Z = 1 Else Z = 0.05
                Do
                    findZV = coeff(3) * Z ^ 3 + coeff(2) * Z ^ 2 + coeff(1) * Z + coeff(0)
                    dfdz = 3 * coeff(3) * Z ^ 2 + 2 * coeff(2) * Z + coeff(1)
                    zant = Z
                    Z = Z - findZV / dfdz
                    If Z < 0 Then Z = 1
                Loop Until Math.Abs(findZV) < 0.0001 Or Double.IsNaN(Z)

            End If

            Dim V = (Z * R * T / P) ' m3/mol

            Dim tmp1 = MMm / V / 1000

            Dim aux1 = -R / 2 * (0.45724 / T) ^ 0.5
            i = 0
            Dim aux2 = 0.0#
            Do
                j = 0
                Do
                    aux2 += Vz(i) * Vz(j) * (1 - VKij(i, j)) * (ci(j) * (ai(i) * Tc(j) / Pc(j)) ^ 0.5 + ci(i) * (ai(j) * Tc(i) / Pc(i)) ^ 0.5)
                    j = j + 1
                Loop Until j = n + 1
                i = i + 1
            Loop Until i = n + 1

            dadT = aux1 * aux2

            Dim V0 As Double = R * 298.15 / 101325
            Dim DSres = R * Math.Log((Z - BG1) / Z) + R * Math.Log(Z) - 1 / (8 ^ 0.5 * bm) * dadT * Math.Log((2 * Z + BG1 * (2 - 8 ^ 0.5)) / (2 * Z + BG1 * (2 + 8 ^ 0.5)))

            If MathEx.Common.Sum(Vz) = 0.0# Then
                S_PR_MIX = 0.0#
            Else
                S_PR_MIX = Sid + DSres / MMm '/ 1000
            End If

        End Function

        Function G_PR_MIX(ByVal TYPE As String, ByVal T As Double, ByVal P As Double, ByVal Vz As Array, ByVal VKij As Object, ByVal VTc As Array, ByVal VPc As Array, ByVal Vw As Array, ByVal VMM As Array, ByVal Sid As Double, ByVal Hid As Double) As Double

            Dim h As Double = H_PR_MIX(TYPE, T, P, Vz, VKij, VTc, VPc, Vw, VMM, Hid)
            Dim s As Double = S_PR_MIX(TYPE, T, P, Vz, VKij, VTc, VPc, Vw, VMM, Sid)

            Return h - T * s

        End Function

        Function ESTIMATE_V(ByVal Vz As Object, ByVal KI As Object) As Double

            Dim n = Vz.Length - 1

            Dim i As Integer

            Dim Vinf, Vsup As Double

            Dim fV, fV_inf, nsub, delta_V As Double

            Vinf = 0
            Vsup = 1

            nsub = 20

            delta_V = (Vsup - Vinf) / nsub

            i = 0
            Do
                i = i + 1
                fV = OF_V(Vinf, Vz, KI)
                Vinf = Vinf + delta_V
                fV_inf = OF_V(Vinf, Vz, KI)
            Loop Until fV * fV_inf < 0 Or Vinf > 1
            Vsup = Vinf
            Vinf = Vinf - delta_V

            'Brent method to find Vc

            Dim aaa, bbb, ccc, ddd, eee, min11, min22, faa, fbb, fcc, ppp, qqq, rrr, sss, tol11, xmm As Double
            Dim ITMAX2 As Integer = 100
            Dim iter2 As Integer

            aaa = Vinf
            bbb = Vsup
            ccc = Vsup

            faa = OF_V(aaa, Vz, KI)
            fbb = OF_V(bbb, Vz, KI)
            fcc = fbb

            iter2 = 0
            Do
                If (fbb > 0 And fcc > 0) Or (fbb < 0 And fcc < 0) Then
                    ccc = aaa
                    fcc = faa
                    ddd = bbb - aaa
                    eee = ddd
                End If
                If Math.Abs(fcc) < Math.Abs(fbb) Then
                    aaa = bbb
                    bbb = ccc
                    ccc = aaa
                    faa = fbb
                    fbb = fcc
                    fcc = faa
                End If
                tol11 = 0.000001
                xmm = 0.5 * (ccc - bbb)
                If (Math.Abs(xmm) <= tol11) Or (fbb = 0) Then GoTo Final3
                If (Math.Abs(eee) >= tol11) And (Math.Abs(faa) > Math.Abs(fbb)) Then
                    sss = fbb / faa
                    If aaa = ccc Then
                        ppp = 2 * xmm * sss
                        qqq = 1 - sss
                    Else
                        qqq = faa / fcc
                        rrr = fbb / fcc
                        ppp = sss * (2 * xmm * qqq * (qqq - rrr) - (bbb - aaa) * (rrr - 1))
                        qqq = (qqq - 1) * (rrr - 1) * (sss - 1)
                    End If
                    If ppp > 0 Then qqq = -qqq
                    ppp = Math.Abs(ppp)
                    min11 = 3 * xmm * qqq - Math.Abs(tol11 * qqq)
                    min22 = Math.Abs(eee * qqq)
                    Dim tvar2 As Double
                    If min11 < min22 Then tvar2 = min11
                    If min11 > min22 Then tvar2 = min22
                    If 2 * ppp < tvar2 Then
                        eee = ddd
                        ddd = ppp / qqq
                    Else
                        ddd = xmm
                        eee = ddd
                    End If
                Else
                    ddd = xmm
                    eee = ddd
                End If
                aaa = bbb
                faa = fbb
                If (Math.Abs(ddd) > tol11) Then
                    bbb += ddd
                Else
                    bbb += Math.Sign(xmm) * tol11
                End If
                fbb = OF_V(bbb, Vz, KI)
                iter2 += 1
            Loop Until iter2 = ITMAX2

Final2:     'bbb = -100

Final3:

            Return bbb

        End Function

        Function OF_V(ByVal V As Double, ByVal Vz As Object, ByVal KI As Object) As Double

            Dim i As Integer
            Dim n = Vz.Length - 1
            Dim result As Double

            i = 0
            Do
                result += Vz(i) * (1 - KI(i)) / (1 - V + V * KI(i))
                i = i + 1
            Loop Until i = n + 1

            Return result

        End Function

        Function GeneratePseudoRoot(ByVal T, ByVal P, ByVal Vx, ByVal VKij, ByVal VTc, ByVal VPc, ByVal Vw, ByVal VTb, ByVal TYPE)

            Dim ai(), bi(), aml2(), amv2() As Double
            Dim n, R, coeff(3), tmp() As Double
            Dim Tc(), Pc(), W(), alpha(), Vant(0, 4), m(), a(,), b(,), Tr() As Double
            Dim beta As Double
            Dim criterionOK As Boolean = False
            Dim hbcIndex, counter As Integer
            Dim sum_x As Double
            Dim ZV As Double
            Dim AG, BG, aml, bml As Double

            n = Vx.Length - 1

            ReDim ai(n), bi(n), tmp(n + 1), a(n, n), b(n, n)
            ReDim aml2(n), amv2(n)
            ReDim Tc(n), Pc(n), W(n), alpha(n), m(n), Tr(n)

            R = 8.314

            Dim i, j As Integer
            i = 0
            Do
                Tc(i) = VTc(i)
                Tr(i) = T / Tc(i)
                Pc(i) = VPc(i)
                W(i) = Vw(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                alpha(i) = (1 + (0.37464 + 1.54226 * W(i) - 0.26992 * W(i) ^ 2) * (1 - (T / Tc(i)) ^ 0.5)) ^ 2
                ai(i) = 0.45724 * alpha(i) * R ^ 2 * Tc(i) ^ 2 / Pc(i)
                bi(i) = 0.0778 * R * Tc(i) / Pc(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                j = 0
                Do
                    a(i, j) = (ai(i) * ai(j)) ^ 0.5 * (1 - VKij(i, j))
                    j = j + 1
                Loop Until j = n + 1
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                aml2(i) = 0
                i = i + 1
            Loop Until i = n + 1

            counter = 0
            Do

                i = 0
                aml = 0
                Do
                    j = 0
                    Do
                        aml = aml + Vx(i) * Vx(j) * a(i, j)
                        aml2(i) = aml2(i) + Vx(j) * a(j, i)
                        j = j + 1
                    Loop Until j = n + 1
                    i = i + 1
                Loop Until i = n + 1

                i = 0
                bml = 0
                Do
                    bml = bml + Vx(i) * bi(i)
                    i = i + 1
                Loop Until i = n + 1

                AG = aml * P / (R * T) ^ 2
                BG = bml * P / (R * T)

                coeff(0) = -AG * BG + BG ^ 2 + BG ^ 3
                coeff(1) = AG - 3 * BG ^ 2 - 2 * BG
                coeff(2) = BG - 1
                coeff(3) = 1

                Dim temp1 = Poly_Roots(coeff)
                Dim tv = 0.0#
                Dim tv2

                If Not IsNumeric(temp1) Then

                    If temp1(0, 0) > temp1(1, 0) Then
                        tv = temp1(1, 0)
                        temp1(1, 0) = temp1(0, 0)
                        temp1(0, 0) = tv
                        tv2 = temp1(1, 1)
                        temp1(1, 1) = temp1(0, 1)
                        temp1(0, 1) = tv2
                    End If
                    If temp1(0, 0) > temp1(2, 0) Then
                        tv = temp1(2, 0)
                        temp1(2, 0) = temp1(0, 0)
                        temp1(0, 0) = tv
                        tv2 = temp1(2, 1)
                        temp1(2, 1) = temp1(0, 1)
                        temp1(0, 1) = tv2
                    End If
                    If temp1(1, 0) > temp1(2, 0) Then
                        tv = temp1(2, 0)
                        temp1(2, 0) = temp1(1, 0)
                        temp1(1, 0) = tv
                        tv2 = temp1(2, 1)
                        temp1(2, 1) = temp1(1, 1)
                        temp1(1, 1) = tv2
                    End If

                    ZV = temp1(2, 0)
                    If temp1(2, 1) <> 0 Then
                        ZV = temp1(1, 0)
                        If temp1(1, 1) <> 0 Then
                            ZV = temp1(0, 0)
                        End If
                    End If

                    ZV = 0
                    If TYPE = "L" Then
                        ZV = temp1(0, 0)
                    ElseIf TYPE = "V" Then
                        ZV = temp1(2, 0)
                    End If

                Else

                    Dim findZV, dfdz, zant As Double
                    If TYPE = "V" Then ZV = 1 Else ZV = 0.05
                    Do
                        findZV = coeff(3) * ZV ^ 3 + coeff(2) * ZV ^ 2 + coeff(1) * ZV + coeff(0)
                        dfdz = 3 * coeff(3) * ZV ^ 2 + 2 * coeff(2) * ZV + coeff(1)
                        zant = ZV
                        ZV = ZV - findZV / dfdz
                        If ZV < 0 Then ZV = 1
                    Loop Until Math.Abs(findZV) < 0.0001 Or Double.IsNaN(ZV)

                End If


                beta = 1 / P * (1 - (BG * ZV ^ 2 + AG * ZV - 6 * BG ^ 2 * ZV - 2 * BG * ZV - 2 * AG * BG + 2 * BG ^ 2 + 2 * BG) / (ZV * (3 * ZV ^ 2 - 2 * ZV + 2 * BG * ZV + AG - 3 * BG ^ 2 - 2 * BG)))

                If TYPE = "L" Then
                    If beta < 0.005 / 101325 Then criterionOK = True
                Else
                    If beta < 3 / (P / 101325) And beta > 0.9 / (P / 101325) Then criterionOK = True
                    If ZV > 0.8 Then criterionOK = True
                End If

                If Not criterionOK Then
                    If TYPE = "L" Then
                        'check which component is the heaviest
                        i = 1
                        'hbc index is the heaviest component of the index
                        hbcIndex = i
                        i = 0
                        Do
                            If VTb(i) > VTb(hbcIndex) And Vx(i) <> 0 Then
                                hbcIndex = i
                            End If
                            i += 1
                        Loop Until i = n + 1
                        'increases the mole fraction of component hbc
                        Vx(hbcIndex) += 1
                        'And then normalizes the composition.
                        i = 0
                        sum_x = 0
                        Do
                            sum_x = sum_x + Vx(i)
                            i = i + 1
                        Loop Until i = n + 1
                        i = 0
                        Do
                            Vx(i) = Vx(i) / sum_x
                            i = i + 1
                        Loop Until i = n + 1
                    Else
                        P = P * 0.75
                    End If
                End If

                If P <= 1000 Then
                    Return Nothing
                End If

                counter += 1

            Loop Until criterionOK = True Or counter > 50

            Return New Object() {ZV, AG, BG, aml, bml}

        End Function

        Function CalcLnFug(ByVal T, ByVal P, ByVal Vx, ByVal VKij, ByVal VTc, ByVal VPc, ByVal Vw, ByVal VTb, ByVal TYPE)

            Dim n, R, coeff(3) As Double
            Dim Vant(0, 4) As Double
            Dim beta As Double
            Dim criterionOK As Boolean = False
            Dim ZV As Double
            Dim AG, BG, aml, bml As Double
            Dim t1, t2, t3, t4, t5 As Double

            n = Vx.Length - 1

            Dim ai(n), bi(n), tmp(n + 1), a(n, n), b(n, n)
            Dim aml2(n), amv2(n), LN_CF(n), PHI(n)
            Dim Tc(n), Pc(n), W(n), alpha(n), m(n), Tr(n)
            Dim rho, rho0, rho_mc, Tmc, dPdrho, dPdrho_, Zcalc As Double

            R = 8.314

            Dim i, j As Integer
            i = 0
            Do
                Tc(i) = VTc(i)
                Tr(i) = T / Tc(i)
                Pc(i) = VPc(i)
                W(i) = Vw(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                alpha(i) = (1 + (0.37464 + 1.54226 * W(i) - 0.26992 * W(i) ^ 2) * (1 - (T / Tc(i)) ^ 0.5)) ^ 2
                ai(i) = 0.45724 * alpha(i) * R ^ 2 * Tc(i) ^ 2 / Pc(i)
                bi(i) = 0.0778 * R * Tc(i) / Pc(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                j = 0
                Do
                    a(i, j) = (ai(i) * ai(j)) ^ 0.5 * (1 - VKij(i, j))
                    j = j + 1
                Loop Until j = n + 1
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                aml2(i) = 0
                i = i + 1
            Loop Until i = n + 1

            i = 0
            aml = 0
            Do
                j = 0
                Do
                    aml = aml + Vx(i) * Vx(j) * a(i, j)
                    aml2(i) = aml2(i) + Vx(j) * a(j, i)
                    j = j + 1
                Loop Until j = n + 1
                i = i + 1
            Loop Until i = n + 1

            i = 0
            bml = 0
            Do
                bml = bml + Vx(i) * bi(i)
                i = i + 1
            Loop Until i = n + 1

            AG = aml * P / (R * T) ^ 2
            BG = bml * P / (R * T)

            coeff(0) = -AG * BG + BG ^ 2 + BG ^ 3
            coeff(1) = AG - 3 * BG ^ 2 - 2 * BG
            coeff(2) = BG - 1
            coeff(3) = 1

            Dim temp1 = Poly_Roots(coeff)
            Dim tv = 0.0#
            Dim tv2

            If Not IsNumeric(temp1) Then

                If temp1(0, 0) > temp1(1, 0) Then
                    tv = temp1(1, 0)
                    temp1(1, 0) = temp1(0, 0)
                    temp1(0, 0) = tv
                    tv2 = temp1(1, 1)
                    temp1(1, 1) = temp1(0, 1)
                    temp1(0, 1) = tv2
                End If
                If temp1(0, 0) > temp1(2, 0) Then
                    tv = temp1(2, 0)
                    temp1(2, 0) = temp1(0, 0)
                    temp1(0, 0) = tv
                    tv2 = temp1(2, 1)
                    temp1(2, 1) = temp1(0, 1)
                    temp1(0, 1) = tv2
                End If
                If temp1(1, 0) > temp1(2, 0) Then
                    tv = temp1(2, 0)
                    temp1(2, 0) = temp1(1, 0)
                    temp1(1, 0) = tv
                    tv2 = temp1(2, 1)
                    temp1(2, 1) = temp1(1, 1)
                    temp1(1, 1) = tv2
                End If

                ZV = temp1(2, 0)
                If temp1(2, 1) <> 0 Then
                    ZV = temp1(1, 0)
                    If temp1(1, 1) <> 0 Then
                        ZV = temp1(0, 0)
                    End If
                End If

                ZV = 0
                If TYPE = "L" Then
                    ZV = temp1(0, 0)
                ElseIf TYPE = "V" Then
                    ZV = temp1(2, 0)
                End If

            Else

                Dim findZV, dfdz, zant As Double
                If TYPE = "V" Then ZV = 1 Else ZV = 0.05
                Do
                    findZV = coeff(3) * ZV ^ 3 + coeff(2) * ZV ^ 2 + coeff(1) * ZV + coeff(0)
                    dfdz = 3 * coeff(3) * ZV ^ 2 + 2 * coeff(2) * ZV + coeff(1)
                    zant = ZV
                    ZV = ZV - findZV / dfdz
                    If ZV < 0 Then ZV = 1
                Loop Until Math.Abs(findZV) < 0.0001 Or Double.IsNaN(ZV)

            End If

            beta = 1 / P * (1 - (BG * ZV ^ 2 + AG * ZV - 6 * BG ^ 2 * ZV - 2 * BG * ZV - 2 * AG * BG + 2 * BG ^ 2 + 2 * BG) / (ZV * (3 * ZV ^ 2 - 2 * ZV + 2 * BG * ZV + AG - 3 * BG ^ 2 - 2 * BG)))

            rho0 = 1 / bml
            rho_mc = 0.2599 / bml
            Tmc = 0.20268 * aml / (R * bml)
            rho = P / (ZV * R * T)
            dPdrho_ = 0.1 * R * T
            dPdrho = bml * rho * R * T * (1 - bml * rho) ^ -2 + R * T * (1 - bml * rho) ^ -1 + _
                    aml * rho ^ 2 * (1 + 2 * bml * rho - (bml * rho) ^ 2) ^ -2 * (2 * bml - 2 * bml ^ 2 * rho) + _
                    2 * aml * rho * (1 + 2 * bml * rho - (bml * rho) ^ 2) ^ -1

            If TYPE = "L" Then
                Zcalc = ZV
                i = 0
                Do
                    t1 = bi(i) * (Zcalc - 1) / bml
                    t2 = -Math.Log(Zcalc - BG)
                    t3 = AG * (2 * aml2(i) / aml - bi(i) / bml)
                    t4 = Math.Log((Zcalc + (1 + 2 ^ 0.5) * BG) / (Zcalc + (1 - 2 ^ 0.5) * BG))
                    t5 = 2 * 2 ^ 0.5 * BG
                    LN_CF(i) = t1 + t2 - (t3 * t4 / t5)
                    LN_CF(i) = LN_CF(i) '* Pcalc / P
                    i = i + 1
                Loop Until i = n + 1
                Return LN_CF
            Else
                Zcalc = ZV
                i = 0
                Do
                    t1 = bi(i) * (Zcalc - 1) / bml
                    t2 = -Math.Log(Zcalc - BG)
                    t3 = AG * (2 * aml2(i) / aml - bi(i) / bml)
                    t4 = Math.Log((Zcalc + (1 + 2 ^ 0.5) * BG) / (Zcalc + (1 - 2 ^ 0.5) * BG))
                    t5 = 2 * 2 ^ 0.5 * BG
                    LN_CF(i) = t1 + t2 - (t3 * t4 / t5)
                    LN_CF(i) = LN_CF(i)
                    i = i + 1
                Loop Until i = n + 1
                Return LN_CF
            End If

        End Function

        Function CalcPartialVolume(ByVal T, ByVal P, ByVal Vx, ByVal VKij, ByVal VTc, ByVal VPc, ByVal Vw, ByVal VTb, ByVal TYPE, ByVal deltaP)

            Dim lnfug1, lnfug2 As Object
            Dim P1, P2 As Double
            P1 = P
            P2 = P + deltaP

            lnfug1 = Me.CalcLnFug(T, P1, Vx, VKij, VTc, VPc, Vw, VTb, TYPE)
            lnfug2 = Me.CalcLnFug(T, P2, Vx, VKij, VTc, VPc, Vw, VTb, TYPE)

            Dim i As Integer
            Dim n As Integer = UBound(lnfug1)

            Dim partvol(n) As Double

            i = 0
            For i = 0 To n
                partvol(i) = (Math.Log(Math.Exp(lnfug2(i)) * Vx(i) * P2) - Math.Log(Math.Exp(lnfug1(i)) * Vx(i) * P1)) / deltaP * (8.314 * T) 'm3/mol
                If Double.IsNaN(partvol(i)) Then partvol(i) = 0
            Next

            Return partvol

        End Function

        Function ESTIMATE_RhoLim(ByVal am As Double, ByVal bm As Double, ByVal T As Double, ByVal P As Double) As Double

            Dim i As Integer

            Dim rinf, rsup As Double

            Dim fr, fr_inf, nsub, delta_r As Double

            rinf = 0
            rsup = P / (8.314 * T)

            nsub = 10

            delta_r = (rsup - rinf) / nsub

            i = 0
            Do
                i = i + 1
                fr = OF_Rho(rinf, am, bm, T)
                rinf = rinf + delta_r
                fr_inf = OF_Rho(rinf, am, bm, T)
            Loop Until fr * fr_inf < 0 Or i = 11
            If i = 11 Then GoTo Final2
            rsup = rinf
            rinf = rinf - delta_r

            'Brent method to find Vc

            Dim aaa, bbb, ccc, ddd, eee, min11, min22, faa, fbb, fcc, ppp, qqq, rrr, sss, tol11, xmm As Double
            Dim ITMAX2 As Integer = 100
            Dim iter2 As Integer

            aaa = rinf
            bbb = rsup
            ccc = rsup

            faa = OF_Rho(aaa, am, bm, T)
            fbb = OF_Rho(bbb, am, bm, T)
            fcc = fbb

            iter2 = 0
            Do
                If (fbb > 0 And fcc > 0) Or (fbb < 0 And fcc < 0) Then
                    ccc = aaa
                    fcc = faa
                    ddd = bbb - aaa
                    eee = ddd
                End If
                If Math.Abs(fcc) < Math.Abs(fbb) Then
                    aaa = bbb
                    bbb = ccc
                    ccc = aaa
                    faa = fbb
                    fbb = fcc
                    fcc = faa
                End If
                tol11 = 0.000001
                xmm = 0.5 * (ccc - bbb)
                If (Math.Abs(xmm) <= tol11) Or (fbb = 0) Then GoTo Final3
                If (Math.Abs(eee) >= tol11) And (Math.Abs(faa) > Math.Abs(fbb)) Then
                    sss = fbb / faa
                    If aaa = ccc Then
                        ppp = 2 * xmm * sss
                        qqq = 1 - sss
                    Else
                        qqq = faa / fcc
                        rrr = fbb / fcc
                        ppp = sss * (2 * xmm * qqq * (qqq - rrr) - (bbb - aaa) * (rrr - 1))
                        qqq = (qqq - 1) * (rrr - 1) * (sss - 1)
                    End If
                    If ppp > 0 Then qqq = -qqq
                    ppp = Math.Abs(ppp)
                    min11 = 3 * xmm * qqq - Math.Abs(tol11 * qqq)
                    min22 = Math.Abs(eee * qqq)
                    Dim tvar2 As Double
                    If min11 < min22 Then tvar2 = min11
                    If min11 > min22 Then tvar2 = min22
                    If 2 * ppp < tvar2 Then
                        eee = ddd
                        ddd = ppp / qqq
                    Else
                        ddd = xmm
                        eee = ddd
                    End If
                Else
                    ddd = xmm
                    eee = ddd
                End If
                aaa = bbb
                faa = fbb
                If (Math.Abs(ddd) > tol11) Then
                    bbb += ddd
                Else
                    bbb += Math.Sign(xmm) * tol11
                End If
                fbb = OF_Rho(bbb, am, bm, T)
                iter2 += 1
            Loop Until iter2 = ITMAX2

Final2:     bbb = -100

Final3:

            Return bbb

        End Function

        Function OF_Rho(ByVal rho As Double, ByVal aml As Double, ByVal bml As Double, ByVal T As Double) As Double

            Dim R As Double = 8.314
            Return 0.1 * 8.314 * T - _
                        bml * rho * R * T * (1 - bml * rho) ^ -2 + R * T * (1 - bml * rho) ^ -1 + _
                        aml * rho ^ 2 * (1 + 2 * bml * rho - (bml * rho) ^ 2) ^ -2 * (2 * bml - 2 * bml ^ 2 * rho) + _
                        2 * aml * rho * (1 + 2 * bml * rho - (bml * rho) ^ 2) ^ -1

        End Function

    End Class

End Namespace

