'    Soave-Redlich-Kwong Property Package 
'    Copyright 2008-2014 Daniel Wagner O. de Medeiros
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

Imports DTL.DTL.MathEx

Namespace DTL.SimulationObjects.PropertyPackages.ThermoPlugs

    <Serializable()> Public Class SRK

        Inherits DTL.SimulationObjects.PropertyPackages.ThermoPlug

        Shared Function ReturnParameters(ByVal T As Double, ByVal P As Double, ByVal Vx As Array, ByVal VKij As Object, ByVal VTc As Array, ByVal VPc As Array, ByVal Vw As Array)

            Dim n, R, coeff(3) As Double
            Dim Vant(0, 4) As Double
            Dim criterioOK As Boolean = False
            Dim AG, BG, aml, bml As Double

            n = UBound(Vx)

            Dim ai(n), bi(n), ci(n), tmp(n + 1), a(n, n), b(n, n)
            Dim aml2(n), amv2(n), LN_CF(n), PHI(n) As Double
            Dim Tc(n), Pc(n), W(n), alpha(n), m(n), Tr(n)

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
                alpha(i) = (1 + (0.48 + 1.574 * W(i) - 0.176 * W(i) ^ 2) * (1 - (T / Tc(i)) ^ 0.5)) ^ 2
                ai(i) = 0.42748 * alpha(i) * R ^ 2 * Tc(i) ^ 2 / Pc(i)
                bi(i) = 0.08664 * R * Tc(i) / Pc(i)
                ci(i) = 0.48 + 1.574 * W(i) - 0.176 * W(i) ^ 2
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

            Dim _zarray As ArrayList, _mingz As Object, Z As Double
            _zarray = CalcZ(T, P, Vx, VKij, VTc, VPc, Vw)
            _mingz = ZtoMinG(_zarray.ToArray, T, P, Vx, VKij, VTc, VPc, Vw)
            Z = _zarray(_mingz(0))

            Dim aux1 = -R / 2 * (0.45724 / T) ^ 0.5
            i = 0
            Dim aux2 = 0.0#
            Do
                j = 0
                Do
                    aux2 += Vx(i) * Vx(j) * (1 - VKij(i, j)) * (ci(j) * (ai(i) * Tc(j) / Pc(j)) ^ 0.5 + ci(i) * (ai(j) * Tc(i) / Pc(i)) ^ 0.5)
                    j = j + 1
                Loop Until j = n + 1
                i = i + 1
            Loop Until i = n + 1

            Dim dadT = aux1 * aux2

            Return New Double() {aml, bml, Z * R * T / P, dadT}

        End Function

        Shared Function ZtoMinG(ByVal Z_ As Array, ByVal T As Double, ByVal P As Double, ByVal Vz As Array, ByVal VKij As Object, ByVal VTc As Array, ByVal VPc As Array, ByVal Vw As Array) As Object

            App.WriteToConsole("SRK min-G root finder (Z) for T = " & T & " K, P = " & P & " Pa and Z = " & DirectCast(Z_, Object()).ToArrayString, 3)

            Dim S, H, Z As Double

            Dim ai(), bi(), ci() As Double
            Dim n, R, dadT As Double
            Dim Tc(), Pc(), Vc(), w(), Zc(), alpha(), m(), a(,), b(,), Tr() As Double
            Dim i, j, k, l As Integer

            n = UBound(Vz)

            Dim G(UBound(Z_)) As Double

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
            Do
                alpha(i) = (1 + (0.48 + 1.574 * w(i) - 0.176 * w(i) ^ 2) * (1 - (T / Tc(i)) ^ 0.5)) ^ 2
                ai(i) = 0.42748 * alpha(i) * R ^ 2 * Tc(i) ^ 2 / Pc(i)
                bi(i) = 0.08664 * R * Tc(i) / Pc(i)
                ci(i) = 0.48 + 1.574 * w(i) - 0.176 * w(i) ^ 2
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

            l = 0
            For Each Z In Z_

                Dim V = (Z * R * T / P) ' m3/mol

                Dim tmp1 = 1 / V / 1000

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
                uu = 1
                ww = 0

                Dim DAres = am / (bm * (uu ^ 2 - 4 * ww) ^ 0.5) * Math.Log((2 * Z + BG1 * (uu - (uu ^ 2 - 4 * ww) ^ 0.5)) / (2 * Z + BG1 * (uu + (uu ^ 2 - 4 * ww) ^ 0.5))) - R * T * Math.Log((Z - BG1) / Z) - R * T * Math.Log(V * 101325 / (R * 298.15))
                Dim V0 As Double = R * 298.15 / 101325
                Dim DSres = R * Math.Log((Z - BG1) / Z) + R * Math.Log(Z) - 1 / (8 ^ 0.5 * bm) * dadT * Math.Log((2 * Z + BG1 * (2 - 8 ^ 0.5)) / (2 * Z + BG1 * (2 + 8 ^ 0.5)))
                Dim DHres = DAres + T * (DSres) + R * T * (Z - 1)

                H = DHres

                S = DSres

                G(l) = H - T * S

                If j = 0 Then
                    k = 0
                Else
                    i = 0
                    Do
                        If G(l) <= G(k) Then k = l
                        i = i + 1
                    Loop Until i = UBound(G) + 1
                End If

                l = l + 1

            Next

            App.WriteToConsole("Result: Min-G Z Index = " & k, 3)

            Return New Object() {k, G(k)}

        End Function

        Public Overrides Function CalcLnFug(ByVal T As Double, ByVal P As Double, ByVal Vx As Array, ByVal VKij As Object, ByVal VTc As Array, ByVal VPc As Array, ByVal Vw As Array, Optional ByVal otherargs As Object = Nothing, Optional ByVal forcephase As String = "")

            Dim n, R, coeff(3) As Double
            Dim Vant(0, 4) As Double
            Dim criterioOK As Boolean = False
            Dim AG, BG, aml, bml As Double
            Dim t1, t2, t3, t4, t5 As Double

            n = UBound(Vx)

            Dim ai(n), bi(n), tmp(n + 1), a(n, n), b(n, n)
            Dim aml2(n), amv2(n), LN_CF(n), PHI(n) As Double
            Dim Tc(n), Pc(n), W(n), alpha(n), m(n), Tr(n)

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
                alpha(i) = (1 + (0.48 + 1.574 * W(i) - 0.176 * W(i) ^ 2) * (1 - (T / Tc(i)) ^ 0.5)) ^ 2
                ai(i) = 0.42748 * alpha(i) * R ^ 2 * Tc(i) ^ 2 / Pc(i)
                bi(i) = 0.08664 * R * Tc(i) / Pc(i)
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

            Dim _zarray As ArrayList, _mingz As Object, Z As Double


            _zarray = CalcZ(T, P, Vx, VKij, VTc, VPc, Vw)
            If forcephase <> "" Then
                If forcephase = "L" Then
                    Z = Common.Min(_zarray.ToArray())
                ElseIf forcephase = "V" Then
                    Z = Common.Max(_zarray.ToArray())
                End If
            Else
                _mingz = ZtoMinG(_zarray.ToArray, T, P, Vx, VKij, VTc, VPc, Vw)
                Z = _zarray(_mingz(0))
            End If

            Dim Pcorr As Double = P
            Dim ZP As Double() = CheckRoot(Z, aml, bml, P, T, forcephase)
            Z = ZP(0)
            Pcorr = ZP(1)

            i = 0
            Do
                t1 = bi(i) * (Z - 1) / bml
                t2 = -Math.Log(Z - BG)
                t3 = AG * (2 * aml2(i) / aml - bi(i) / bml)
                t4 = Math.Log((Z + BG) / Z)
                t5 = BG
                LN_CF(i) = t1 + t2 - (t3 * t4 / t5) + Math.Log(Pcorr / P)
                i = i + 1
            Loop Until i = n + 1

            Return LN_CF

        End Function

        Shared Function CalcZ(ByVal T, ByVal P, ByVal Vx, ByVal VKij, ByVal VTc, ByVal VPc, ByVal Vw) As ArrayList

            Dim ai(), bi(), aml2(), amv2() As Double
            Dim n, R, coeff(3), tmp() As Double
            Dim Tc(), Pc(), W(), alpha(), Vant(0, 4), m(), a(,), b(,), Tr() As Double

            n = UBound(Vx)

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
                alpha(i) = (1 + (0.48 + 1.574 * W(i) - 0.176 * W(i) ^ 2) * (1 - (T / Tc(i)) ^ 0.5)) ^ 2
                ai(i) = 0.42748 * alpha(i) * R ^ 2 * Tc(i) ^ 2 / Pc(i)
                bi(i) = 0.08664 * R * Tc(i) / Pc(i)
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

            coeff(0) = -AG * BG
            coeff(1) = AG - BG - BG ^ 2
            coeff(2) = -1
            coeff(3) = 1

            Dim temp1 = Poly_Roots(coeff)
            Dim tv = 0.0#
            Dim ZV, tv2

            Dim result As New ArrayList

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

                If temp1(0, 1) = 0.0# And temp1(0, 0) > 0.0# Then result.Add(temp1(0, 0))
                If temp1(1, 1) = 0.0# And temp1(1, 0) > 0.0# Then result.Add(temp1(1, 0))
                If temp1(2, 1) = 0.0# And temp1(2, 0) > 0.0# Then result.Add(temp1(2, 0))

            Else

                Dim findZV
                Dim lc As Integer = 0
                ZV = 1
                Do
                    findZV = coeff(3) * ZV ^ 3 + coeff(2) * ZV ^ 2 + coeff(1) * ZV + coeff(0)
                    ZV -= 0.00001
                    If ZV < 0 Then ZV = 1
                    lc += 1
                Loop Until Math.Abs(findZV) < 0.0001 Or lc > 10000

                result.Add(ZV)

                Return result

            End If

            Return result

        End Function

        Public Overrides Function PhaseType(ByVal T As Double, ByVal P As Double, ByVal Vx As Array, ByVal VKij As Object, ByVal VTc As Array, ByVal VPc As Array, ByVal Vw As Array, Optional ByVal otherargs As Object = Nothing)

            Dim n, R, coeff(3) As Double
            Dim Vant(0, 4) As Double
            Dim beta As Double
            Dim criterioOK As Boolean = False
            Dim AG, BG, aml, bml As Double

            n = UBound(Vx)

            Dim ai(n), bi(n), tmp(n + 1), a(n, n), b(n, n)
            Dim aml2(n), amv2(n), LN_CF(n), PHI(n) As Double
            Dim Tc(n), Pc(n), W(n), alpha(n), m(n), Tr(n)

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
                alpha(i) = (1 + (0.48 + 1.574 * W(i) - 0.176 * W(i) ^ 2) * (1 - (T / Tc(i)) ^ 0.5)) ^ 2
                ai(i) = 0.42748 * alpha(i) * R ^ 2 * Tc(i) ^ 2 / Pc(i)
                bi(i) = 0.08664 * R * Tc(i) / Pc(i)
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

            Dim _zarray As ArrayList, _mingz As Object, Z As Double

            _zarray = CalcZ(T, P, Vx, VKij, VTc, VPc, Vw)
            _mingz = ZtoMinG(_zarray.ToArray, T, P, Vx, VKij, VTc, VPc, Vw)
            Z = _zarray(_mingz(0))

            beta = 1 / P * (1 - (2 * AG * BG + Z * BG + 2 * BG ^ 2 * Z - AG * Z) / (Z * (3 * Z ^ 2 - 2 * Z + AG - BG - BG ^ 2)))
            Dim phase As String = "Unknown"

            If beta < 0.005 / 101322 Then phase = "L" Else phase = "V"
            'If beta > 0.9 / P And beta < 3 / P Then phase = "V"

            Return New Object() {phase, beta}

        End Function

        Public Overrides Function CalcEnthalpy(ByVal phasetype As String, ByVal T As Double, ByVal P As Double, ByVal Vz As System.Array, ByVal VKij As Object, ByVal VTc As System.Array, ByVal VPc As System.Array, ByVal Vw As System.Array, ByVal Hid As Double, Optional ByVal otherargs As Object = Nothing) As Double

            Dim ai(), bi(), ci() As Double
            Dim n, R As Double
            Dim Tc(), Pc(), Vc(), w(), Zc(), alpha(), m(), a(,), b(,), Z, Tr() As Double
            Dim i, j, dadT

            n = UBound(Vz)

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
            Do
                alpha(i) = (1 + (0.48 + 1.574 * w(i) - 0.176 * w(i) ^ 2) * (1 - (T / Tc(i)) ^ 0.5)) ^ 2
                ai(i) = 0.42748 * alpha(i) * R ^ 2 * Tc(i) ^ 2 / Pc(i)
                bi(i) = 0.08664 * R * Tc(i) / Pc(i)
                ci(i) = 0.48 + 1.574 * w(i) - 0.176 * w(i) ^ 2
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

            coeff(0) = -AG1 * BG1
            coeff(1) = AG1 - BG1 - BG1 ^ 2
            coeff(2) = -1
            coeff(3) = 1

            Dim temp1 = Poly_Roots(coeff)
            Dim tv
            Dim tv2
            Try

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

                If phasetype = "L" Then
                    Z = temp1(0, 0)
                    If temp1(0, 1) <> 0 Then
                        Z = temp1(1, 0)
                        If temp1(1, 1) <> 0 Then
                            Z = temp1(2, 0)
                        End If
                    End If
                ElseIf phasetype = "V" Then
                    Z = temp1(2, 0)
                    If temp1(2, 1) <> 0 Then
                        Z = temp1(1, 0)
                        If temp1(1, 1) <> 0 Then
                            Z = temp1(0, 0)
                        End If
                    End If
                End If
            Catch

                Dim findZ

                If phasetype = "V" Then

                    Z = 1
                    Do
                        findZ = coeff(3) * Z ^ 3 + coeff(2) * Z ^ 2 + coeff(1) * Z + coeff(0)
                        Z -= 0.00001
                        If Z < 0 Then Throw New Exception(App.GetLocalString("PropPack_ZError"))
                    Loop Until Math.Abs(findZ) < 0.0001

                Else

                    Z = 0
                    Do
                        findZ = coeff(3) * Z ^ 3 + coeff(2) * Z ^ 2 + coeff(1) * Z + coeff(0)
                        Z += 0.00001
                        If Z > 1 Then Throw New Exception(App.GetLocalString("PropPack_ZError"))
                    Loop Until Math.Abs(findZ) < 0.0001

                End If

            End Try

            Dim V = (Z * R * T / P) ' m3/mol

            Dim aux1 = -R / 2 * (0.42748 / T) ^ 0.5
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
            uu = 1
            ww = 0

            Dim DAres = am / (bm * (uu ^ 2 - 4 * ww) ^ 0.5) * Math.Log((2 * Z + BG1 * (uu - (uu ^ 2 - 4 * ww) ^ 0.5)) / (2 * Z + BG1 * (uu + (uu ^ 2 - 4 * ww) ^ 0.5))) - R * T * Math.Log((Z - BG1) / Z) - R * T * Math.Log(V * 101325 / (R * 298.15))
            Dim DSres = R * Math.Log((Z - BG1) / Z) + R * Math.Log(V * 101325 / (R * 298.15)) - 1 / (8 ^ 0.5 * bm) * dadT * Math.Log((2 * Z + BG1 * (2 - 8 ^ 0.5)) / (2 * Z + BG1 * (2 + 8 ^ 0.5)))
            Dim DHres = DAres + T * (DSres) + R * T * (Z - 1)

            Return Hid + DHres

        End Function

        Public Overrides Function CalcEntropy(ByVal phasetype As String, ByVal T As Double, ByVal P As Double, ByVal Vz As System.Array, ByVal VKij As Object, ByVal VTc As System.Array, ByVal VPc As System.Array, ByVal Vw As System.Array, ByVal Sid As Double, Optional ByVal otherargs As Object = Nothing) As Double

            Dim ai(), bi(), ci() As Double
            Dim n, R As Double
            Dim Tc(), Pc(), Vc(), w(), Zc(), alpha(), m(), a(,), b(,), Z, Tr() As Double
            Dim i, j, dadT

            n = UBound(Vz)

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
            Do
                alpha(i) = (1 + (0.48 + 1.574 * w(i) - 0.176 * w(i) ^ 2) * (1 - (T / Tc(i)) ^ 0.5)) ^ 2
                ai(i) = 0.42748 * alpha(i) * R ^ 2 * Tc(i) ^ 2 / Pc(i)
                bi(i) = 0.08664 * R * Tc(i) / Pc(i)
                ci(i) = 0.48 + 1.574 * w(i) - 0.176 * w(i) ^ 2
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

            coeff(0) = -AG1 * BG1
            coeff(1) = AG1 - BG1 - BG1 ^ 2
            coeff(2) = -1
            coeff(3) = 1

            Dim temp1 = Poly_Roots(coeff)
            Dim tv
            Dim tv2
            Try

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

                If phasetype = "L" Then
                    Z = temp1(0, 0)
                    If temp1(0, 1) <> 0 Then
                        Z = temp1(1, 0)
                        If temp1(1, 1) <> 0 Then
                            Z = temp1(2, 0)
                        End If
                    End If
                ElseIf phasetype = "V" Then
                    Z = temp1(2, 0)
                    If temp1(2, 1) <> 0 Then
                        Z = temp1(1, 0)
                        If temp1(1, 1) <> 0 Then
                            Z = temp1(0, 0)
                        End If
                    End If
                End If
            Catch

                Dim findZ

                If phasetype = "V" Then

                    Z = 1
                    Do
                        findZ = coeff(3) * Z ^ 3 + coeff(2) * Z ^ 2 + coeff(1) * Z + coeff(0)
                        Z -= 0.00001
                        If Z < 0 Then Throw New Exception(App.GetLocalString("PropPack_ZError"))
                    Loop Until Math.Abs(findZ) < 0.0001

                Else

                    Z = 0
                    Do
                        findZ = coeff(3) * Z ^ 3 + coeff(2) * Z ^ 2 + coeff(1) * Z + coeff(0)
                        Z += 0.00001
                        If Z > 1 Then Throw New Exception(App.GetLocalString("PropPack_ZError"))
                    Loop Until Math.Abs(findZ) < 0.0001

                End If

            End Try

            Dim V = (Z * R * T / P) ' m3/mol

            Dim aux1 = -R / 2 * (0.42748 / T) ^ 0.5
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
            uu = 1
            ww = 0

            Dim DAres = am / (bm * (uu ^ 2 - 4 * ww) ^ 0.5) * Math.Log((2 * Z + BG1 * (uu - (uu ^ 2 - 4 * ww) ^ 0.5)) / (2 * Z + BG1 * (uu + (uu ^ 2 - 4 * ww) ^ 0.5))) - R * T * Math.Log((Z - BG1) / Z) - R * T * Math.Log(V * 101325 / (R * 298.15))
            Dim DSres = R * Math.Log((Z - BG1) / Z) + R * Math.Log(V * 101325 / (R * 298.15)) - 1 / (8 ^ 0.5 * bm) * dadT * Math.Log((2 * Z + BG1 * (2 - 8 ^ 0.5)) / (2 * Z + BG1 * (2 + 8 ^ 0.5)))
            Dim DHres = DAres + T * (DSres) + R * T * (Z - 1)

            Return Sid + DSres

        End Function

        Public Overrides Function CalcGibbsEnergy(ByVal phasetype As String, ByVal T As Double, ByVal P As Double, ByVal Vz As System.Array, ByVal VKij As Object, ByVal VTc As System.Array, ByVal VPc As System.Array, ByVal Vw As System.Array, ByVal Gid As Double, Optional ByVal otherargs As Object = Nothing) As Double

            Dim ai(), bi(), ci() As Double
            Dim n, R As Double
            Dim Tc(), Pc(), Vc(), w(), Zc(), alpha(), m(), a(,), b(,), Z, Tr() As Double
            Dim i, j, dadT

            n = UBound(Vz)

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
            Do
                alpha(i) = (1 + (0.48 + 1.574 * w(i) - 0.176 * w(i) ^ 2) * (1 - (T / Tc(i)) ^ 0.5)) ^ 2
                ai(i) = 0.42748 * alpha(i) * R ^ 2 * Tc(i) ^ 2 / Pc(i)
                bi(i) = 0.08664 * R * Tc(i) / Pc(i)
                ci(i) = 0.48 + 1.574 * w(i) - 0.176 * w(i) ^ 2
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

            coeff(0) = -AG1 * BG1
            coeff(1) = AG1 - BG1 - BG1 ^ 2
            coeff(2) = -1
            coeff(3) = 1

            Dim temp1 = Poly_Roots(coeff)
            Dim tv
            Dim tv2
            Try

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

                If phasetype = "L" Then
                    Z = temp1(0, 0)
                    If temp1(0, 1) <> 0 Then
                        Z = temp1(1, 0)
                        If temp1(1, 1) <> 0 Then
                            Z = temp1(2, 0)
                        End If
                    End If
                ElseIf phasetype = "V" Then
                    Z = temp1(2, 0)
                    If temp1(2, 1) <> 0 Then
                        Z = temp1(1, 0)
                        If temp1(1, 1) <> 0 Then
                            Z = temp1(0, 0)
                        End If
                    End If
                End If
            Catch

                Dim findZ

                If phasetype = "V" Then

                    Z = 1
                    Do
                        findZ = coeff(3) * Z ^ 3 + coeff(2) * Z ^ 2 + coeff(1) * Z + coeff(0)
                        Z -= 0.00001
                        If Z < 0 Then Throw New Exception(App.GetLocalString("PropPack_ZError"))
                    Loop Until Math.Abs(findZ) < 0.0001

                Else

                    Z = 0
                    Do
                        findZ = coeff(3) * Z ^ 3 + coeff(2) * Z ^ 2 + coeff(1) * Z + coeff(0)
                        Z += 0.00001
                        If Z > 1 Then Throw New Exception(App.GetLocalString("PropPack_ZError"))
                    Loop Until Math.Abs(findZ) < 0.0001

                End If

            End Try

            Dim V = (Z * R * T / P) ' m3/mol

            Dim aux1 = -R / 2 * (0.42748 / T) ^ 0.5
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
            uu = 1
            ww = 0

            Dim DAres = am / (bm * (uu ^ 2 - 4 * ww) ^ 0.5) * Math.Log((2 * Z + BG1 * (uu - (uu ^ 2 - 4 * ww) ^ 0.5)) / (2 * Z + BG1 * (uu + (uu ^ 2 - 4 * ww) ^ 0.5))) - R * T * Math.Log((Z - BG1) / Z) - R * T * Math.Log(V * 101325 / (R * 298.15))
            Dim DSres = R * Math.Log((Z - BG1) / Z) + R * Math.Log(V * 101325 / (R * 298.15)) - 1 / (8 ^ 0.5 * bm) * dadT * Math.Log((2 * Z + BG1 * (2 - 8 ^ 0.5)) / (2 * Z + BG1 * (2 + 8 ^ 0.5)))
            Dim DHres = DAres + T * (DSres) + R * T * (Z - 1)

            Return Gid + (DHres - T * DSres)

        End Function

        Public Overrides Function CalcHelmoltzEnergy(ByVal phasetype As String, ByVal T As Double, ByVal P As Double, ByVal Vz As System.Array, ByVal VKij As Object, ByVal VTc As System.Array, ByVal VPc As System.Array, ByVal Vw As System.Array, ByVal Aid As Double, Optional ByVal otherargs As Object = Nothing) As Double

            Dim ai(), bi(), ci() As Double
            Dim n, R As Double
            Dim Tc(), Pc(), Vc(), w(), Zc(), alpha(), m(), a(,), b(,), Z, Tr() As Double
            Dim i, j, dadT

            n = UBound(Vz)

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
            Do
                alpha(i) = (1 + (0.48 + 1.574 * w(i) - 0.176 * w(i) ^ 2) * (1 - (T / Tc(i)) ^ 0.5)) ^ 2
                ai(i) = 0.42748 * alpha(i) * R ^ 2 * Tc(i) ^ 2 / Pc(i)
                bi(i) = 0.08664 * R * Tc(i) / Pc(i)
                ci(i) = 0.48 + 1.574 * w(i) - 0.176 * w(i) ^ 2
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

            coeff(0) = -AG1 * BG1
            coeff(1) = AG1 - BG1 - BG1 ^ 2
            coeff(2) = -1
            coeff(3) = 1

            Dim temp1 = Poly_Roots(coeff)
            Dim tv
            Dim tv2
            Try

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

                If phasetype = "L" Then
                    Z = temp1(0, 0)
                    If temp1(0, 1) <> 0 Then
                        Z = temp1(1, 0)
                        If temp1(1, 1) <> 0 Then
                            Z = temp1(2, 0)
                        End If
                    End If
                ElseIf phasetype = "V" Then
                    Z = temp1(2, 0)
                    If temp1(2, 1) <> 0 Then
                        Z = temp1(1, 0)
                        If temp1(1, 1) <> 0 Then
                            Z = temp1(0, 0)
                        End If
                    End If
                End If
            Catch

                Dim findZ

                If phasetype = "V" Then

                    Z = 1
                    Do
                        findZ = coeff(3) * Z ^ 3 + coeff(2) * Z ^ 2 + coeff(1) * Z + coeff(0)
                        Z -= 0.00001
                        If Z < 0 Then Throw New Exception(App.GetLocalString("PropPack_ZError"))
                    Loop Until Math.Abs(findZ) < 0.0001

                Else

                    Z = 0
                    Do
                        findZ = coeff(3) * Z ^ 3 + coeff(2) * Z ^ 2 + coeff(1) * Z + coeff(0)
                        Z += 0.00001
                        If Z > 1 Then Throw New Exception(App.GetLocalString("PropPack_ZError"))
                    Loop Until Math.Abs(findZ) < 0.0001

                End If

            End Try

            Dim V = (Z * R * T / P) ' m3/mol
            Dim aux1 = -R / 2 * (0.42748 / T) ^ 0.5
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
            uu = 1
            ww = 0

            Dim DAres = am / (bm * (uu ^ 2 - 4 * ww) ^ 0.5) * Math.Log((2 * Z + BG1 * (uu - (uu ^ 2 - 4 * ww) ^ 0.5)) / (2 * Z + BG1 * (uu + (uu ^ 2 - 4 * ww) ^ 0.5))) - R * T * Math.Log((Z - BG1) / Z) - R * T * Math.Log(V * 101325 / (R * 298.15))
            Dim DSres = R * Math.Log((Z - BG1) / Z) + R * Math.Log(V * 101325 / (R * 298.15)) - 1 / (8 ^ 0.5 * bm) * dadT * Math.Log((2 * Z + BG1 * (2 - 8 ^ 0.5)) / (2 * Z + BG1 * (2 + 8 ^ 0.5)))
            Dim DHres = DAres + T * (DSres) + R * T * (Z - 1)

            Return Aid + DAres

        End Function

        Public Overrides Function CalcP(ByVal V As Double, ByVal T As Double, ByVal Vx As System.Array, ByVal VKij As Object, ByVal VTc As System.Array, ByVal VPc As System.Array, ByVal Vw As System.Array, Optional ByVal otherargs As Object = Nothing) As Object

            Dim ai(), bi() As Double
            Dim n, R, P, coeff(3), tmp() As Double
            Dim Tc(), Pc(), W(), alpha(), Vant(0, 4), m(), a(,), b(,), Tr() As Double

            n = UBound(Vx)

            ReDim ai(n), bi(n), tmp(n + 1), a(n, n), b(n, n)
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
                alpha(i) = (1 + (0.48 + 1.574 * W(i) - 0.176 * W(i) ^ 2) * (1 - (T / Tc(i)) ^ 0.5)) ^ 2
                ai(i) = 0.42748 * alpha(i) * R ^ 2 * Tc(i) ^ 2 / Pc(i)
                bi(i) = 0.08664 * R * Tc(i) / Pc(i)
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
            Dim aml = 0.0#
            Do
                j = 0
                Do
                    aml = aml + Vx(i) * Vx(j) * a(i, j)
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

            P = R * T / (V - bml) - aml / (V * (V + bml))

            Return P

        End Function

        Public Overrides Function CalcLnFugTV(ByVal T As Double, ByVal V As Double, ByVal Vx As System.Array, ByVal VKij As Object, ByVal VTc As System.Array, ByVal VPc As System.Array, ByVal Vw As System.Array, Optional ByVal otherargs As Object = Nothing, Optional ByVal forcephase As String = "") As Object

            Dim P As Double = Me.CalcP(V, T, Vx, VKij, VTc, VPc, Vw, otherargs)

            Dim n, R, coeff(3) As Double
            Dim Vant(0, 4) As Double
            Dim criterioOK As Boolean = False
            Dim AG, BG, aml, bml As Double
            Dim t1, t2, t3, t4, t5 As Double

            n = UBound(Vx)

            Dim ai(n), bi(n), tmp(n + 1), a(n, n), b(n, n)
            Dim aml2(n), amv2(n), LN_CF(n), PHI(n) As Double
            Dim Tc(n), Pc(n), W(n), alpha(n), m(n), Tr(n)

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
                alpha(i) = (1 + (0.48 + 1.574 * W(i) - 0.176 * W(i) ^ 2) * (1 - (T / Tc(i)) ^ 0.5)) ^ 2
                ai(i) = 0.42748 * alpha(i) * R ^ 2 * Tc(i) ^ 2 / Pc(i)
                bi(i) = 0.08664 * R * Tc(i) / Pc(i)
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

            Dim Z = P * V / (R * T)

            i = 0
            Do
                t1 = bi(i) * (Z - 1) / bml
                t2 = -Math.Log(Z - BG)
                t3 = AG * (2 * aml2(i) / aml - bi(i) / bml)
                t4 = Math.Log((Z + BG) / Z)
                t5 = BG
                LN_CF(i) = t1 + t2 - (t3 * t4 / t5)
                i = i + 1
            Loop Until i = n + 1

            Return LN_CF

        End Function

        ''' <summary>
        ''' This procedure checks if the compressibility factor is within the allowable region for the specified phase. 
        ''' If not, it generates a pseudo-root cabable of generate properties for the specified phase in order to keep 
        ''' the flash convergence process going forward.
        ''' </summary>
        ''' <param name="Z">The calculated compressibility factor, coming from the EOS</param>
        ''' <param name="a">EOS 'a' mixture parameter</param>
        ''' <param name="b">EOS 'b' mixture parameter</param>
        ''' <param name="P">Pressure in Pa</param>
        ''' <param name="T">Temperature in K</param>
        ''' <param name="phaselabel">'L' for Liquid, 'V' for Vapor.</param>
        ''' <returns>A vector containing the calculated compressibility factor and pressure, if required. 
        ''' If the given compressibility factor is within the allowable range, it is returned together with 
        ''' the specified pressure (no pseudoroot calculation is required).</returns>
        ''' <remarks>This procedure is based on the paper: 
        ''' Mathias, P. M., Boston, J. F. and Watanasiri, S. (1984), 
        ''' Effective utilization of equations of state for thermodynamic properties in process simulation. 
        ''' AIChE J., 30: 182–186. doi: 10.1002/aic.690300203</remarks>
        Public Shared Function CheckRoot(Z As Double, a As Double, b As Double, P As Double, T As Double, phaselabel As String) As Double()

            If a * b = 0.0# Then Return New Double() {Z, P}

            Dim rho, dPdrho, R, Pnew, P_, rho_, Tmc,
                Zcorr, rhomax, rhomin, rhomc, dPdrholim, C0, C1, rho2 As Double
            Dim i As Integer

            R = 8.314

            Tmc = 0.20268 * a / (R * b)
            If T > Tmc Then T = Tmc * 0.9

            rho = P / (Z * R * T)

            dPdrholim = 0.1 * R * T

            'find rhomax

            Dim fx, dfdx As Double
            rhomax = rho
            i = 0
            Do
                fx = -(1 - b ^ 2 * rhomax ^ 2) / (a * b * rhomax ^ 3 + (b * R * T - a) * rhomax ^ 2 + R * T * rhomax)
                dfdx = (R * T * (b * rhomax + 1) ^ 2 - a * rhomax * (b * rhomax - 1) ^ 2 * (b * rhomax + 2)) / (rhomax ^ 2 * (a * rhomax * (b * rhomax - 1) + R * (b * T * rhomax + T)) ^ 2)
                rhomax = rhomax - 0.7 * fx / dfdx
                If rhomax < 0 Then rhomax = -rhomax
                i += 1
            Loop Until Math.Abs(fx) < 0.000001 Or i = 100

            'find rhomin

            rhomin = rho
            i = 0
            Do
                fx = (a * b * rhomin ^ 3 + (b * R * T - a) * rhomin ^ 2 + R * T * rhomin) / (1 - b ^ 2 * rhomin ^ 2)
                dfdx = (R * T * (b * rhomin + 1) ^ 2 - a * rhomin * (b * rhomin - 1) ^ 2 * (b * rhomin + 2)) / (b ^ 2 * rhomin ^ 2 - 1) ^ 2
                rhomin = rhomin - 0.7 * fx / dfdx
                If rhomin < 0 Then rhomin = -rhomin
                i += 1
            Loop Until Math.Abs(fx) < 0.000001 Or i = 100

            'find rhomc

            i = 0
            rhomc = (rhomax - rhomin) / 2
            Do
                fx = -(2 * (a * (b * rhomc - 1) ^ 3 + b * R * T * (b * rhomc + 1) ^ 3)) / (b ^ 2 * rhomc ^ 2 - 1) ^ 3
                dfdx = (2 * (3 * a * rhomc * (b * rhomc - 1) ^ 4 + R * T * (b * rhomc + 1) ^ 4 * (2 * b * rhomc + 1))) / (b ^ 2 * rhomc ^ 2 - 1) ^ 4
                rhomc = rhomc - 0.7 * fx / dfdx
                i += 1
            Loop Until Math.Abs(fx) < 0.000001 Or i = 100

            dPdrho = (R * T * (b * rho + 1) ^ 2 - a * rho * (b * rho - 1) ^ 2 * (b * rho + 2)) / (b ^ 2 * rho ^ 2 - 1) ^ 2

            If phaselabel = "L" Then
                If dPdrho > dPdrholim And rho > rhomc Then
                    Return New Double() {Z, P}
                End If
            Else
                If dPdrho > dPdrholim Then
                    Return New Double() {Z, P}
                End If
            End If

            If phaselabel = "L" Then

                'find rho*, P*

                i = 0
                rho_ = rhomc * 1.1
                Do
                    fx = -0.1 * R * T + (R * T * (b * rho_ + 1) ^ 2 - a * rho_ * (b * rho_ - 1) ^ 2 * (b * rho_ + 2)) / (b ^ 2 * rho_ ^ 2 - 1) ^ 2
                    dfdx = -(2 * (a * (b * rho_ - 1) ^ 3 + b * R * T * (b * rho_ + 1) ^ 3)) / (b ^ 2 * rho_ ^ 2 - 1) ^ 3
                    rho_ = rho_ - fx / dfdx
                    If rho_ < rhomc Then rho_ = rhomc * 1.02
                    i += 1
                Loop Until Math.Abs(fx) < 0.000001 Or i = 100

                P_ = (a * b * rho_ ^ 3 + (b * R * T - a) * rho_ ^ 2 + R * T * rho_) / (1 - b ^ 2 * rho_ ^ 2)

                C1 = 0.1 * R * T * (rho_ - 0.7 * rhomc)

                C0 = P_ - C1 * Math.Log(rho_ - 0.7 * rhomc)

                rho = 0.7 * rhomc + Math.Exp((P - C0) / C1)

                Pnew = (a * b * rho ^ 3 + (b * R * T - a) * rho ^ 2 + R * T * rho) / (1 - b ^ 2 * rho ^ 2)

                Zcorr = Pnew / (rho * R * T)

                If Double.IsNaN(Zcorr) Or Double.IsNaN(Pnew) Or Double.IsInfinity(Zcorr) Or Double.IsInfinity(Pnew) Then
                    Return New Double() {Z, P}
                Else
                    If Zcorr < 0.0# Or Pnew < 0.0# Then
                        Return New Double() {Z, P}
                    Else
                        Return New Double() {Zcorr, Pnew}
                    End If
                End If

            Else

                'find rho*, P*

                i = 0
                rho_ = rhomc * 0.9
                Do
                    fx = -0.1 * R * T + (R * T * (b * rho_ + 1) ^ 2 - a * rho_ * (b * rho_ - 1) ^ 2 * (b * rho_ + 2)) / (b ^ 2 * rho_ ^ 2 - 1) ^ 2
                    dfdx = -(2 * (a * (b * rho_ - 1) ^ 3 + b * R * T * (b * rho_ + 1) ^ 3)) / (b ^ 2 * rho_ ^ 2 - 1) ^ 3
                    rho_ = rho_ - fx / dfdx
                    If rho_ < rhomc Then rho_ = rhomc * 0.98
                    i += 1
                Loop Until Math.Abs(fx) < 0.000001 Or i = 100

                P_ = (a * b * rho_ ^ 3 + (b * R * T - a) * rho_ ^ 2 + R * T * rho_) / (1 - b ^ 2 * rho_ ^ 2)

                rho2 = (rho_ + rhomc) / 2

                C0 = -2 * rho_ * rho2 ^ 2 / (0.1 * R * T * (rho_ ^ 2 - rho2 ^ 2))

                C1 = 2 * rho_ / (0.1 * R * T * (rho_ ^ 2 - rho2 ^ 2))

                rho = (((1 / P) - C0) / C1) ^ 0.5

                Pnew = (a * b * rho ^ 3 + (b * R * T - a) * rho ^ 2 + R * T * rho) / (1 - b ^ 2 * rho ^ 2)

                Zcorr = Pnew / (rho * R * T)

                If Double.IsNaN(Zcorr) Or Double.IsNaN(Pnew) Or Double.IsInfinity(Zcorr) Or Double.IsInfinity(Pnew) Then
                    Return New Double() {Z, P}
                Else
                    If Zcorr < 0.0# Or Pnew < 0.0# Then
                        Return New Double() {Z, P}
                    Else
                        Return New Double() {Zcorr, Pnew}
                    End If
                End If

            End If

            'SRK EOS P=f(rho) derivatives
            'P = (a * b * rho ^ 3 + (b * R * T - a) * rho ^ 2 + R * T * rho) / (1 - b ^ 2 * rho ^ 2)
            'dPdrho = (R * T * (b * rho + 1) ^ 2 - a * rho * (b * rho - 1) ^ 2 * (b * rho + 2)) / (b ^ 2 * rho ^ 2 - 1) ^ 2
            'd2Pdrho2 = -(2 * (a * (b * rho - 1) ^ 3 + b * R * T * (b * rho + 1) ^ 3)) / (b ^ 2 * rho ^ 2 - 1) ^ 3
            'd3Pdrho3 = (2 * (3 * a * rho * (b * rho - 1) ^ 4 + R * T * (b * rho + 1) ^ 4 * (2 * b * rho + 1))) / (b ^ 2 * rho ^ 2 - 1) ^ 4


        End Function


    End Class

End Namespace
