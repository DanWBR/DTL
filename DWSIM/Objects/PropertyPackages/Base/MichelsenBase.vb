'    Advanced Flash Calculation Routines (Michelsen's Stability Test & Phase Split Algorithm)
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

Imports System.Math
Imports DTL.DTL.MathEx

Namespace DTL.SimulationObjects.PropertyPackages

    <System.Serializable()> Public MustInherit Class ThermoPlug

        Public MustOverride Function CalcP(ByVal V As Double, ByVal T As Double, ByVal Vx As Array, ByVal VKij As Object, ByVal VTc As Array, ByVal VPc As Array, ByVal Vw As Array, Optional ByVal otherargs As Object = Nothing)

        Public MustOverride Function CalcLnFug(ByVal T As Double, ByVal P As Double, ByVal Vx As Array, ByVal VKij As Object, ByVal VTc As Array, ByVal VPc As Array, ByVal Vw As Array, Optional ByVal otherargs As Object = Nothing, Optional ByVal forcephase As String = "")

        Public MustOverride Function CalcLnFugTV(ByVal T As Double, ByVal V As Double, ByVal Vx As Array, ByVal VKij As Object, ByVal VTc As Array, ByVal VPc As Array, ByVal Vw As Array, Optional ByVal otherargs As Object = Nothing, Optional ByVal forcephase As String = "")

        Public MustOverride Function PhaseType(ByVal T As Double, ByVal P As Double, ByVal Vx As Array, ByVal VKij As Object, ByVal VTc As Array, ByVal VPc As Array, ByVal Vw As Array, Optional ByVal otherargs As Object = Nothing)

        Public MustOverride Function CalcEnthalpy(ByVal phasetype As String, ByVal T As Double, ByVal P As Double, ByVal Vz As Array, ByVal VKij As Object, ByVal VTc As Array, ByVal VPc As Array, ByVal Vw As Array, ByVal Hid As Double, Optional ByVal otherargs As Object = Nothing) As Double

        Public MustOverride Function CalcEntropy(ByVal phasetype As String, ByVal T As Double, ByVal P As Double, ByVal Vz As Array, ByVal VKij As Object, ByVal VTc As Array, ByVal VPc As Array, ByVal Vw As Array, ByVal Sid As Double, Optional ByVal otherargs As Object = Nothing) As Double

        Public MustOverride Function CalcGibbsEnergy(ByVal phasetype As String, ByVal T As Double, ByVal P As Double, ByVal Vz As Array, ByVal VKij As Object, ByVal VTc As Array, ByVal VPc As Array, ByVal Vw As Array, ByVal Gid As Double, Optional ByVal otherargs As Object = Nothing) As Double

        Public MustOverride Function CalcHelmoltzEnergy(ByVal phasetype As String, ByVal T As Double, ByVal P As Double, ByVal Vz As Array, ByVal VKij As Object, ByVal VTc As Array, ByVal VPc As Array, ByVal Vw As Array, ByVal Aid As Double, Optional ByVal otherargs As Object = Nothing) As Double

    End Class

    <System.Serializable()> Public Class MichelsenFlash

        Public thermobase As ThermoPlug
        Public ppbase As PropertyPackage

#Region "        Enums"

        Public Enum ErrorCode

            MaximumIterationsReached = 0
            IterationDiverged = 1

        End Enum

        Public Enum FallbackType
            VL = 0
            VLL = 1
        End Enum

#End Region

#Region "        Properties"

        Private m_flashtype As String = "LL"
        Private m_flashfallback As FallbackType = FallbackType.VLL
        Private m_tolerance As Double = 0.00001
        Private m_miti As Integer = 200
        Private m_mite As Integer = 200
        Private m_keycomp As Integer = 0

        Public ReadOnly Property TBase() As ThermoPlug
            Get
                Return thermobase
            End Get
        End Property

        Public Property FlashType() As String
            Get
                Return m_flashtype
            End Get
            Set(ByVal value As String)
                m_flashtype = value
            End Set
        End Property

        Public Property FlashFallbackType() As FallbackType
            Get
                Return m_flashfallback
            End Get
            Set(ByVal value As FallbackType)
                m_flashfallback = value
            End Set
        End Property

        Public Property Tolerance() As Double
            Get
                Return m_tolerance
            End Get
            Set(ByVal value As Double)
                m_tolerance = value
            End Set
        End Property

        Public Property MaxIntIterations() As Integer
            Get
                Return m_miti
            End Get
            Set(ByVal value As Integer)
                m_miti = value
            End Set
        End Property

        Public Property MaxExtIterations() As Integer
            Get
                Return m_mite
            End Get
            Set(ByVal value As Integer)
                m_mite = value
            End Set
        End Property

        Public Property KeyComponentIndex() As Integer
            Get
                Return m_keycomp
            End Get
            Set(ByVal value As Integer)
                m_keycomp = value
            End Set
        End Property

#End Region

#Region "        Base Flash and Stability Codes"

        Private Function StabTest(ByVal T As Double, ByVal P As Double, ByVal Vz As Array, ByVal VKij As Object, ByVal VTc As Array, ByVal VPc As Array, ByVal Vw As Array, Optional ByVal VzArray(,) As Double = Nothing, Optional ByVal otherargs As Object = Nothing)

            Dim i, j, c, n, o, l As Integer
            n = UBound(Vz)

            Dim Y, K As Double(,)

            If m_flashtype = "VL" Then
                ReDim Y(1, n), K(0, n)
            Else
                If Not VzArray Is Nothing Then ReDim Y(n + 2, n), K(0, n) Else ReDim Y(n + 1, n), K(0, n)
            End If

            Dim m As Integer = UBound(Y, 1)

            Dim h(n), lnfi_z(n), Y_ant(m, n) As Double

            lnfi_z = thermobase.CalcLnFug(T, P, Vz, VKij, VTc, VPc, Vw)

            i = 0
            Do
                h(i) = Log(Vz(i)) + lnfi_z(i)
                i = i + 1
            Loop Until i = n + 1

            If m_flashtype = "VL" Then
                i = 0
                Do
                    K(0, i) = VPc(i) / P * Exp(5.42 * (1 - VTc(i) / T))
                    i = i + 1
                Loop Until i = n + 1
                i = 0
                Do
                    Y(0, i) = K(0, i) * Vz(i)
                    Y(1, i) = Vz(i) / K(0, i)
                    i = i + 1
                Loop Until i = n + 1
            Else
                i = 0
                Do
                    j = 0
                    Do
                        If i = j Then Y(i, j) = 1 Else Y(i, j) = 0.0000000001
                        j = j + 1
                    Loop Until j = n + 1
                    i = i + 1
                Loop Until i = n + 1
                If Not VzArray Is Nothing Then
                    Dim sum0(n) As Double
                    i = 0
                    Do
                        sum0(i) = 0
                        j = 0
                        Do
                            sum0(i) += VzArray(j, i)
                            j = j + 1
                        Loop Until j = UBound(VzArray, 1) + 1
                        i = i + 1
                    Loop Until i = n + 1
                    i = 0
                    Do
                        Y(n + 1, i) = sum0(i) / UBound(VzArray, 1)
                        Y(n + 2, i) = Exp(h(i))
                        i = i + 1
                    Loop Until i = n + 1
                Else
                    i = 0
                    Do
                        Y(n + 1, i) = Exp(h(i))
                        i = i + 1
                    Loop Until i = n + 1
                End If
            End If

            Dim lnfi(m, n), beta(m), r(m), r_ant(m) As Double
            Dim currcomp(n) As Double
            Dim dgdY(m, n), g_(m), tmpfug(n), dY(m, n), sum3 As Double
            Dim excidx As New ArrayList
            Dim finish As Boolean = True

            c = 0
            Do

                'start stability test for each one of the initial estimate vectors
                i = 0
                Do
                    If Not excidx.Contains(i) Then
                        j = 0
                        sum3 = 0
                        Do
                            If Y(i, j) > 0 Then sum3 += Y(i, j)
                            j = j + 1
                        Loop Until j = n + 1
                        j = 0
                        Do
                            If Y(i, j) > 0 Then currcomp(j) = Y(i, j) / sum3 Else currcomp(j) = 0
                            j = j + 1
                        Loop Until j = n + 1
                        tmpfug = thermobase.CalcLnFug(T, P, currcomp, VKij, VTc, VPc, Vw)
                        j = 0
                        Do
                            lnfi(i, j) = tmpfug(j)
                            j = j + 1
                        Loop Until j = n + 1
                        j = 0
                        Do
                            dgdY(i, j) = Log(Y(i, j)) + lnfi(i, j) - h(j)
                            j = j + 1
                        Loop Until j = n + 1
                        j = 0
                        beta(i) = 0
                        Do
                            beta(i) += (Y(i, j) - Vz(j)) * dgdY(i, j)
                            j = j + 1
                        Loop Until j = n + 1
                        g_(i) = 1
                        j = 0
                        Do
                            g_(i) += Y(i, j) * (Log(Y(i, j)) + lnfi(i, j) - h(j) - 1)
                            j = j + 1
                        Loop Until j = n + 1
                        If i > 0 Then r_ant(i) = r(i) Else r_ant(i) = 0
                        r(i) = 2 * g_(i) / beta(i)
                    End If
                    i = i + 1
                Loop Until i = m + 1

                i = 0
                Do
                    If (Abs(g_(i)) < 0.0000000001 And r(i) > 0.9 And r(i) < 1.1) Then
                        If Not excidx.Contains(i) Then excidx.Add(i)
                        'ElseIf c > 4 And r(i) > r_ant(i) Then
                        '    If Not excidx.Contains(i) Then excidx.Add(i)
                    End If
                    i = i + 1
                Loop Until i = m + 1

                i = 0
                Do
                    If Not excidx.Contains(i) Then
                        j = 0
                        Do
                            Y_ant(i, j) = Y(i, j)
                            Y(i, j) = Exp(h(j) - lnfi(i, j))
                            dY(i, j) = Y(i, j) - Y_ant(i, j)
                            If Y(i, j) < 0 Then Y(i, j) = 0
                            j = j + 1
                        Loop Until j = n + 1
                    End If
                    i = i + 1
                Loop Until i = m + 1

                'check convergence

                finish = True
                i = 0
                Do
                    If Not excidx.Contains(i) Then
                        j = 0
                        Do
                            If Abs(dY(i, j)) > m_tolerance Then finish = False
                            j = j + 1
                        Loop Until j = n + 1
                    End If
                    i = i + 1
                Loop Until i = m + 1

                c = c + 1

                If c > m_mite Then Throw New Exception(ErrorCode.MaximumIterationsReached)

            Loop Until finish = True

            ' search for trivial solutions

            Dim sum As Double
            i = 0
            Do
                If Not excidx.Contains(i) Then
                    j = 0
                    sum = 0
                    Do
                        sum += Abs(Y(i, j) - Vz(j))
                        j = j + 1
                    Loop Until j = n + 1
                    If sum < 0.001 Then
                        If Not excidx.Contains(i) Then excidx.Add(i)
                    End If
                End If
                i = i + 1
            Loop Until i = m + 1

            ' search for trivial solutions

            Dim sum5 As Double
            i = 0
            Do
                If Not excidx.Contains(i) Then
                    j = 0
                    sum5 = 0
                    Do
                        sum5 += Y(i, j)
                        j = j + 1
                    Loop Until j = n + 1
                    If sum5 < 1 Then
                        'phase is stable
                        If Not excidx.Contains(i) Then excidx.Add(i)
                    End If
                End If
                i = i + 1
            Loop Until i = m + 1

            ' join similar solutions

            Dim similar As Boolean

            i = 0
            Do
                If Not excidx.Contains(i) Then
                    o = 0
                    Do
                        If Not excidx.Contains(o) And i <> o Then
                            similar = True
                            j = 0
                            Do
                                If Abs(Y(i, j) - Y(o, j)) > 0.00001 Then
                                    similar = False
                                End If
                                j = j + 1
                            Loop Until j = n + 1
                            If similar Then
                                excidx.Add(o)
                                Exit Do
                            End If
                        End If
                        o = o + 1
                    Loop Until o = m + 1
                End If
                i = i + 1
            Loop Until i = m + 1

            l = excidx.Count
            Dim sum2 As Double
            Dim isStable As Boolean

            If m + 1 - l > 0 Then

                'the phase is unstable

                isStable = False

                'normalize initial estimates

                Dim inest(m - l, n) As Double
                i = 0
                l = 0
                Do
                    If Not excidx.Contains(i) Then
                        j = 0
                        sum2 = 0
                        Do
                            sum2 += Y(i, j)
                            j = j + 1
                        Loop Until j = n + 1
                        j = 0
                        Do
                            inest(l, j) = Y(i, j) / sum2
                            j = j + 1
                        Loop Until j = n + 1
                        l = l + 1
                    End If
                    i = i + 1
                Loop Until i = m + 1
                Return New Object() {isStable, inest}
            Else

                'the phase is stable

                isStable = True
                Return New Object() {isStable, Nothing}
            End If

        End Function

        Private Function Flash_TP_RR(ByVal T As Double, ByVal P As Double, ByVal Vz As Array, ByVal VKij As Object, ByVal VTc As Array, ByVal VPc As Array, ByVal Vw As Array, Optional ByVal otherargs As Object = Nothing)

            Dim n, i, ci As Integer

            n = UBound(Vz)

            Dim K(n), Kant(n), lnfugl(n), lnfugv(n), x(n), y(n), xsum, ysum, v, vant, l As Double

            'generate ideal K-values
            i = 0
            Do
                K(i) = VPc(i) / P * Exp(5.42 * (1 - VTc(i) / T))
                i = i + 1
            Loop Until i = n + 1
            v = CalcV(Vz, K)
            If v = -100 Then
                v = 0.5
            End If
            If v > 1 Then v = 1
            If v < 0 Then v = 0
            l = 1 - v
            i = 0
            Do
                If Vz(i) <> 0 Then
                    y(i) = Vz(i) * K(i) / ((K(i) - 1) * v + 1)
                    x(i) = y(i) / K(i)
                Else
                    y(i) = 0
                    x(i) = 0
                End If
                i = i + 1
            Loop Until i = n + 1
            i = 0
            xsum = 0
            ysum = 0
            Do
                xsum += x(i)
                ysum += y(i)
                i = i + 1
            Loop Until i = n + 1
            i = 0
            Do
                x(i) = x(i) / xsum
                y(i) = y(i) / ysum
                i = i + 1
            Loop Until i = n + 1

            Dim pf(1) As Double, pty(1) As String, pic(1) As Double
            Dim tmpres As Object

            ci = 0
            Do

                'vapor phase fugacity coefficients
                lnfugv = thermobase.CalcLnFug(T, P, y, VKij, VTc, VPc, Vw, otherargs)
                tmpres = thermobase.PhaseType(T, P, y, VKij, VTc, VPc, Vw, otherargs)
                pf(0) = v
                pty(0) = tmpres(0)
                pic(0) = tmpres(1)

                'liquid phase fugacity coefficients
                lnfugl = thermobase.CalcLnFug(T, P, x, VKij, VTc, VPc, Vw, otherargs)
                tmpres = thermobase.PhaseType(T, P, x, VKij, VTc, VPc, Vw, otherargs)
                pf(1) = l
                pty(1) = tmpres(0)
                pic(1) = tmpres(1)

                'update K-values
                i = 0
                Do
                    Kant(i) = K(i)
                    K(i) = Exp(lnfugl(i) - lnfugv(i))
                    i = i + 1
                Loop Until i = n + 1

                'check convergence
                Dim sum1 As Double = 0
                i = 0
                Do
                    sum1 += Abs(K(i) - Kant(i))
                    i = i + 1
                Loop Until i = n + 1

                If ci >= m_mite Then Throw New Exception(ErrorCode.MaximumIterationsReached)
                If Double.IsNaN(sum1) Then Throw New Exception(ErrorCode.IterationDiverged)

                If sum1 <= m_tolerance Then Exit Do

                'update vapor fraction by newton's method
                vant = v
                Dim F = 0.0#
                Dim dF = 0.0#
                i = 0
                Do
                    If Vz(i) > 0 Then
                        F = F + Vz(i) * (K(i) - 1) / (1 + v * (K(i) - 1))
                        dF = dF - Vz(i) * (K(i) - 1) ^ 2 / (1 + v * (K(i) - 1)) ^ 2
                    End If
                    i = i + 1
                Loop Until i = n + 1
                If Math.Abs(F / dF) > 1 Then
                    v = -F / dF * 0.5 * v + v
                Else
                    v = -F / dF + v
                End If
                l = 1 - v
                If v > 1 Then
                    v = 1
                    l = 0
                    i = 0
                    Do
                        y(i) = Vz(i)
                        x(i) = 0
                        i = i + 1
                    Loop Until i = n + 1
                    pf(0) = v
                    pf(1) = l
                    Exit Do
                ElseIf v < 0 Then
                    v = 0
                    l = 1
                    i = 0
                    Do
                        x(i) = Vz(i)
                        y(i) = 0
                        i = i + 1
                    Loop Until i = n + 1
                    pf(0) = v
                    pf(1) = l
                    Exit Do
                Else
                    i = 0
                    Do
                        If Vz(i) <> 0 Then
                            y(i) = Vz(i) * K(i) / ((K(i) - 1) * v + 1)
                            x(i) = y(i) / K(i)
                        Else
                            y(i) = 0
                            x(i) = 0
                        End If
                        i = i + 1
                    Loop Until i = n + 1
                    i = 0
                    xsum = 0
                    ysum = 0
                    Do
                        xsum += x(i)
                        ysum += y(i)
                        i = i + 1
                    Loop Until i = n + 1
                    i = 0
                    Do
                        x(i) = x(i) / xsum
                        y(i) = y(i) / ysum
                        i = i + 1
                    Loop Until i = n + 1
                End If

                ci += 1

            Loop

            ' return phase amounts, compositions, fugacity coefficients,
            ' compr. factors and isothermal compressibilities

            Dim result(1, 4) As Object

            result(0, 0) = pf(0)
            result(0, 1) = y
            result(0, 2) = lnfugv
            result(0, 3) = pic(0)

            result(1, 0) = pf(1)
            result(1, 1) = x
            result(1, 2) = lnfugl
            result(1, 3) = pic(1)

            Return result

        End Function

        Private Function Flash_TP_RR_3P(ByVal T As Double, ByVal P As Double, ByVal Vz As Array, ByVal VKij As Object, ByVal VTc As Array, ByVal VPc As Array, ByVal Vw As Array, Optional ByVal otherargs As Object = Nothing)

            Dim n, i, ci As Integer

            n = UBound(Vz)

            Dim K(n), K1(n), Kant(n), K1ant(n), b1(n), b2(n), lnfugl(n), lnfugl1(n), lnfugv(n), x(n), x1(n), y(n), xsum, xsum1, ysum, v, l, lant, l1, l1ant As Double

            'generate ideal K-values
            i = 0
            Do
                K(i) = VPc(i) / P * Exp(5.42 * (1 - VTc(i) / T))
                i = i + 1
            Loop Until i = n + 1
            v = CalcV(Vz, K)
            If v = -100 Then
                v = 0.5
            End If
            If v > 1 Then v = 0.999999
            If v < 0 Then v = 0.000001
            Dim Vp(n) As Double
            i = 0
            Do
                Vp(i) = Me.ppbase.AUX_PVAPi(i, T)
                i = i + 1
            Loop Until i = n + 1
            If Common.Max(Vp) < P Then v = 0.000001
            l1 = Vz(m_keycomp)
            If 1 - v - l1 < 0 Then v = 1 - l1 - 0.01
            l = 1 - v - l1
            If l < 0 Then l = 0
            i = 0
            Do
                If i = m_keycomp Then
                    K(i) = 10000
                    K1(i) = 0.001
                Else
                    K(i) = VPc(i) / P * Exp(5.42 * (1 - VTc(i) / T))
                    K1(i) = 10000
                End If
                i = i + 1
            Loop Until i = n + 1
            i = 0
            Do
                b1(i) = 1 - K(i) ^ -1
                b2(i) = 1 - K1(i) ^ -1
                i = i + 1
            Loop Until i = n + 1
            i = 0
            Do
                If Vz(i) <> 0 Then
                    y(i) = Vz(i) / (1 - b1(i) * l - b2(i) * l1)
                    x(i) = y(i) * (1 - b1(i))
                    x1(i) = y(i) * (1 - b2(i))
                Else
                    y(i) = 0
                    x1(i) = 0
                    x(i) = 0
                End If
                i += 1
            Loop Until i = n + 1
            i = 0
            xsum = 0
            xsum1 = 0
            ysum = 0
            Do
                xsum += x(i)
                xsum1 += x1(i)
                ysum += y(i)
                i = i + 1
            Loop Until i = n + 1
            i = 0
            Do
                x(i) = x(i) / xsum
                x1(i) = x1(i) / xsum1
                y(i) = y(i) / ysum
                i = i + 1
            Loop Until i = n + 1

            Dim pf(2) As Double, pty(2) As String, pic(2) As Double
            Dim tmpres As Object

            ci = 0
            Do

                'vapor phase fugacity coefficients
                lnfugv = thermobase.CalcLnFug(T, P, y, VKij, VTc, VPc, Vw, otherargs)
                tmpres = thermobase.PhaseType(T, P, y, VKij, VTc, VPc, Vw, otherargs)
                pf(0) = v
                pty(0) = tmpres(0)
                pic(0) = tmpres(1)

                'liquid phase (1) fugacity coefficients
                lnfugl = thermobase.CalcLnFug(T, P, x, VKij, VTc, VPc, Vw, otherargs)
                tmpres = thermobase.PhaseType(T, P, x, VKij, VTc, VPc, Vw, otherargs)
                pf(1) = l
                pty(1) = tmpres(0)
                pic(1) = tmpres(1)

                'liquid phase (2) fugacity coefficients
                lnfugl1 = thermobase.CalcLnFug(T, P, x1, VKij, VTc, VPc, Vw, otherargs)
                tmpres = thermobase.PhaseType(T, P, x1, VKij, VTc, VPc, Vw, otherargs)
                pf(2) = l1
                pty(2) = tmpres(0)
                pic(2) = tmpres(1)

                'update K-values
                i = 0
                Do
                    Kant(i) = K(i)
                    K1ant(i) = K1(i)
                    K(i) = Exp(lnfugl(i) - lnfugv(i))
                    K1(i) = Exp(lnfugl1(i) - lnfugv(i))
                    i = i + 1
                Loop Until i = n + 1

                i = 0
                Do
                    b1(i) = 1 - K(i) ^ -1
                    b2(i) = 1 - K1(i) ^ -1
                    i = i + 1
                Loop Until i = n + 1

                'update vapor/liquid fractions by newton's method
                Dim f As Mapack.Matrix = New Mapack.Matrix(2, 1)
                Dim dl As Mapack.Matrix = New Mapack.Matrix(2, 1)
                Dim df As Mapack.Matrix = New Mapack.Matrix(2, 2)

                f(0, 0) = 0
                f(1, 0) = 0
                df(0, 0) = 0
                df(1, 0) = 0
                df(0, 1) = 0
                df(1, 1) = 0
                i = 0
                Do
                    f(0, 0) += b1(i) * Vz(i) / (1 - b1(i) * l - b2(i) * l1)
                    f(1, 0) += b2(i) * Vz(i) / (1 - b1(i) * l - b2(i) * l1)
                    df(0, 0) += b1(i) * Vz(i) * (-b1(i)) / (1 - b1(i) * l - b2(i) * l1) ^ 2
                    df(0, 1) += b1(i) * Vz(i) * (-b2(i)) / (1 - b1(i) * l - b2(i) * l1) ^ 2
                    df(1, 0) += b2(i) * Vz(i) * (-b1(i)) / (1 - b1(i) * l - b2(i) * l1) ^ 2
                    df(1, 1) += b2(i) * Vz(i) * (-b2(i)) / (1 - b1(i) * l - b2(i) * l1) ^ 2
                    i = i + 1
                Loop Until i = n + 1

                Try
                    'WriteDebugInfo(df.ToString)
                    dl = df.Solve(f)
                    'WriteDebugInfo(dl.ToString())
                    'Console.ReadKey()

                Catch ex As Exception

                    dl(0, 0) = 0
                    dl(1, 0) = 0

                End Try

                l1ant = l1
                lant = l

                l += dl(0, 0)
                l1 += dl(1, 0)

                v = 1 - l - l1

                If v < 0 Then
                    v = 0
                    l = Abs(l) / (Abs(l) + Abs(l1))
                    l1 = 1 - l
                ElseIf v > 1 Then
                    v = 1
                    l = 0
                    l1 = 0
                End If

                'normalize
                Dim sumfrac As Double = 0
                sumfrac = Abs(v) + Abs(l) + Abs(l1)

                v = Abs(v) / sumfrac
                l = Abs(l) / sumfrac
                l1 = Abs(l1) / sumfrac

                'check convergence
                Dim sum1 As Double = 0
                i = 0
                Do
                    sum1 += Abs(l - lant) + Abs(l1 - l1ant)
                    i = i + 1
                Loop Until i = n + 1

                If sum1 <= m_tolerance Then Exit Do
                If ci >= m_mite Then Throw New Exception(ErrorCode.MaximumIterationsReached)
                If Double.IsNaN(sum1) Then Throw New Exception(ErrorCode.IterationDiverged)

                i = 0
                Do
                    If Vz(i) <> 0 Then
                        y(i) = Vz(i) / (1 - b1(i) * l - b2(i) * l1)
                        x(i) = y(i) * (1 - b1(i))
                        x1(i) = y(i) * (1 - b2(i))
                    Else
                        y(i) = 0
                        x1(i) = 0
                        x(i) = 0
                    End If
                    i += 1
                Loop Until i = n + 1
                i = 0
                xsum = 0
                xsum1 = 0
                ysum = 0
                Do
                    xsum += x(i)
                    xsum1 += x1(i)
                    ysum += y(i)
                    i = i + 1
                Loop Until i = n + 1
                i = 0
                Do
                    x(i) = x(i) / xsum
                    x1(i) = x1(i) / xsum1
                    y(i) = y(i) / ysum
                    i = i + 1
                Loop Until i = n + 1

                ci += 1

            Loop

            ' return phase amounts, compositions, fugacity coefficients,
            ' compr. factors and isothermal compressibilities

            Dim result(2, 3) As Object

            result(0, 0) = pf(0)
            result(0, 1) = y
            result(0, 2) = lnfugv
            result(0, 3) = pic(0)

            result(1, 0) = pf(1)
            result(1, 1) = x
            result(1, 2) = lnfugl
            result(1, 3) = pic(1)

            result(2, 0) = pf(2)
            result(2, 1) = x1
            result(2, 2) = lnfugl1
            result(2, 3) = pic(2)

            Return result

        End Function

        Private Function Flash_TP_Ideal(ByVal T As Double, ByVal P As Double, ByVal Vz As Array, ByVal VKij As Object, ByVal VTc As Array, ByVal VPc As Array, ByVal Vw As Array, Optional ByVal otherargs As Object = Nothing)

            Dim n, i As Integer

            n = UBound(Vz)

            Dim K(n), Kant(n), lnfugl(n), lnfugv(n), x(n), y(n), xsum, ysum, v, vant, l As Double

            'generate ideal K-values
            i = 0
            Do
                K(i) = VPc(i) / P * Exp(5.42 * (1 - VTc(i) / T))
                i = i + 1
            Loop Until i = n + 1
            v = CalcV(Vz, K)
            If v = -100 Then
                v = 0.5
            End If
            If v > 1 Then v = 1
            If v < 0 Then v = 0
            l = 1 - v
            i = 0
            Do
                If Vz(i) <> 0 Then
                    y(i) = Vz(i) * K(i) / ((K(i) - 1) * v + 1)
                    x(i) = y(i) / K(i)
                Else
                    y(i) = 0
                    x(i) = 0
                End If
                i = i + 1
            Loop Until i = n + 1
            i = 0
            xsum = 0
            ysum = 0
            Do
                xsum += x(i)
                ysum += y(i)
                i = i + 1
            Loop Until i = n + 1
            i = 0
            Do
                x(i) = x(i) / xsum
                y(i) = y(i) / ysum
                i = i + 1
            Loop Until i = n + 1

            Dim pf(1) As Double, pty(1) As String, pic(1) As Double
            Dim tmpres As Object

            'vapor phase properties
            tmpres = thermobase.PhaseType(T, P, y, VKij, VTc, VPc, Vw, otherargs)
            pf(0) = v
            pty(0) = tmpres(0)
            pic(0) = tmpres(1)

            'liquid phase properties
            tmpres = thermobase.PhaseType(T, P, x, VKij, VTc, VPc, Vw, otherargs)
            pf(1) = l
            pty(1) = tmpres(0)
            pic(1) = tmpres(1)

            'update vapor fraction by newton's method
            vant = v
            Dim F = 0.0#
            Dim dF = 0.0#
            i = 0
            Do
                If Vz(i) > 0 Then
                    F = F + Vz(i) * (K(i) - 1) / (1 + v * (K(i) - 1))
                    dF = dF - Vz(i) * (K(i) - 1) ^ 2 / (1 + v * (K(i) - 1)) ^ 2
                End If
                i = i + 1
            Loop Until i = n + 1
            If Math.Abs(F / dF) > 1 Then
                v = -F / dF * 0.5 * v + v
            Else
                v = -F / dF + v
            End If
            l = 1 - v
            If v > 1 Then
                v = 1
                l = 0
                i = 0
                Do
                    y(i) = Vz(i)
                    x(i) = 0
                    i = i + 1
                Loop Until i = n + 1
                pf(0) = v
                pf(1) = l
            ElseIf v < 0 Then
                v = 0
                l = 1
                i = 0
                Do
                    x(i) = Vz(i)
                    y(i) = 0
                    i = i + 1
                Loop Until i = n + 1
                pf(0) = v
                pf(1) = l
            Else
                i = 0
                Do
                    If Vz(i) <> 0 Then
                        y(i) = Vz(i) * K(i) / ((K(i) - 1) * v + 1)
                        x(i) = y(i) / K(i)
                    Else
                        y(i) = 0
                        x(i) = 0
                    End If
                    i = i + 1
                Loop Until i = n + 1
                i = 0
                xsum = 0
                ysum = 0
                Do
                    xsum += x(i)
                    ysum += y(i)
                    i = i + 1
                Loop Until i = n + 1
                i = 0
                Do
                    x(i) = x(i) / xsum
                    y(i) = y(i) / ysum
                    i = i + 1
                Loop Until i = n + 1
            End If

            ' return phase amounts, compositions, fugacity coefficients,
            ' compr. factors and isothermal compressibilities

            Dim result(1, 4) As Object

            result(0, 0) = pf(0)
            result(0, 1) = y
            result(0, 2) = lnfugv
            result(0, 3) = pic(0)

            result(1, 0) = pf(1)
            result(1, 1) = x
            result(1, 2) = lnfugl
            result(1, 3) = pic(1)

            Return result

        End Function

        Private Function CalcV(ByVal Vz As Object, ByVal KI As Object) As Double

            Dim n = UBound(Vz)

            Dim i As Integer

            Dim Vinf, Vsup As Double

            Dim fV, fV_inf, nsub, delta_V As Double

            Vinf = 0
            Vsup = 1

            nsub = 10

            delta_V = (Vsup - Vinf) / nsub

            i = 0
            Do
                i = i + 1
                fV = OFunc_V(Vinf, Vz, KI)
                Vinf = Vinf + delta_V
                fV_inf = OFunc_V(Vinf, Vz, KI)
            Loop Until fV * fV_inf < 0 Or i = 11
            If i = 11 Then GoTo Final2
            Vsup = Vinf
            Vinf = Vinf - delta_V

            'método de Brent para encontrar Vc

            Dim aaa, bbb, ccc, ddd, eee, min11, min22, faa, fbb, fcc, ppp, qqq, rrr, sss, tol11, xmm As Double
            Dim ITMAX2 As Integer = 100
            Dim iter2 As Integer

            aaa = Vinf
            bbb = Vsup
            ccc = Vsup

            faa = OFunc_V(aaa, Vz, KI)
            fbb = OFunc_V(bbb, Vz, KI)
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
                fbb = OFunc_V(bbb, Vz, KI)
                iter2 += 1
            Loop Until iter2 = ITMAX2

Final2:     bbb = -100

Final3:

            Return bbb

        End Function

        Private Function OFunc_V(ByVal V As Double, ByVal Vz As Object, ByVal KI As Object) As Double

            Dim i As Integer
            Dim n = UBound(Vz)
            Dim result As Double

            i = 0
            Do
                result += Vz(i) * (1 - KI(i)) / (1 - V + V * KI(i))
                i = i + 1
            Loop Until i = n + 1

            Return result

        End Function

#End Region

#Region "        Bubble/Dew Point Functions"

        Function _dKdT(ByVal T As Double, ByVal P As Double, ByVal Vx As Object, ByVal Vy As Object, ByVal VKij As Object, ByVal VTc As Object, ByVal VPc As Object, ByVal Vw As Object, Optional ByVal otherargs As Object = Nothing) As Object

            Dim n = UBound(Vx)

            Dim LN_CFL(n), LN_CFV(n) As Double
            Dim i As Integer
            Dim h = 0.0001

            Dim KI(n), KI2(n), KI_ant(n), dKdT(n)
            Dim chk As Boolean = False

Start:

            LN_CFL = thermobase.CalcLnFug(T, P, Vx, VKij, VTc, VPc, Vw, otherargs)
            LN_CFV = thermobase.CalcLnFug(T, P, Vy, VKij, VTc, VPc, Vw, otherargs)

            i = 0
            Do
                KI(i) = (Math.Exp(LN_CFL(i)) / (Math.Exp(LN_CFV(i))))
                If chk = False Then KI2(i) = (Math.Exp(LN_CFL(i)) / (Math.Exp(LN_CFV(i))))
                i = i + 1
            Loop Until i = n + 1

            If chk = True Then GoTo Final

            T = T + h

            chk = True

            GoTo Start

Final:

            i = 0
            Do
                dKdT(i) = (KI(i) - KI2(i)) / h
                i = i + 1
            Loop Until i = n + 1

            Return dKdT

        End Function

        Function _dKdP(ByVal T As Double, ByVal P As Double, ByVal Vx As Object, ByVal Vy As Object, ByVal VKij As Object, ByVal VTc As Object, ByVal VPc As Object, ByVal Vw As Object, Optional ByVal otherargs As Object = Nothing) As Object

            Dim n = UBound(Vx)

            Dim LN_CFL(n), LN_CFV(n) As Double
            Dim i As Integer
            Dim h = 0.00001

            Dim KI(n), KI2(n), KI_ant(n), dKdP(n)

            Dim chk As Boolean = False

Start:

            i = 0
            Do
                KI_ant(i) = Vy(i) / Vx(i)
                i = i + 1
            Loop Until i = n + 1

            LN_CFL = thermobase.CalcLnFug(T, P, Vx, VKij, VTc, VPc, Vw, otherargs)
            LN_CFV = thermobase.CalcLnFug(T, P, Vy, VKij, VTc, VPc, Vw, otherargs)

            i = 0
            Do
                KI(i) = (Math.Exp(LN_CFL(i)) / (Math.Exp(LN_CFV(i))))
                If chk = False Then KI2(i) = (Math.Exp(LN_CFL(i)) / (Math.Exp(LN_CFV(i))))
                i = i + 1
            Loop Until i = n + 1

            If chk = True Then GoTo Final

            P = P + h

            chk = True

            GoTo Start

Final:

            i = 0
            Do
                dKdP(i) = (KI(i) - KI2(i)) / h
                i = i + 1
            Loop Until i = n + 1

            Return dKdP

        End Function

        Function DewP(ByVal T As Double, ByVal Vy As Object, ByVal VKij As Object, ByVal VTc As Object, ByVal VPc As Object, ByVal Vw As Object, ByVal KI As Object, Optional ByVal P As Double = 0, Optional ByVal otherargs As Object = Nothing) As Object

            Dim dFdP As Double
            Dim cnt As Integer = 0

            Dim n = UBound(Vy)

            Dim chk As Boolean = False
            Dim marcador2 As Integer

            Dim F As Double
            Dim dKdP(n), dKdPi(n) As Object
            Dim Vx(n), LN_CFL(n), LN_CFV(n) As Double
            Dim Vp(n), R, coeff(3) As Double
            Dim i As Integer
            Dim soma_x As Double
            Dim marcador
            Dim stmp4_ant, stmp4, Pant As Double
            stmp4_ant = 0
            stmp4 = 0
            Pant = 0

            R = 8.314

            i = 0
            Do
                Vp(i) = ppbase.AUX_PVAPi(i, T)
                i = i + 1
            Loop Until i = n + 1

            If P = 0 Then
                'generate first estimate for P
                i = 0
                Do
                    P = P + (Vy(i) / Vp(i))
                    i = i + 1
                Loop Until i = n + 1
                P = 1 / P
            End If

            i = 0
            Do
                If KI(i) = 0 Then KI(i) = Vp(i) / P
                dKdPi(i) = -Vp(i) / P ^ 2
                Vx(i) = Vy(i) / KI(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            soma_x = 0
            Do
                soma_x = soma_x + Vx(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                Vx(i) = Vx(i) / soma_x
                i = i + 1
            Loop Until i = n + 1

            Do

                LN_CFV = thermobase.CalcLnFug(T, P, Vy, VKij, VTc, VPc, Vw, otherargs)

                Dim cont_int = 1
                Do

                    LN_CFL(i) = thermobase.CalcLnFug(T, P, Vx, VKij, VTc, VPc, Vw, otherargs)

                    i = 0
                    Do
                        KI(i) = (Math.Exp(LN_CFL(i)) / (Math.Exp(LN_CFV(i))))
                        i = i + 1
                    Loop Until i = n + 1

                    marcador = 0
                    If stmp4_ant <> 0 Then
                        marcador = 1
                    End If
                    stmp4_ant = stmp4

                    i = 0
                    stmp4 = 0
                    Do
                        stmp4 = stmp4 + Vy(i) / KI(i)
                        i = i + 1
                    Loop Until i = n + 1

                    i = 0
                    Do
                        Vx(i) = (Vy(i) / KI(i)) / stmp4
                        i = i + 1
                    Loop Until i = n + 1

                    marcador2 = 0
                    If marcador = 1 Then
                        If Math.Abs(stmp4_ant - stmp4) < m_tolerance Then
                            marcador2 = 1
                        End If
                    End If

                    cont_int = cont_int + 1

                    If cont_int >= m_miti Then Throw New Exception(ErrorCode.MaximumIterationsReached)
                    If Double.IsNaN(stmp4) Then Throw New Exception(ErrorCode.IterationDiverged)

                Loop Until marcador2 = 1

                dKdP = _dKdP(T, P, Vx, Vy, VKij, VTc, VPc, Vw, otherargs)

                F = stmp4 - 1

                cnt += 1

                i = 0
                dFdP = 0
                Do
                    dFdP = dFdP - Vy(i) / (KI(i) ^ 2) * dKdP(i)
                    i = i + 1
                Loop Until i = n + 1

                Pant = P
                P = P - F / dFdP

                If cnt >= m_mite Then Throw New Exception(ErrorCode.MaximumIterationsReached)
                If Double.IsNaN(P) Then Throw New Exception(ErrorCode.IterationDiverged)

            Loop Until Math.Abs(F) < m_tolerance

            'check for trivial solution

            'Dim sumk As Double = 0
            'i = 0
            'Do
            '    sumk += KI(i) / n
            '    i = i + 1
            'Loop Until i = n + 1

            'If Abs(sumk - 1) < 0.1 Then

            '    i = 0
            '    P = 0
            '    Do
            '        P = P + Vy(i) / Vp(i)
            '        i = i + 1
            '    Loop Until i = n + 1
            '    P = 1 / P
            'End If

            Dim tmp2(n + 1)
            tmp2(0) = Pant
            i = 0
            Do
                tmp2(i + 1) = KI(i)
                i = i + 1
            Loop Until i = n + 1

            Return tmp2

        End Function

        Function BubP(ByVal T As Double, ByVal Vx As Object, ByVal VKij As Object, ByVal VTc As Object, ByVal VPc As Object, ByVal Vw As Object, ByVal KI As Object, Optional ByVal P As Double = 0, Optional ByVal otherargs As Object = Nothing) As Object

            Dim dFdP As Double
            Dim cnt As Integer = 0

            Dim n = UBound(Vx)

            Dim chk As Boolean = False
            Dim marcador2 As Integer

            Dim F As Double
            Dim dKdP(n), dKdPi(n) As Object
            Dim Vy(n), Vx_ant(n), Vy_ant(n), LN_CFL(n), LN_CFV(n) As Double
            Dim Vp(n), R, coeff(3), tmp(2, 2) As Double
            Dim i As Integer
            Dim soma_y As Double
            Dim marcador3, marcador
            Dim stmp4_ant, stmp4, Pant As Double
            stmp4_ant = 0
            stmp4 = 0
            Pant = 0

            R = 8.314

            i = 0
            Do
                Vp(i) = ppbase.AUX_PVAPi(i, T)
                i = i + 1
            Loop Until i = n + 1

            If P = 0 Then
                'generate first estimate for T
                i = 0
                Do
                    P = P + Vx(i) * Vp(i)
                    i = i + 1
                Loop Until i = n + 1
            End If

            Dim ai_(n)

            i = 0
            Do
                If KI(i) = 0 Then KI(i) = Vp(i) / P
                dKdPi(i) = -Vp(i) / P ^ 2
                Vy(i) = Vx(i) * KI(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            soma_y = 0
            Do
                soma_y = soma_y + Vy(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                Vy(i) = Vy(i) / soma_y
                i = i + 1
            Loop Until i = n + 1

            Do

                LN_CFL = thermobase.CalcLnFug(T, P, Vx, VKij, VTc, VPc, Vw, otherargs)

                marcador3 = 0

                Dim cont_int = 0
                Do

                    LN_CFV = thermobase.CalcLnFug(T, P, Vy, VKij, VTc, VPc, Vw, otherargs)

                    i = 0
                    Do
                        KI(i) = (Math.Exp(LN_CFL(i) - LN_CFV(i)))
                        i = i + 1
                    Loop Until i = n + 1

                    marcador = 0
                    If stmp4_ant <> 0 Then
                        marcador = 1
                    End If
                    stmp4_ant = stmp4

                    i = 0
                    stmp4 = 0
                    Do
                        stmp4 = stmp4 + KI(i) * Vx(i)
                        i = i + 1
                    Loop Until i = n + 1

                    i = 0
                    Do
                        Vy_ant(i) = Vy(i)
                        Vy(i) = KI(i) * Vx(i) / stmp4
                        i = i + 1
                    Loop Until i = n + 1

                    marcador2 = 0
                    If marcador = 1 Then
                        If Math.Abs(stmp4_ant - stmp4) < m_tolerance Then
                            marcador2 = 1
                        End If
                    End If

                    cont_int = cont_int + 1

                    If cont_int >= m_miti Then Throw New Exception(ErrorCode.MaximumIterationsReached)
                    If Double.IsNaN(stmp4) Then Throw New Exception(ErrorCode.IterationDiverged)

                Loop Until marcador2 = 1

                dKdP = _dKdP(T, P, Vx, Vy, VKij, VTc, VPc, Vw, otherargs)

                F = stmp4 - 1

                cnt += 1

                i = 0
                dFdP = 0
                Do
                    dFdP = dFdP + Vx(i) * dKdP(i)
                    i = i + 1
                Loop Until i = n + 1

                Pant = P
                P = P - F / dFdP

                If cnt >= m_mite Then Throw New Exception(ErrorCode.MaximumIterationsReached)
                If Double.IsNaN(P) Then Throw New Exception(ErrorCode.IterationDiverged)

            Loop Until Math.Abs(F) < m_tolerance

            ''check for trivial solution

            'Dim sumk As Double = 0
            'i = 0
            'Do
            '    sumk += KI(i) / n
            '    i = i + 1
            'Loop Until i = n + 1

            'If Abs(sumk - 1) < 0.1 Then

            '    i = 0
            '    P = 0
            '    Do
            '        P = P + Vx(i) * Vp(i)
            '        i = i + 1
            '    Loop Until i = n + 1

            'End If

            Dim tmp2(n + 1)
            tmp2(0) = P
            i = 0
            Do
                tmp2(i + 1) = KI(i)
                i = i + 1
            Loop Until i = n + 1

            Return tmp2

        End Function

        Function DewT(ByVal P As Double, ByVal Vy As Object, ByVal VKij As Object, ByVal VTc As Object, ByVal VPc As Object, ByVal Vw As Object, ByVal KI As Object, Optional ByVal T As Double = 0, Optional ByVal otherargs As Object = Nothing) As Object

            Dim dFdT As Double
            Dim cnt As Integer = 0

            Dim n = UBound(Vy)

            Dim chk As Boolean = False
            Dim marcador2 As Integer

            Dim F As Double
            Dim dKdT(n), dKdTi(n) As Object
            Dim Vx(n), LN_CFL(n), LN_CFV(n) As Double
            Dim Vp(n) As Double
            Dim i As Integer
            Dim soma_x As Double
            Dim marcador
            Dim stmp4_ant, stmp4, Tant As Double
            stmp4_ant = 0
            stmp4 = 0
            Tant = 0

            If T = 0 Then
                'generate first estimate for T
                i = 0
                T = 0
                Do
                    T = T + Vy(i) * ppbase.AUX_TSATi(P, i)
                    i = i + 1
                Loop Until i = n + 1
            End If

            i = 0
            Do
                Vp(i) = ppbase.AUX_PVAPi(i, T)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                KI(i) = VPc(i) / P * Exp(5.42 * (1 - VTc(i) / T))
                dKdTi(i) = 5.42 * KI(i) * VTc(i) / T ^ 2
                Vx(i) = Vy(i) / KI(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            soma_x = 0
            Do
                soma_x = soma_x + Vx(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                Vx(i) = Vx(i) / soma_x
                i = i + 1
            Loop Until i = n + 1

            Do

                i = 0
                Do
                    Vp(i) = ppbase.AUX_PVAPi(i, T)
                    i = i + 1
                Loop Until i = n + 1

                Dim ai_(n)

                LN_CFV = thermobase.CalcLnFug(T, P, Vy, VKij, VTc, VPc, Vw, otherargs)

                Dim cont_int = 1
                Do

                    LN_CFL = thermobase.CalcLnFug(T, P, Vx, VKij, VTc, VPc, Vw, otherargs)

                    i = 0
                    Do
                        KI(i) = (Math.Exp(LN_CFL(i)) / (Math.Exp(LN_CFV(i))))
                        i = i + 1
                    Loop Until i = n + 1

                    marcador = 0
                    If stmp4_ant <> 0 Then
                        marcador = 1
                    End If
                    stmp4_ant = stmp4

                    i = 0
                    stmp4 = 0
                    Do
                        stmp4 = stmp4 + Vy(i) / KI(i)
                        i = i + 1
                    Loop Until i = n + 1

                    i = 0
                    Do
                        Vx(i) = (Vy(i) / KI(i)) / stmp4
                        i = i + 1
                    Loop Until i = n + 1

                    marcador2 = 0
                    If marcador = 1 Then
                        If Math.Abs(stmp4_ant - stmp4) < m_tolerance Then
                            marcador2 = 1
                        End If
                    End If

                    cont_int = cont_int + 1

                    If cont_int >= m_miti Then Throw New Exception(ErrorCode.MaximumIterationsReached)
                    If Double.IsNaN(stmp4) Then Throw New Exception(ErrorCode.IterationDiverged)

                Loop Until marcador2 = 1

                dKdT = _dKdT(T, P, Vx, Vy, VKij, VTc, VPc, Vw, otherargs)

                F = stmp4 - 1

                i = 0
                dFdT = 0
                Do
                    dFdT = dFdT - Vy(i) / (KI(i) ^ 2) * dKdT(i)
                    i = i + 1
                Loop Until i = n + 1

                Tant = T
                T = T - F / dFdT

                cnt += 1

                If cnt >= m_mite Then Throw New Exception(ErrorCode.MaximumIterationsReached)
                If Double.IsNaN(T) Then Throw New Exception(ErrorCode.IterationDiverged)

            Loop Until Math.Abs(F) < m_tolerance

            'check for trivial solution

            'Dim sumk As Double = 0
            'i = 0
            'Do
            '    sumk += KI(i) / n
            '    i = i + 1
            'Loop Until i = n + 1

            'If Abs(sumk - 1) < 0.1 Then

            '    i = 0
            '    T = 0
            '    Do
            '        T = T + Vy(i) * ppbase.AUX_TSATi(P, i)
            '        i = i + 1
            '    Loop Until i = n + 1

            'End If

            Dim tmp2(n + 1)
            tmp2(0) = T
            i = 0
            Do
                tmp2(i + 1) = KI(i)
                i = i + 1
            Loop Until i = n + 1

            Return tmp2

        End Function

        Function BubT(ByVal P As Double, ByVal Vx As Object, ByVal VKij As Object, ByVal VTc As Object, ByVal VPc As Object, ByVal Vw As Object, ByVal KI As Object, Optional ByVal T As Double = 0, Optional ByVal otherargs As Object = Nothing) As Object

            Dim dFdT As Double
            Dim cnt As Integer = 0

            Dim n = UBound(Vx)

            Dim chk As Boolean = False
            Dim marcador2 As Integer

            Dim F As Double
            Dim dKdT(n) As Object
            Dim dKdTi(n) As Object
            Dim Vy(n), Vx_ant(n), Vy_ant(n), LN_CFL(n), LN_CFV(n) As Double
            Dim Vp(n) As Double
            Dim i As Integer
            Dim soma_y As Double
            Dim marcador
            Dim stmp4_ant, stmp4, Tant As Double
            stmp4_ant = 0
            stmp4 = 0
            Tant = 0

            If T = 0 Then
                'generate first estimate for T
                i = 0
                Do
                    If T <= VTc(i) Then T = T + Vx(i) * ppbase.AUX_TSATi(P, i)
                    i = i + 1
                Loop Until i = n + 1
            End If

            i = 0
            Do
                Vp(i) = ppbase.AUX_PVAPi(i, T)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                If KI(i) = 0 Then KI(i) = Vp(i) / P
                dKdTi(i) = 5.42 * KI(i) * VTc(i) / T ^ 2
                Vy(i) = Vx(i) * KI(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            soma_y = 0
            Do
                soma_y = soma_y + Vy(i)
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                Vy(i) = Vy(i) / soma_y
                i = i + 1
            Loop Until i = n + 1

            Do

                i = 0
                Do
                    Vp(i) = ppbase.AUX_PVAPi(i, T)
                    i = i + 1
                Loop Until i = n + 1

                LN_CFL = thermobase.CalcLnFug(T, P, Vx, VKij, VTc, VPc, Vw, otherargs)

                Dim cont_int = 1
                Do

                    LN_CFV = thermobase.CalcLnFug(T, P, Vy, VKij, VTc, VPc, Vw, otherargs)

                    i = 0
                    Do
                        KI(i) = (Math.Exp(LN_CFL(i)) / (Math.Exp(LN_CFV(i))))
                        i = i + 1
                    Loop Until i = n + 1

                    marcador = 0
                    If stmp4_ant <> 0 Then
                        marcador = 1
                    End If
                    stmp4_ant = stmp4

                    i = 0
                    stmp4 = 0
                    Do
                        stmp4 = stmp4 + KI(i) * Vx(i)
                        i = i + 1
                    Loop Until i = n + 1

                    i = 0
                    Do
                        Vy(i) = KI(i) * Vx(i) / stmp4
                        i = i + 1
                    Loop Until i = n + 1

                    marcador2 = 0
                    If marcador = 1 Then
                        If Math.Abs(stmp4_ant - stmp4) < m_tolerance Then
                            marcador2 = 1
                        End If
                    End If

                    cont_int += 1

                    If cont_int >= m_miti Then Throw New Exception(ErrorCode.MaximumIterationsReached)
                    If Double.IsNaN(stmp4) Then Throw New Exception(ErrorCode.IterationDiverged)

                Loop Until marcador2 = 1

                dKdT = _dKdT(T, P, Vx, Vy, VKij, VTc, VPc, Vw)

                F = stmp4 - 1

                i = 0
                dFdT = 0
                Do
                    dFdT = dFdT + Vx(i) * dKdT(i)
                    i = i + 1
                Loop Until i = n + 1

                Tant = T
                T = T - F / dFdT

                cnt += 1

                If cnt >= m_mite Then Throw New Exception(ErrorCode.MaximumIterationsReached)
                If Double.IsNaN(T) Then Throw New Exception(ErrorCode.IterationDiverged)

            Loop Until Math.Abs(F) < m_tolerance

            'check for trivial solution

            'Dim sumk As Double = 0
            'i = 0
            'Do
            '    sumk += KI(i) / n
            '    i = i + 1
            'Loop Until i = n + 1

            'If Abs(sumk - 1) < 0.1 Then

            '    i = 0
            '    T = 0
            '    Do
            '        T = T + Vx(i) * ppbase.AUX_TSATi(P, i)
            '        i = i + 1
            '    Loop Until i = n + 1

            'End If

            Dim tmp2(n + 1)
            tmp2(0) = Tant
            i = 0
            Do
                tmp2(i + 1) = KI(i)
                i = i + 1
            Loop Until i = n + 1

            Return tmp2

        End Function

#End Region

#Region "        Critical Point General Calculation Routines (EXPERIMENTAL)"

        Public Function dlnfug_i_dn_j(ByVal jidx As Integer, ByVal T As Double, ByVal V As Double, ByVal Vz As Array, ByVal VKij As Object, ByVal VTc As Array, ByVal VPc As Array, ByVal Vw As Array, Optional ByVal otherargs As Object = Nothing)

            'calculate numerical derivatives with a 4-point differencing scheme

            Dim n As Integer = UBound(Vz)

            Dim i As Integer

            Dim mres(n) As Double

            Dim points1(n), points2(n), points3(n), points4(n) As Double

            Dim h As Double = 0.0001

            points1 = thermobase.CalcLnFugTV(T, V, perturb_n(jidx, -2 * h, Vz), VKij, VTc, VPc, Vw, otherargs)
            points2 = thermobase.CalcLnFugTV(T, V, perturb_n(jidx, -h, Vz), VKij, VTc, VPc, Vw, otherargs)
            points3 = thermobase.CalcLnFugTV(T, V, perturb_n(jidx, h, Vz), VKij, VTc, VPc, Vw, otherargs)
            points4 = thermobase.CalcLnFugTV(T, V, perturb_n(jidx, 2 * h, Vz), VKij, VTc, VPc, Vw, otherargs)

            i = 0
            Do
                mres(i) = (points1(i) - 8 * points2(i) + 8 * points3(i) + points4(i)) / (12 * h)
                i = i + 1
            Loop Until i = n + 1

            Return mres

        End Function

        Public Function d2lnfug_i_dn_j_dn_k(ByVal jidx As Integer, ByVal kidx As Integer, ByVal T As Double, ByVal V As Double, ByVal Vz As Array, ByVal VKij As Object, ByVal VTc As Array, ByVal VPc As Array, ByVal Vw As Array, Optional ByVal otherargs As Object = Nothing)

            'calculate numerical derivatives with a 4-point differencing scheme

            Dim n As Integer = UBound(Vz)

            Dim i As Integer

            Dim mres(n) As Double

            Dim points1(n), points2(n), points3(n), points4(n) As Double

            Dim h As Double = 0.0001

            points1 = dlnfug_i_dn_j(jidx, T, V, perturb_n(kidx, -2 * h, Vz), VKij, VTc, VPc, Vw, otherargs)
            points2 = dlnfug_i_dn_j(jidx, T, V, perturb_n(kidx, -h, Vz), VKij, VTc, VPc, Vw, otherargs)
            points3 = dlnfug_i_dn_j(jidx, T, V, perturb_n(kidx, h, Vz), VKij, VTc, VPc, Vw, otherargs)
            points4 = dlnfug_i_dn_j(jidx, T, V, perturb_n(kidx, 2 * h, Vz), VKij, VTc, VPc, Vw, otherargs)

            i = 0
            Do
                mres(i) = (points1(i) - 8 * points2(i) + 8 * points3(i) + points4(i)) / (12 * h)
                i = i + 1
            Loop Until i = n + 1

            Return mres

        End Function

        Private Function perturb_n(ByVal i As Integer, ByVal h As Double, ByVal Vx As Array)

            Dim n As Integer = UBound(Vx)
            Dim j As Integer = 0

            Dim ntot As Double = 1

            Dim Vn(n), Vn2(n)

            j = 0
            Do
                Vn(j) = Vx(j) * ntot
                j = j + 1
            Loop Until j = n + 1

            Vn(i) = Vn(i) + h

            j = 0
            Do
                Vn2(j) = Vn(j) / (ntot + h)
                j = j + 1
            Loop Until j = n + 1


            Return Vn2

        End Function

        Private Function Qij(ByVal T As Double, ByVal V As Double, ByVal Vz As Array, ByVal VTc As Array, ByVal VPc As Array, ByVal Vw As Array, ByVal VKIj As Array, Optional ByVal otherargs As Object = Nothing) As Mapack.Matrix

            Dim n As Integer = UBound(Vz)

            Dim mat As Mapack.Matrix = New Mapack.Matrix(n + 1, n + 1)
            Dim el(n) As Object


            Dim i, j As Integer

            i = 0
            Do
                el(i) = dlnfug_i_dn_j(i, T, V, Vz, VKIj, VTc, VPc, Vw, otherargs)
                i = i + 1
            Loop Until i = n + 1


            i = 0
            Do
                j = 0
                Do
                    mat(i, j) = el(i)(j)
                    j = j + 1
                Loop Until j = n + 1
                i = i + 1
            Loop Until i = n + 1

            Return mat

        End Function

        Private Function QijDetBrent(ByVal x As Double, ByVal otherargs As Object) As Double

            Dim T As Double = x

            'ByVal T As Double, ByVal V As Double, ByVal Vz As Array, ByVal VTc As Array, 
            'ByVal VPc As Array, ByVal Vw As Array, ByVal VKIj As Array, Optional ByVal otherargs

            Dim V As Double = otherargs(0)
            Dim Vz As Array = otherargs(1)
            Dim VTc As Array = otherargs(2)
            Dim VPc As Array = otherargs(3)
            Dim Vw As Array = otherargs(4)
            Dim VKij As Object = otherargs(5)
            Dim oa1 As Object = otherargs(6)

            Dim mat As Mapack.Matrix = Me.Qij(T, V, Vz, VTc, VPc, Vw, VKij, oa1)

            Return mat.Determinant

        End Function

        Private Function TripleSum(ByVal Dn As Array, ByVal T As Double, ByVal V As Double, ByVal Vz As Array, ByVal VTc As Array, ByVal VPc As Array, ByVal Vw As Array, ByVal VKIj As Object, Optional ByVal otherargs As Object = Nothing) As Double

            Dim n As Integer = UBound(Vz)

            Dim mat(n, n) As Object
            Dim el(n) As Object


            Dim i, j, k As Integer

            i = 0
            Do
                j = 0
                Do
                    mat(i, j) = d2lnfug_i_dn_j_dn_k(i, j, T, V, Vz, VKIj, VTc, VPc, Vw, otherargs)
                    j = j + 1
                Loop Until j = n + 1
                i = i + 1
            Loop Until i = n + 1

            Dim ts As Double = 0.0

            i = 0
            Do
                j = 0
                Do
                    k = 0
                    Do
                        ts += mat(j, k)(i) * Dn(i) * Dn(j) * Dn(k) * 8.314 * T
                        k = k + 1
                    Loop Until k = n + 1
                    j = j + 1
                Loop Until j = n + 1
                i = i + 1
            Loop Until i = n + 1

            Return ts

        End Function

        Private Function TripleSum2(ByVal V As Double, ByVal Vz As Array, ByVal VTc As Array, ByVal VPc As Array, ByVal Vw As Array, ByVal VKIj As Object, Optional ByVal otherargs As Object = Nothing) As Double

            Dim T, Tc_sup, Tc_inf As Double

            Dim i As Integer
            Dim n As Double

            n = UBound(Vz)

            Dim Dn(n)

            Tc_inf = Common.Min(VTc) * 0.5
            Tc_sup = Common.Max(VTc) * 1.5

            Dim nsub, delta_Tc As Double
            delta_Tc = (Tc_sup - Tc_inf) / nsub
            nsub = 10

            Dim brentsolver As New BrentOpt.Brent
            brentsolver.DefineFuncDelegate(AddressOf QijDetBrent)

            Dim args As Object = New Object() {V, Vz, VTc, VPc, Vw, VKIj, otherargs}

            T = brentsolver.BrentOpt(Tc_inf, Tc_sup, nsub, 1.0E-20, 100, args)

            Dim MA As Mapack.Matrix, Dn0(n) As Double
            Dim MA_(n, n), MB_(n), Dn0_(n) As Single

            'Dim MP As New DLLXnumbers.Xnumbers
            MA = Qij(T, V, Vz, VTc, VPc, Vw, VKIj, otherargs)

            Dim m2 As Mapack.Matrix = New Mapack.Matrix(MA.Rows, 1)

            i = 0
            Do
                m2(i, 0) = 0.001
                i = i + 1
            Loop Until i = n + 1

            Try
                Dim m3 As Mapack.Matrix = MA.Solve(m2)
                i = 0
                Do
                    Dn0(i) = m3(i, 0)
                    i = i + 1
                Loop Until i = n + 1
            Catch ex As Exception
                i = 0
                Do
                    Dn0(i) = 0
                    i = i + 1
                Loop Until i = n + 1
            End Try

            Dim soma_Dn = 0
            i = 0
            Do
                soma_Dn += (Dn0(i) ^ 2) ^ 0.5
                i = i + 1
            Loop Until i = n + 1

            i = 0
            Do
                Dn(i) = Dn0(i) / soma_Dn
                i = i + 1
            Loop Until i = n + 1

            Return Me.TripleSum(Dn, T, V, Vz, VTc, VPc, Vw, VKIj, otherargs)

        End Function

        Private Function TripleSumBrent(ByVal x As Double, ByVal otherargs As Object) As Double

            Dim V As Double = x
            Dim Vz As Array = otherargs(0)
            Dim VTc As Array = otherargs(1)
            Dim VPc As Array = otherargs(2)
            Dim Vw As Array = otherargs(3)
            Dim VKij As Object = otherargs(4)
            Dim oa1 As Object = otherargs(5)

            Return Me.TripleSum2(V, Vz, VTc, VPc, Vw, VKij, oa1)

        End Function

        Function CriticalPoints(ByVal Vz As Array, ByVal VTc As Array, ByVal VPc As Array, ByVal Vw As Array, ByVal VKIj As Array, Optional ByVal otherargs As Object = Nothing) As ArrayList

            Dim res As New ArrayList

            Dim V, Vc_sup, Vc_inf As Double

            Dim stmp(2)
            Dim n, R As Double
            Dim i, nsub As Integer

            n = UBound(Vz)

            Dim b As Double

            R = 8.314
            i = 0
            b = 0
            Do
                b += Vz(i) * 0.0778 * R * VTc(i) / VPc(i)
                i = i + 1
            Loop Until i = n + 1

            Vc_inf = b * 5
            Vc_sup = b
            nsub = 25

            Dim brentsolver As New BrentOpt.Brent
            brentsolver.DefineFuncDelegate(AddressOf TripleSumBrent)

            Dim args As Object = New Object() {Vz, VTc, VPc, Vw, VKIj, otherargs}

restart:    V = brentsolver.BrentOpt(Vc_inf, Vc_sup, nsub, 1.0E-20, 100, args)

            If Abs(Vc_sup - Vc_inf) < 0.000001 Then GoTo Final2

            If Double.IsNaN(V) Then
                Vc_inf += (Vc_sup - Vc_inf) / nsub
                GoTo restart
            End If

            Dim T, P As Double

            Dim Tc_sup, Tc_inf As Double

            Tc_inf = Common.Min(VTc) * 0.5
            Tc_sup = Common.Max(VTc) * 1.5

            brentsolver = New BrentOpt.Brent
            brentsolver.DefineFuncDelegate(AddressOf QijDetBrent)

            args = New Object() {V, Vz, VTc, VPc, Vw, VKIj, otherargs}

restart2:   T = brentsolver.BrentOpt(Tc_inf, Tc_sup, nsub, 1.0E-20, 100, args)

            If Double.IsNaN(T) Then
                Tc_inf += (Tc_sup - Tc_inf) / nsub
                GoTo restart2
            End If

            P = thermobase.CalcP(V, T, Vz, VKIj, VTc, VPc, Vw, otherargs)

            'If Abs(Vc_sup - Vc_inf) < 0.000001 Then GoTo Final2

            If P < 0 Then
                Vc_inf += (Vc_sup - Vc_inf) / nsub
                GoTo restart
            End If

            stmp(0) = T
            stmp(1) = P
            stmp(2) = V

            'test stability of the critical point
            Dim stresult As Object
            stresult = Me.StabTest(T, P, Vz, VKIj, VTc, VPc, Vw, Nothing, otherargs)

            'add the calculated CP to results only if it is stable
            If stresult(0) = True Then res.Add(stmp.Clone)

            'If Vc_inf <= b Then
            '    GoTo Final2
            'Else
            '    Vc_inf += 2 * (Vc_sup - Vc_inf) / nsub
            '    GoTo restart
            'End If

Final2:

            Return res

        End Function

        Function StabilityCurve(ByVal Vz As Object, ByVal VTc As Object, ByVal VPc As Object, ByVal VVc As Object, ByVal Vw As Object, ByVal VKIj As Object, Optional ByVal Vmax As Double = 0, Optional ByVal delta As Double = 40, Optional ByVal multipl As Integer = 15, Optional ByVal otherargs As Object = Nothing) As ArrayList

            Dim V, Vmin, deltaV As Double

            Dim stmp(2)
            Dim n, R, P, T As Double
            Dim i As Integer

            n = UBound(Vz)

            Dim Tc(n), Pc(n)
            Dim b As Double

            'estimar temperatura e pressão críticas iniciais

            R = 8.314

            i = 0
            Do
                If Vz(i) <> 0 Then
                    Tc(i) = VTc(i)
                    Pc(i) = VPc(i)
                End If
                i = i + 1
            Loop Until i = n + 1

            i = 0
            b = 0
            Do
                b += Vz(i) * 0.0778 * R * Tc(i) / Pc(i)
                i = i + 1
            Loop Until i = n + 1

            'estimar temperatura e pressão críticas iniciais

            If Vmax = 0 Then Vmax = b * multipl
            Vmin = b * 1.05

            deltaV = (Vmax - Vmin) / 500 ' delta

            Dim result As ArrayList = New ArrayList()

            Dim Tc_sup, Tc_inf As Double

            Dim brentsolver As New BrentOpt.Brent
            brentsolver.DefineFuncDelegate(AddressOf QijDetBrent)

            V = Vmax
            Do

                Tc_inf = Common.Min(VTc) * 0.5
                Tc_sup = Common.Max(VTc) * 1.5
                T = brentsolver.BrentOpt(Tc_inf, Tc_sup, 20, 0.0000000001, 100, otherargs)
                P = Me.thermobase.CalcP(V, T, Vz, VKIj, VTc, VPc, Vw)
                If P < 0 Then Exit Do
                result.Add(New Object() {T, P})
                V -= deltaV
            Loop Until V <= Vmin

            Return result

        End Function

#End Region

    End Class

End Namespace
