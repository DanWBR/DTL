'    DWSIM Nested Loops Flash Algorithms
'    Copyright 2010-2015 Daniel Wagner O. de Medeiros, Gregor Reichert
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
Imports DTL.DTL.SimulationObjects
Imports DTL.DTL.MathEx
Imports DTL.DTL.MathEx.Common
Imports System.Threading.Tasks
Imports System.Linq

Namespace DTL.SimulationObjects.PropertyPackages.Auxiliary.FlashAlgorithms

    ''' <summary>
    ''' The Flash algorithms in this class are based on the Nested Loops approach to solve equilibrium calculations.
    ''' </summary>
    ''' <remarks></remarks>
    <System.Serializable()> Public Class DWSIMDefault

        Inherits FlashAlgorithm

        Dim etol As Double = 0.000001
        Dim itol As Double = 0.000001
        Dim maxit_i As Integer = 100
        Dim maxit_e As Integer = 100
        Dim Hv0, Hvid, Hlid, Hf, Hv, Hl As Double
        Dim Sv0, Svid, Slid, Sf, Sv, Sl As Double

        Public Overrides Function Flash_PT(ByVal Vz As Double(), ByVal P As Double, ByVal T As Double, ByVal PP As PropertyPackages.PropertyPackage, Optional ByVal ReuseKI As Boolean = False, Optional ByVal PrevKi As Double() = Nothing) As Object

            Dim i, n, ecount As Integer
            Dim Pb, Pd, Pmin, Pmax, Px As Double
            Dim d1, d2 As Date, dt As TimeSpan
            Dim L, V, Vant As Double

            d1 = Date.Now

            etol = CDbl(PP.Parameters("PP_PTFELT"))
            maxit_e = CInt(PP.Parameters("PP_PTFMEI"))
            itol = CDbl(PP.Parameters("PP_PTFILT"))
            maxit_i = CInt(PP.Parameters("PP_PTFMII"))

            n = UBound(Vz)

            Dim Vn(n) As String, Vx(n), Vy(n), Vx_ant(n), Vy_ant(n), Vp(n), Ki(n), Ki_ant(n), fi(n) As Double

            Vn = PP.RET_VNAMES()
            fi = Vz.Clone

            'Calculate Ki`s

            If Not ReuseKI Then
                i = 0
                Do
                    Vp(i) = PP.AUX_PVAPi(i, T)
                    Ki(i) = Vp(i) / P
                    i += 1
                Loop Until i = n + 1
            Else
                For i = 0 To n
                    Vp(i) = PP.AUX_PVAPi(i, T)
                    Ki(i) = PrevKi(i)
                Next
            End If

            'Estimate V

            If T > DTL.MathEx.Common.Max(PP.RET_VTC, Vz) Then
                Vy = Vz
                V = 1
                L = 0
                GoTo out
            End If

            i = 0
            Px = 0
            Do
                If Vp(i) <> 0.0# Then Px = Px + (Vz(i) / Vp(i))
                i = i + 1
            Loop Until i = n + 1
            Px = 1 / Px
            Pmin = Px
            i = 0
            Px = 0
            Do
                Px = Px + Vz(i) * Vp(i)
                i = i + 1
            Loop Until i = n + 1
            Pmax = Px
            Pb = Pmax
            Pd = Pmin

            If Abs(Pb - Pd) / Pb < 0.0000001 Then
                'one comp only
                If Px <= P Then
                    L = 1
                    V = 0
                    Vx = Vz
                    GoTo out
                Else
                    L = 0
                    V = 1
                    Vy = Vz
                    GoTo out
                End If
            End If

            Dim Vmin, Vmax, g As Double
            Vmin = 1.0#
            Vmax = 0.0#
            For i = 0 To n
                If (Ki(i) * Vz(i) - 1) / (Ki(i) - 1) < Vmin Then Vmin = (Ki(i) * Vz(i) - 1) / (Ki(i) - 1)
                If (1 - Vz(i)) / (1 - Ki(i)) > Vmax Then Vmax = (1 - Vz(i)) / (1 - Ki(i))
            Next

            If Vmin < 0.0# Then Vmin = 0.0#
            If Vmin = 1.0# Then Vmin = 0.0#
            If Vmax = 0.0# Then Vmax = 1.0#
            If Vmax > 1.0# Then Vmax = 1.0#

            V = (Vmin + Vmax) / 2

            g = 0.0#
            For i = 0 To n
                g += Vz(i) * (Ki(i) - 1) / (V + (1 - V) * Ki(i))
            Next

            If g > 0 Then Vmin = V Else Vmax = V

            V = Vmin + (Vmax - Vmin) / 2
            'V = (P - Pd) / (Pb - Pd)

            L = 1 - V

            If n = 0 Then
                If Vp(0) <= P Then
                    L = 1
                    V = 0
                Else
                    L = 0
                    V = 1
                End If
            End If

            i = 0
            Do
                If Vz(i) <> 0 Then
                    Vy(i) = Vz(i) * Ki(i) / ((Ki(i) - 1) * V + 1)
                    If Ki(i) <> 0 Then Vx(i) = Vy(i) / Ki(i) Else Vx(i) = Vz(i)
                    If Vy(i) < 0 Then Vy(i) = 0
                    If Vx(i) < 0 Then Vx(i) = 0
                Else
                    Vy(i) = 0
                    Vx(i) = 0
                End If
                i += 1
            Loop Until i = n + 1

            Vy_ant = Vy.Clone
            Vx_ant = Vx.Clone

            Vy = Vz.MultiplyY(Ki).DivideY(Ki.AddConstY(-1).MultiplyConstY(V).AddConstY(1))
            Vx = Vy.DivideY(Ki)

            Vx = Vx.NormalizeY
            Vy = Vy.NormalizeY

            ecount = 0
            Dim convergiu As Integer = 0
            Dim F, dF, e1, e2, e3 As Double

            Do

                Ki_ant = Ki.Clone
                Ki = PP.DW_CalcKvalue(Vx, Vy, T, P)

                'i = 0
                'Do
                '    If Vz(i) <> 0 Then
                '        Vy_ant(i) = Vy(i)
                '        Vx_ant(i) = Vx(i)
                '        Vy(i) = Vz(i) * Ki(i) / ((Ki(i) - 1) * V + 1)
                '        Vx(i) = Vy(i) / Ki(i)
                '    Else
                '        Vy(i) = 0
                '        Vx(i) = 0
                '    End If
                '    i += 1
                'Loop Until i = n + 1

                Vy_ant = Vy.Clone
                Vx_ant = Vx.Clone

                Vy = Vz.MultiplyY(Ki).DivideY(Ki.AddConstY(-1).MultiplyConstY(V).AddConstY(1))
                Vx = Vy.DivideY(Ki)

                Vx = Vx.NormalizeY
                Vy = Vy.NormalizeY

                e1 = Vx.SubtractY(Vx_ant).AbsSumY
                e2 = Vy.SubtractY(Vy_ant).AbsSumY

                e3 = (V - Vant)

                If Double.IsNaN(e1 + e2) Then

                    Throw New Exception(DTL.App.GetLocalString("PropPack_FlashError"))

                ElseIf Math.Abs(e3) < 0.0000000001 And ecount > 0 Then

                    convergiu = 1

                    Exit Do

                Else

                    Vant = V

                    'F = 0.0#
                    'dF = 0.0#
                    'i = 0
                    'Do
                    '    If Vz(i) > 0 Then
                    '        F = F + Vz(i) * (Ki(i) - 1) / (1 + V * (Ki(i) - 1))
                    '        dF = dF - Vz(i) * (Ki(i) - 1) ^ 2 / (1 + V * (Ki(i) - 1)) ^ 2
                    '    End If
                    '    i = i + 1
                    'Loop Until i = n + 1

                    F = Vz.MultiplyY(Ki.AddConstY(-1).DivideY(Ki.AddConstY(-1).MultiplyConstY(V).AddConstY(1))).SumY
                    dF = Vz.NegateY.MultiplyY(Ki.AddConstY(-1).MultiplyY(Ki.AddConstY(-1)).DivideY(Ki.AddConstY(-1).MultiplyConstY(V).AddConstY(1)).DivideY(Ki.AddConstY(-1).MultiplyConstY(V).AddConstY(1))).SumY

                    If Abs(F) < etol / 100 Then Exit Do

                    V = -F / dF + Vant

                    'If V >= 1.01 Or V <= -0.01 Then V = -0.1 * F / dF + Vant

                End If

                If V < 0.0# Then V = 0.0#
                If V > 1.0# Then V = 1.0#

                L = 1 - V

                ecount += 1

                If Double.IsNaN(V) Then
                    Throw New Exception(DTL.App.GetLocalString("PropPack_FlashTPVapFracError"))
                End If
                If ecount > maxit_e Then
                    Throw New Exception(DTL.App.GetLocalString("PropPack_FlashMaxIt2"))
                End If

                WriteDebugInfo("PT Flash [NL]: Iteration #" & ecount & ", VF = " & V)

            Loop Until convergiu = 1

            d2 = Date.Now

            dt = d2 - d1

            WriteDebugInfo("PT Flash [NL]: Converged in " & ecount & " iterations. Time taken: " & dt.TotalMilliseconds & " ms. Error function value: " & F)

out:        Return New Object() {L, V, Vx, Vy, ecount, 0.0#, PP.RET_NullVector, 0.0#, PP.RET_NullVector}

        End Function

        Public Overrides Function Flash_PH(ByVal Vz As Double(), ByVal P As Double, ByVal H As Double, ByVal Tref As Double, ByVal PP As PropertyPackages.PropertyPackage, Optional ByVal ReuseKI As Boolean = False, Optional ByVal PrevKi As Double() = Nothing) As Object
            If PP.Parameters("PP_FLASHALGORITHMFASTMODE") = 1 Then
                Return Flash_PH_1(Vz, P, H, Tref, PP, ReuseKI, PrevKi)
            Else
                Return Flash_PH_2(Vz, P, H, Tref, PP, ReuseKI, PrevKi)
            End If
        End Function

        Public Overrides Function Flash_PS(ByVal Vz As Double(), ByVal P As Double, ByVal S As Double, ByVal Tref As Double, ByVal PP As PropertyPackages.PropertyPackage, Optional ByVal ReuseKI As Boolean = False, Optional ByVal PrevKi As Double() = Nothing) As Object
            If PP.Parameters("PP_FLASHALGORITHMFASTMODE") = 1 Then
                Return Flash_PS_1(Vz, P, S, Tref, PP, ReuseKI, PrevKi)
            Else
                Return Flash_PS_2(Vz, P, S, Tref, PP, ReuseKI, PrevKi)
            End If
        End Function

        Public Function Flash_PH_1(ByVal Vz As Double(), ByVal P As Double, ByVal H As Double, ByVal Tref As Double, ByVal PP As PropertyPackages.PropertyPackage, Optional ByVal ReuseKI As Boolean = False, Optional ByVal PrevKi As Double() = Nothing) As Object

            Dim doparallel As Boolean = My.MyApplication._EnableParallelProcessing

            Dim i, j, n, ecount As Integer
            Dim d1, d2 As Date, dt As TimeSpan
            Dim L, V, T, Pf As Double

            d1 = Date.Now

            n = UBound(Vz)

            PP = PP
            Hf = H
            Pf = P

            Dim Vn(n) As String, Vx(n), Vy(n), Vx_ant(n), Vy_ant(n), Vp(n), Ki(n), Ki_ant(n), fi(n) As Double

            Vn = PP.RET_VNAMES()
            fi = Vz.Clone

            Dim maxitINT As Integer = CInt(PP.Parameters("PP_PHFMII"))
            Dim maxitEXT As Integer = CInt(PP.Parameters("PP_PHFMEI"))
            Dim tolINT As Double = CDbl(PP.Parameters("PP_PHFILT"))
            Dim tolEXT As Double = CDbl(PP.Parameters("PP_PHFELT"))

            Dim Tmin, Tmax, epsilon(4), maxDT As Double

            Tmax = 4000.0#
            Tmin = 50.0#
            maxDT = 30.0#

            epsilon(0) = 0.1
            epsilon(1) = 0.01
            epsilon(2) = 0.001
            epsilon(3) = 0.0001
            epsilon(4) = 0.00001

            Dim fx, fx2, dfdx, x1, dx As Double

            Dim cnt As Integer

            If Tref = 0.0# Then Tref = 298.15

            For j = 0 To 4

                cnt = 0
                x1 = Tref

                Do

                    If My.MyApplication._EnableParallelProcessing Then
                        My.MyApplication.IsRunningParallelTasks = True
                        If My.MyApplication._EnableGPUProcessing Then
                            'My.MyApplication.gpu.EnableMultithreading()
                        End If
                        Try
                            Dim task1 = Task.Factory.StartNew(Sub()
                                                                  fx = Herror("PT", x1, P, Vz, PP)(0)
                                                              End Sub)
                            Dim task2 = Task.Factory.StartNew(Sub()
                                                                  fx2 = Herror("PT", x1 + epsilon(j), P, Vz, PP)(0)
                                                              End Sub)
                            Task.WaitAll(task1, task2)
                        Catch ae As AggregateException
                            Throw ae.Flatten().InnerException
                        Finally
                            'If My.MyApplication._EnableGPUProcessing Then
                            '    My.MyApplication.gpu.DisableMultithreading()
                            '    My.MyApplication.gpu.FreeAll()
                            'End If
                        End Try
                        My.MyApplication.IsRunningParallelTasks = False
                    Else
                        fx = Herror("PT", x1, P, Vz, PP)(0)
                        fx2 = Herror("PT", x1 + epsilon(j), P, Vz, PP)(0)
                    End If

                    If Abs(fx) <= tolEXT Then Exit Do

                    dfdx = (fx2 - fx) / epsilon(j)
                    dx = fx / dfdx

                    If Abs(dx) > maxDT Then dx = maxDT * Sign(dx)

                    x1 = x1 - dx

                    cnt += 1

                Loop Until cnt > maxitEXT Or Double.IsNaN(x1) Or x1 < 0.0#

                T = x1

                If Not Double.IsNaN(T) And Not Double.IsInfinity(T) And Not cnt > maxitEXT Then
                    If T > Tmin And T < Tmax Then Exit For
                End If

            Next

            If Double.IsNaN(T) Or T <= Tmin Or T >= Tmax Or cnt > maxitEXT Or Abs(fx) > tolEXT Then Throw New Exception("PH Flash [NL]: Invalid result: Temperature did not converge.")

            Dim tmp As Object = Flash_PT(Vz, P, T, PP)

            L = tmp(0)
            V = tmp(1)
            Vx = tmp(2)
            Vy = tmp(3)
            ecount = tmp(4)

            For i = 0 To n
                Ki(i) = Vy(i) / Vx(i)
            Next

            d2 = Date.Now

            dt = d2 - d1

            WriteDebugInfo("PH Flash [NL]: Converged in " & ecount & " iterations. Time taken: " & dt.TotalMilliseconds & " ms.")

            Return New Object() {L, V, Vx, Vy, T, ecount, Ki, 0.0#, PP.RET_NullVector, 0.0#, PP.RET_NullVector}

        End Function

        Public Function Flash_PH_2(ByVal Vz As Double(), ByVal P As Double, ByVal H As Double, ByVal Tref As Double, ByVal PP As PropertyPackages.PropertyPackage, Optional ByVal ReuseKI As Boolean = False, Optional ByVal PrevKi As Double() = Nothing) As Object

            Dim doparallel As Boolean = My.MyApplication._EnableParallelProcessing

            Dim i, n, ecount As Integer
            Dim d1, d2 As Date, dt As TimeSpan
            Dim L, V, T, Pf As Double

            Dim resultFlash As Object
            Dim Tb, Td, Hb, Hd As Double
            Dim ErrRes As Object

            d1 = Date.Now

            n = UBound(Vz)

            Hf = H
            Pf = P

            Dim Vn(n) As String, Vx(n), Vy(n), Vx_ant(n), Vy_ant(n), Vp(n), Ki(n), Ki_ant(n), fi(n) As Double

            Vn = PP.RET_VNAMES()
            fi = Vz.Clone

            Dim maxitINT As Integer = CInt(PP.Parameters("PP_PHFMII"))
            Dim maxitEXT As Integer = CInt(PP.Parameters("PP_PHFMEI"))
            Dim tolINT As Double = CDbl(PP.Parameters("PP_PHFILT"))
            Dim tolEXT As Double = CDbl(PP.Parameters("PP_PHFELT"))

            Dim Tmin, Tmax As Double

            Tmax = 2000.0#
            Tmin = 50.0#

            If Tref = 0.0# Then Tref = 298.15

            'calculate dew point and boiling point

            Dim alreadymt As Boolean = False

            If My.MyApplication._EnableParallelProcessing Then
                My.MyApplication.IsRunningParallelTasks = True
                Try
                    Dim task1 = Task.Factory.StartNew(Sub()
                                                          Dim ErrRes1 = Herror("PV", 0, P, Vz, PP)
                                                          Hb = ErrRes1(0)
                                                          Tb = ErrRes1(1)
                                                      End Sub)
                    Dim task2 = Task.Factory.StartNew(Sub()
                                                          Dim ErrRes2 = Herror("PV", 1, P, Vz, PP)
                                                          Hd = ErrRes2(0)
                                                          Td = ErrRes2(1)
                                                      End Sub)
                    Task.WaitAll(task1, task2)
                Catch ae As AggregateException
                    Throw ae.Flatten().InnerException
                End Try
                My.MyApplication.IsRunningParallelTasks = False
            Else
                ErrRes = Herror("PV", 0, P, Vz, PP)
                Hb = ErrRes(0)
                Tb = ErrRes(1)
                ErrRes = Herror("PV", 1, P, Vz, PP)
                Hd = ErrRes(0)
                Td = ErrRes(1)
            End If

            If Hb > 0 And Hd < 0 Then

                'specified enthalpy requires partial evaporation 
                'calculate vapour fraction

                Dim H1, H2, V1, V2 As Double
                ecount = 0
                V = 0
                H1 = Hb
                Do

                    ecount += 1
                    V1 = V
                    If V1 < 1 Then
                        V2 = V1 + 0.01
                    Else
                        V2 = V1 - 0.01
                    End If
                    H2 = Herror("PV", V2, P, Vz, PP)(0)
                    V = V1 + (V2 - V1) * (0 - H1) / (H2 - H1)
                    If V < 0 Then V = 0.0#
                    If V > 1 Then V = 1.0#
                    resultFlash = Herror("PV", V, P, Vz, PP)
                    H1 = resultFlash(0)
                Loop Until Abs(H1) < itol Or ecount > maxitEXT

                T = resultFlash(1)
                L = resultFlash(3)
                Vy = resultFlash(4)
                Vx = resultFlash(5)

                For i = 0 To n
                    Ki(i) = Vy(i) / Vx(i)
                Next

            ElseIf Hd > 0 Then

                'only gas phase
                'calculate temperature

                Dim H1, H2, T1, T2 As Double
                ecount = 0
                T = Td
                H1 = Hd
                Do
                    ecount += 1
                    T1 = T
                    T2 = T1 + 1
                    H2 = Hf - PP.DW_CalcEnthalpy(Vz, T2, P, State.Vapor)
                    T = T1 + (T2 - T1) * (0 - H1) / (H2 - H1)
                    H1 = Hf - PP.DW_CalcEnthalpy(Vz, T, P, State.Vapor)
                Loop Until Abs(H1) < itol Or ecount > maxitEXT

                L = 0
                V = 1
                Vy = Vz.Clone
                Vx = Vz.Clone
                L = 0
                For i = 0 To n
                    Ki(i) = 1
                Next

                If T <= Tmin Or T >= Tmax Or ecount > maxitEXT Then Throw New Exception("PH Flash [NL3PV3]: Invalid result: Temperature did not converge.")
            Else

                'specified enthalpy requires pure liquid 
                'calculate temperature

                Dim H1, H2, T1, T2 As Double
                ecount = 0
                T = Tb
                H1 = Hb
                Do
                    ecount += 1
                    T1 = T
                    T2 = T1 - 1
                    H2 = Herror("PT", T2, P, Vz, PP)(0)
                    T = T1 + (T2 - T1) * (0 - H1) / (H2 - H1)
                    resultFlash = Herror("PT", T, P, Vz, PP)
                    H1 = resultFlash(0)
                Loop Until Abs(H1) < itol Or ecount > maxitEXT

                V = 0
                L = resultFlash(3)
                Vy = resultFlash(4)
                Vx = resultFlash(5)

                For i = 0 To n
                    Ki(i) = Vy(i) / Vx(i)
                Next

                If T <= Tmin Or T >= Tmax Or ecount > maxitEXT Then Throw New Exception("PH Flash [NL]: Invalid result: Temperature did not converge.")
            End If

            d2 = Date.Now

            dt = d2 - d1

            WriteDebugInfo("PH Flash [NL]: Converged in " & ecount & " iterations. Time taken: " & dt.TotalMilliseconds & " ms")

            Return New Object() {L, V, Vx, Vy, T, ecount, Ki, 0.0#, PP.RET_NullVector, 0.0#, PP.RET_NullVector}

        End Function

        Public Function Flash_PS_1(ByVal Vz As Double(), ByVal P As Double, ByVal S As Double, ByVal Tref As Double, ByVal PP As PropertyPackages.PropertyPackage, Optional ByVal ReuseKI As Boolean = False, Optional ByVal PrevKi As Double() = Nothing) As Object

            Dim doparallel As Boolean = My.MyApplication._EnableParallelProcessing

            Dim Vn(1) As String, Vx(1), Vy(1), Vx_ant(1), Vy_ant(1), Vp(1), Ki(1), Ki_ant(1), fi(1) As Double
            Dim i, j, n, ecount As Integer
            Dim d1, d2 As Date, dt As TimeSpan
            Dim L, V, T, Pf As Double

            d1 = Date.Now

            n = UBound(Vz)

            PP = PP
            Sf = S
            Pf = P

            ReDim Vn(n), Vx(n), Vy(n), Vx_ant(n), Vy_ant(n), Vp(n), Ki(n), fi(n)

            Vn = PP.RET_VNAMES()
            fi = Vz.Clone

            Dim maxitINT As Integer = CInt(PP.Parameters("PP_PSFMII"))
            Dim maxitEXT As Integer = CInt(PP.Parameters("PP_PSFMEI"))
            Dim tolINT As Double = CDbl(PP.Parameters("PP_PSFILT"))
            Dim tolEXT As Double = CDbl(PP.Parameters("PP_PSFELT"))

            Dim Tmin, Tmax, epsilon(4) As Double

            Tmax = 2000.0#
            Tmin = 50.0#

            epsilon(0) = 0.001
            epsilon(1) = 0.01
            epsilon(2) = 0.1
            epsilon(3) = 1
            epsilon(4) = 10

            Dim fx, fx2, dfdx, x1, dx As Double

            Dim cnt As Integer

            If Tref = 0 Then Tref = 298.15

            For j = 0 To 4

                cnt = 0
                x1 = Tref

                Do

                    If My.MyApplication._EnableParallelProcessing Then
                        My.MyApplication.IsRunningParallelTasks = True
                        If My.MyApplication._EnableGPUProcessing Then
                            'My.MyApplication.gpu.EnableMultithreading()
                        End If
                        Try
                            Dim task1 = Task.Factory.StartNew(Sub()
                                                                  fx = Serror("PT", x1, P, Vz, PP)(0)
                                                              End Sub)
                            Dim task2 = Task.Factory.StartNew(Sub()
                                                                  fx2 = Serror("PT", x1 + epsilon(j), P, Vz, PP)(0)
                                                              End Sub)
                            Task.WaitAll(task1, task2)
                        Catch ae As AggregateException
                            Throw ae.Flatten().InnerException
                        Finally
                            'If My.MyApplication._EnableGPUProcessing Then
                            '    My.MyApplication.gpu.DisableMultithreading()
                            '    My.MyApplication.gpu.FreeAll()
                            'End If
                        End Try
                        My.MyApplication.IsRunningParallelTasks = False
                    Else
                        fx = Serror("PT", x1, P, Vz, PP)(0)
                        fx2 = Serror("PT", x1 + epsilon(j), P, Vz, PP)(0)
                    End If

                    If Abs(fx) < tolEXT Then Exit Do

                    dfdx = (fx2 - fx) / epsilon(j)
                    dx = fx / dfdx

                    x1 = x1 - dx

                    cnt += 1

                Loop Until cnt > maxitEXT Or Double.IsNaN(x1)

                T = x1

                If Not Double.IsNaN(T) And Not Double.IsInfinity(T) And Not cnt > maxitEXT Then
                    If T > Tmin And T < Tmax Then Exit For
                End If

            Next

            If Double.IsNaN(T) Or T <= Tmin Or T >= Tmax Then Throw New Exception("PS Flash [NL]: Invalid result: Temperature did not converge.")

            Dim tmp As Object = Flash_PT(Vz, P, T, PP)

            L = tmp(0)
            V = tmp(1)
            Vx = tmp(2)
            Vy = tmp(3)
            ecount = tmp(4)

            For i = 0 To n
                Ki(i) = Vy(i) / Vx(i)
            Next

            d2 = Date.Now

            dt = d2 - d1

            WriteDebugInfo("PS Flash [NL]: Converged in " & ecount & " iterations. Time taken: " & dt.TotalMilliseconds & " ms.")

            Return New Object() {L, V, Vx, Vy, T, ecount, Ki, 0.0#, PP.RET_NullVector, 0.0#, PP.RET_NullVector}

        End Function


        Public Function Flash_PS_2(ByVal Vz As Double(), ByVal P As Double, ByVal S As Double, ByVal Tref As Double, ByVal PP As PropertyPackages.PropertyPackage, Optional ByVal ReuseKI As Boolean = False, Optional ByVal PrevKi As Double() = Nothing) As Object

            Dim doparallel As Boolean = My.MyApplication._EnableParallelProcessing

            Dim Vn(1) As String, Vx(1), Vy(1), Vx_ant(1), Vy_ant(1), Vp(1), Ki(1), Ki_ant(1), fi(1) As Double
            Dim i, n, ecount As Integer
            Dim d1, d2 As Date, dt As TimeSpan
            Dim L, V, T, Pf As Double

            Dim resultFlash As Object
            Dim Tb, Td, Sb, Sd As Double
            Dim ErrRes As Object

            d1 = Date.Now

            n = UBound(Vz)

            Sf = S
            Pf = P

            ReDim Vn(n), Vy(n), Vp(n), Ki(n), fi(n)

            Vn = PP.RET_VNAMES()
            fi = Vz.Clone

            Dim maxitINT As Integer = CInt(PP.Parameters("PP_PSFMII"))
            Dim maxitEXT As Integer = CInt(PP.Parameters("PP_PSFMEI"))
            Dim tolINT As Double = CDbl(PP.Parameters("PP_PSFILT"))
            Dim tolEXT As Double = CDbl(PP.Parameters("PP_PSFELT"))

            Dim Tmin, Tmax As Double

            Tmax = 2000.0#
            Tmin = 50.0#

            If Tref = 0.0# Then Tref = 298.15

            'calculate dew point and boiling point

            Dim alreadymt As Boolean = False

            If My.MyApplication._EnableParallelProcessing Then
                My.MyApplication.IsRunningParallelTasks = True
                If My.MyApplication._EnableGPUProcessing Then
                    'If Not My.MyApplication.gpu.IsMultithreadingEnabled Then
                    '    My.MyApplication.gpu.EnableMultithreading()
                    'Else
                    '    alreadymt = True
                    'End If
                End If
                Try
                    Dim task1 = Task.Factory.StartNew(Sub()
                                                          Dim ErrRes1 = Serror("PV", 0, P, Vz, PP)
                                                          Sb = ErrRes1(0)
                                                          Tb = ErrRes1(1)
                                                      End Sub)
                    Dim task2 = Task.Factory.StartNew(Sub()
                                                          Dim ErrRes2 = Serror("PV", 1, P, Vz, PP)
                                                          Sd = ErrRes2(0)
                                                          Td = ErrRes2(1)
                                                      End Sub)
                    Task.WaitAll(task1, task2)
                Catch ae As AggregateException
                    Throw ae.Flatten().InnerException
                Finally
                    'If My.MyApplication._EnableGPUProcessing Then
                    '    If Not alreadymt Then
                    '        My.MyApplication.gpu.DisableMultithreading()
                    '        My.MyApplication.gpu.FreeAll()
                    '    End If
                    'End If
                End Try
                My.MyApplication.IsRunningParallelTasks = False
            Else
                ErrRes = Serror("PV", 0, P, Vz, PP)
                Sb = ErrRes(0)
                Tb = ErrRes(1)
                ErrRes = Serror("PV", 1, P, Vz, PP)
                Sd = ErrRes(0)
                Td = ErrRes(1)
            End If

            Dim S1, S2, T1, T2, V1, V2 As Double

            If Sb > 0 And Sd < 0 Then

                'specified entropy requires partial evaporation 
                'calculate vapour fraction

                ecount = 0
                V = 0
                S1 = Sb
                Do
                    ecount += 1
                    V1 = V
                    If V1 < 1 Then
                        V2 = V1 + 0.01
                    Else
                        V2 = V1 - 0.01
                    End If

                    S2 = Serror("PV", V2, P, Vz, PP)(0)
                    V = V1 + (V2 - V1) * (0 - S1) / (S2 - S1)
                    If V < 0 Then V = 0
                    If V > 1 Then V = 1
                    resultFlash = Serror("PV", V, P, Vz, PP)
                    S1 = resultFlash(0)
                Loop Until Abs(S1) < itol Or ecount > maxitEXT

                T = resultFlash(1)
                L = resultFlash(3)
                Vy = resultFlash(4)
                Vx = resultFlash(5)
                For i = 0 To n
                    Ki(i) = Vy(i) / Vx(i)
                Next

            ElseIf Sd > 0 Then

                'only gas phase
                'calculate temperature

                ecount = 0
                T = Td
                S1 = Sd
                Do
                    ecount += 1
                    T1 = T
                    T2 = T1 + 1
                    S2 = Sf - PP.DW_CalcEntropy(Vz, T2, P, State.Vapor)
                    T = T1 + (T2 - T1) * (0 - S1) / (S2 - S1)
                    S1 = Sf - PP.DW_CalcEntropy(Vz, T, P, State.Vapor)
                Loop Until Abs(S1) < itol Or ecount > maxitEXT

                L = 0
                V = 1
                Vy = Vz.Clone
                Vx = Vz.Clone
                L = 0
                For i = 0 To n
                    Ki(i) = 1
                Next

                If T <= Tmin Or T >= Tmax Then Throw New Exception("PS Flash [NL]: Invalid result: Temperature did not converge.")

            Else

                'specified enthalpy requires pure liquid 
                'calculate temperature

                ecount = 0
                T = Tb
                S1 = Sb
                Do
                    ecount += 1
                    T1 = T
                    T2 = T1 - 1
                    S2 = Serror("PT", T2, P, Vz, PP)(0)
                    T = T1 + (T2 - T1) * (0 - S1) / (S2 - S1)
                    resultFlash = Serror("PT", T, P, Vz, PP)
                    S1 = resultFlash(0)
                Loop Until Abs(S1) < itol Or ecount > maxitEXT

                V = 0
                L = resultFlash(3)
                Vy = resultFlash(4)
                Vx = resultFlash(5)

                For i = 0 To n
                    Ki(i) = Vy(i) / Vx(i)
                Next

                If T <= Tmin Or T >= Tmax Then Throw New Exception("PS Flash [NL]: Invalid result: Temperature did not converge.")
            End If

            d2 = Date.Now

            dt = d2 - d1

            WriteDebugInfo("PS Flash [NL]: Converged in " & ecount & " iterations. Time taken: " & dt.TotalMilliseconds & " ms.")

            Return New Object() {L, V, Vx, Vy, T, ecount, Ki, 0.0#, PP.RET_NullVector, 0.0#, PP.RET_NullVector}

        End Function

        Public Overrides Function Flash_TV(ByVal Vz As Double(), ByVal T As Double, ByVal V As Double, ByVal Pref As Double, ByVal PP As PropertyPackages.PropertyPackage, Optional ByVal ReuseKI As Boolean = False, Optional ByVal PrevKi As Double() = Nothing) As Object

            Dim Vn(1) As String, Vx(1), Vy(1), Vx_ant(1), Vy_ant(1), Vp(1), Ki(1), Ki_ant(1), fi(1) As Double
            Dim i, n, ecount As Integer
            Dim d1, d2 As Date, dt As TimeSpan
            Dim Pmin, Pmax, soma_x, soma_y As Double
            Dim L, Lf, Vf, P, Pf, deltaP As Double

            d1 = Date.Now

            etol = CDbl(PP.Parameters("PP_PTFELT"))
            maxit_e = CInt(PP.Parameters("PP_PTFMEI"))
            itol = CDbl(PP.Parameters("PP_PTFILT"))
            maxit_i = CInt(PP.Parameters("PP_PTFMII"))

            n = UBound(Vz)

            PP = PP
            Vf = V
            L = 1 - V
            Lf = 1 - Vf
            Pf = P

            ReDim Vn(n), Vx(n), Vy(n), Vx_ant(n), Vy_ant(n), Vp(n), Ki(n), fi(n)
            Dim dFdP As Double

            Dim VTc = PP.RET_VTC()

            Vn = PP.RET_VNAMES()
            fi = Vz.Clone

            If Pref = 0.0# Then

                i = 0
                Do
                    Vp(i) = PP.AUX_PVAPi(Vn(i), T)
                    i += 1
                Loop Until i = n + 1

                Pmin = Common.Min(Vp)
                Pmax = Common.Max(Vp)

                Pref = Pmin + (1 - V) * (Pmax - Pmin)

            Else

                Pmin = Pref * 0.8
                Pmax = Pref * 1.2

            End If

            P = Pref

            'Calculate Ki`s

            If Not ReuseKI Then
                i = 0
                Do
                    Vp(i) = PP.AUX_PVAPi(Vn(i), T)
                    Ki(i) = Vp(i) / P
                    i += 1
                Loop Until i = n + 1
            Else
                If Not PP.AUX_CheckTrivial(PrevKi) Then
                    For i = 0 To n
                        Vp(i) = PP.AUX_PVAPi(Vn(i), T)
                        Ki(i) = PrevKi(i)
                    Next
                Else
                    i = 0
                    Do
                        Vp(i) = PP.AUX_PVAPi(Vn(i), T)
                        Ki(i) = Vp(i) / P
                        i += 1
                    Loop Until i = n + 1
                End If
            End If

            i = 0
            Do
                If Vz(i) <> 0 Then
                    Vy(i) = Vz(i) * Ki(i) / ((Ki(i) - 1) * V + 1)
                    Vx(i) = Vy(i) / Ki(i)
                Else
                    Vy(i) = 0
                    Vx(i) = 0
                End If
                i += 1
            Loop Until i = n + 1

            i = 0
            soma_x = 0
            soma_y = 0
            Do
                soma_x = soma_x + Vx(i)
                soma_y = soma_y + Vy(i)
                i = i + 1
            Loop Until i = n + 1
            i = 0
            Do
                Vx(i) = Vx(i) / soma_x
                Vy(i) = Vy(i) / soma_y
                i = i + 1
            Loop Until i = n + 1

            If PP.AUX_IS_SINGLECOMP(Vz) Then
                WriteDebugInfo("TV Flash [NL]: Converged in 1 iteration.")
                P = 0
                For i = 0 To n
                    P += Vz(i) * PP.AUX_PVAPi(i, T)
                Next
                Return New Object() {L, V, Vx, Vy, P, 0, Ki, 0.0#, PP.RET_NullVector, 0.0#, PP.RET_NullVector}
            End If

            Dim marcador3, marcador2, marcador As Integer
            Dim stmp4_ant, stmp4, Pant, fval As Double
            Dim chk As Boolean = False

            If V = 1.0# Or V = 0.0# Then

                ecount = 0
                Do

                    marcador3 = 0

                    Dim cont_int = 0
                    Do


                        Ki = PP.DW_CalcKvalue(Vx, Vy, T, P)

                        marcador = 0
                        If stmp4_ant <> 0 Then
                            marcador = 1
                        End If
                        stmp4_ant = stmp4

                        If V = 0 Then
                            i = 0
                            stmp4 = 0
                            Do
                                stmp4 = stmp4 + Ki(i) * Vx(i)
                                i = i + 1
                            Loop Until i = n + 1
                        Else
                            i = 0
                            stmp4 = 0
                            Do
                                stmp4 = stmp4 + Vy(i) / Ki(i)
                                i = i + 1
                            Loop Until i = n + 1
                        End If

                        If V = 0 Then
                            i = 0
                            Do
                                Vy_ant(i) = Vy(i)
                                Vy(i) = Ki(i) * Vx(i) / stmp4
                                i = i + 1
                            Loop Until i = n + 1
                        Else
                            i = 0
                            Do
                                Vx_ant(i) = Vx(i)
                                Vx(i) = (Vy(i) / Ki(i)) / stmp4
                                i = i + 1
                            Loop Until i = n + 1
                        End If

                        marcador2 = 0
                        If marcador = 1 Then
                            If V = 0 Then
                                If Math.Abs(Vy(0) - Vy_ant(0)) < itol Then
                                    marcador2 = 1
                                End If
                            Else
                                If Math.Abs(Vx(0) - Vx_ant(0)) < itol Then
                                    marcador2 = 1
                                End If
                            End If
                        End If

                        cont_int = cont_int + 1

                    Loop Until marcador2 = 1 Or Double.IsNaN(stmp4) Or cont_int > maxit_i

                    Dim K1(n), K2(n), dKdP(n) As Double

                    K1 = PP.DW_CalcKvalue(Vx, Vy, T, P)
                    K2 = PP.DW_CalcKvalue(Vx, Vy, T, P * 1.001)

                    For i = 0 To n
                        dKdP(i) = (K2(i) - K1(i)) / (0.001 * P)
                    Next

                    fval = stmp4 - 1

                    ecount += 1

                    i = 0
                    dFdP = 0
                    Do
                        If V = 0 Then
                            dFdP = dFdP + Vx(i) * dKdP(i)
                        Else
                            dFdP = dFdP - Vy(i) / (Ki(i) ^ 2) * dKdP(i)
                        End If
                        i = i + 1
                    Loop Until i = n + 1

                    If (P - fval / dFdP) < 0 Then

                        P = (P + Pant) / 2

                    Else

                        Pant = P

                        deltaP = -fval / dFdP

                        If Abs(deltaP) > 0.1 * P Then
                            P = P + Sign(deltaP) * 0.1 * P
                        Else
                            P = P + deltaP
                        End If

                    End If

                    WriteDebugInfo("TV Flash [NL]: Iteration #" & ecount & ", P = " & P & ", VF = " & V)

                Loop Until Math.Abs(fval) < etol Or Double.IsNaN(P) = True Or ecount > maxit_e

            Else

                ecount = 0

                Do

                    Ki = PP.DW_CalcKvalue(Vx, Vy, T, P)

                    i = 0
                    Do
                        If Vz(i) <> 0 Then
                            Vy_ant(i) = Vy(i)
                            Vx_ant(i) = Vx(i)
                            Vy(i) = Vz(i) * Ki(i) / ((Ki(i) - 1) * V + 1)
                            Vx(i) = Vy(i) / Ki(i)
                        Else
                            Vy(i) = 0
                            Vx(i) = 0
                        End If
                        i += 1
                    Loop Until i = n + 1
                    i = 0
                    soma_x = 0
                    soma_y = 0
                    Do
                        soma_x = soma_x + Vx(i)
                        soma_y = soma_y + Vy(i)
                        i = i + 1
                    Loop Until i = n + 1
                    i = 0
                    Do
                        Vx(i) = Vx(i) / soma_x
                        Vy(i) = Vy(i) / soma_y
                        i = i + 1
                    Loop Until i = n + 1

                    If V <= 0.5 Then

                        i = 0
                        stmp4 = 0
                        Do
                            stmp4 = stmp4 + Ki(i) * Vx(i)
                            i = i + 1
                        Loop Until i = n + 1

                        Dim K1(n), K2(n), dKdP(n) As Double

                        K1 = PP.DW_CalcKvalue(Vx, Vy, T, P)
                        K2 = PP.DW_CalcKvalue(Vx, Vy, T, P * 1.001)

                        For i = 0 To n
                            dKdP(i) = (K2(i) - K1(i)) / (0.001 * P)
                        Next

                        i = 0
                        dFdP = 0
                        Do
                            dFdP = dFdP + Vx(i) * dKdP(i)
                            i = i + 1
                        Loop Until i = n + 1

                    Else

                        i = 0
                        stmp4 = 0
                        Do
                            stmp4 = stmp4 + Vy(i) / Ki(i)
                            i = i + 1
                        Loop Until i = n + 1

                        Dim K1(n), K2(n), dKdP(n) As Double

                        K1 = PP.DW_CalcKvalue(Vx, Vy, T, P)
                        K2 = PP.DW_CalcKvalue(Vx, Vy, T, P * 1.001)

                        For i = 0 To n
                            dKdP(i) = (K2(i) - K1(i)) / (0.001 * P)
                        Next

                        i = 0
                        dFdP = 0
                        Do
                            dFdP = dFdP - Vy(i) / (Ki(i) ^ 2) * dKdP(i)
                            i = i + 1
                        Loop Until i = n + 1
                    End If

                    ecount += 1

                    fval = stmp4 - 1

                    If (P - fval / dFdP) < 0 Then

                        P = (P + Pant) / 2

                    Else

                        Pant = P

                        deltaP = -fval / dFdP

                        If Abs(deltaP) > 0.1 * P Then
                            P = P + Sign(deltaP) * 0.1 * P
                        Else
                            P = P + deltaP
                        End If

                    End If

                    WriteDebugInfo("TV Flash [NL]: Iteration #" & ecount & ", P = " & P & ", VF = " & V)

                Loop Until Math.Abs(fval) < etol Or Double.IsNaN(P) = True Or ecount > maxit_e

            End If

            d2 = Date.Now

            dt = d2 - d1

            If ecount > maxit_e Then Throw New Exception(DTL.App.GetLocalString("PropPack_FlashMaxIt2"))

            If PP.AUX_CheckTrivial(Ki) Then Throw New Exception("TV Flash [NL]: Invalid result: converged to the trivial solution (P = " & P & " ).")

            WriteDebugInfo("TV Flash [NL]: Converged in " & ecount & " iterations. Time taken: " & dt.TotalMilliseconds & " ms.")

            Return New Object() {L, V, Vx, Vy, P, ecount, Ki, 0.0#, PP.RET_NullVector, 0.0#, PP.RET_NullVector}

        End Function

        Public Overrides Function Flash_PV(ByVal Vz As Double(), ByVal P As Double, ByVal V As Double, ByVal Tref As Double, ByVal PP As PropertyPackages.PropertyPackage, Optional ByVal ReuseKI As Boolean = False, Optional ByVal PrevKi As Double() = Nothing) As Object

            Dim i, n, ecount As Integer
            Dim d1, d2 As Date, dt As TimeSpan
            Dim L, Lf, Vf, T, Tf, deltaT As Double
            Dim e1 As Double
            Dim AF As Double = 1

            d1 = Date.Now

            etol = CDbl(PP.Parameters("PP_PTFELT"))
            maxit_e = CInt(PP.Parameters("PP_PTFMEI"))
            itol = CDbl(PP.Parameters("PP_PTFILT"))
            maxit_i = CInt(PP.Parameters("PP_PTFMII"))

            n = UBound(Vz)

            PP = PP
            Vf = V
            L = 1 - V
            Lf = 1 - Vf
            Tf = T

            Dim Vn(n) As String, Vx(n), Vy(n), Vx_ant(n), Vy_ant(n), Vp(n), Ki(n), fi(n) As Double
            Dim Vt(n), VTc(n), Tmin, Tmax, dFdT, Tsat(n) As Double

            Vn = PP.RET_VNAMES()
            VTc = PP.RET_VTC()
            fi = Vz.Clone

            Tmin = 0.0#
            Tmax = 0.0#

            If Tref = 0.0# Then
                i = 0
                Tref = 0.0#
                Do
                    Tref += 0.8 * Vz(i) * VTc(i)
                    Tmin += 0.1 * Vz(i) * VTc(i)
                    Tmax += 2.0 * Vz(i) * VTc(i)
                    i += 1
                Loop Until i = n + 1
            Else
                Tmin = Tref - 50
                Tmax = Tref + 50
            End If

            T = Tref

            'Calculate Ki`s

            If Not ReuseKI Then
                i = 0
                Do
                    Vp(i) = PP.AUX_PVAPi(Vn(i), T)
                    Ki(i) = Vp(i) / P
                    i += 1
                Loop Until i = n + 1
            Else
                If Not PP.AUX_CheckTrivial(PrevKi) And Not Double.IsNaN(PrevKi(0)) Then
                    For i = 0 To n
                        Vp(i) = PP.AUX_PVAPi(Vn(i), T)
                        Ki(i) = PrevKi(i)
                    Next
                Else
                    i = 0
                    Do
                        Vp(i) = PP.AUX_PVAPi(Vn(i), T)
                        Ki(i) = Vp(i) / P
                        i += 1
                    Loop Until i = n + 1
                End If
            End If

            i = 0
            Do
                If Vz(i) <> 0 Then
                    Vy(i) = Vz(i) * Ki(i) / ((Ki(i) - 1) * V + 1)
                    If Double.IsInfinity(Vy(i)) Then Vy(i) = 0.0#
                    Vx(i) = Vy(i) / Ki(i)
                Else
                    Vy(i) = 0
                    Vx(i) = 0
                End If
                i += 1
            Loop Until i = n + 1

            Vx = Vx.NormalizeY()
            Vy = Vy.NormalizeY()

            If PP.AUX_IS_SINGLECOMP(Vz) Then
                WriteDebugInfo("PV Flash [NL]: Converged in 1 iteration.")
                T = 0
                For i = 0 To n
                    T += Vz(i) * PP.AUX_TSATi(P, i)
                Next
                Return New Object() {L, V, Vx, Vy, T, 0, Ki, 0.0#, PP.RET_NullVector, 0.0#, PP.RET_NullVector}
            End If

            Dim marcador3, marcador2, marcador As Integer
            Dim stmp4_ant, stmp4, Tant, fval As Double
            Dim chk As Boolean = False

            If V = 1.0# Or V = 0.0# Then

                ecount = 0
                Do

                    marcador3 = 0

                    Dim cont_int = 0
                    Do

                        Ki = PP.DW_CalcKvalue(Vx, Vy, T, P)

                        marcador = 0
                        If stmp4_ant <> 0 Then
                            marcador = 1
                        End If
                        stmp4_ant = stmp4

                        If V = 0 Then
                            stmp4 = Ki.MultiplyY(Vx).SumY
                            'i = 0
                            'stmp4 = 0
                            'Do
                            '    stmp4 = stmp4 + Ki(i) * Vx(i)
                            '    i = i + 1
                            'Loop Until i = n + 1
                        Else
                            stmp4 = Vy.DivideY(Ki).SumY
                            'i = 0
                            'stmp4 = 0
                            'Do
                            '    stmp4 = stmp4 + Vy(i) / Ki(i)
                            '    i = i + 1
                            'Loop Until i = n + 1
                        End If

                        If V = 0 Then
                            Vy_ant = Vy.Clone
                            Vy = Ki.MultiplyY(Vx).MultiplyConstY(1 / stmp4)
                            'i = 0
                            'Do
                            '    Vy_ant(i) = Vy(i)
                            '    Vy(i) = Ki(i) * Vx(i) / stmp4
                            '    i = i + 1
                            'Loop Until i = n + 1
                        Else
                            Vx_ant = Vx.Clone
                            Vx = Vy.DivideY(Ki).MultiplyConstY(1 / stmp4)
                            'i = 0
                            'Do
                            '    Vx_ant(i) = Vx(i)
                            '    Vx(i) = (Vy(i) / Ki(i)) / stmp4
                            '    i = i + 1
                            'Loop Until i = n + 1
                        End If

                        marcador2 = 0
                        If marcador = 1 Then
                            If V = 0 Then
                                If Math.Abs(Vy(0) - Vy_ant(0)) < itol Then
                                    marcador2 = 1
                                End If
                            Else
                                If Math.Abs(Vx(0) - Vx_ant(0)) < itol Then
                                    marcador2 = 1
                                End If
                            End If
                        End If

                        cont_int = cont_int + 1

                    Loop Until marcador2 = 1 Or Double.IsNaN(stmp4) Or cont_int > maxit_i

                    Dim K1(n), K2(n), dKdT(n) As Double

                    K1 = PP.DW_CalcKvalue(Vx, Vy, T, P)
                    K2 = PP.DW_CalcKvalue(Vx, Vy, T + 0.01, P)

                    dKdT = K2.SubtractY(K1).MultiplyConstY(1 / 0.01)
                    'For i = 0 To n
                    '    dKdT(i) = (K2(i) - K1(i)) / 0.01
                    'Next

                    fval = stmp4 - 1

                    ecount += 1

                    i = 0
                    dFdT = 0
                    Do
                        If V = 0 Then
                            dFdT = Vx.MultiplyY(dKdT).SumY
                            'dFdT = dFdT + Vx(i) * dKdT(i)
                        Else
                            dFdT = -Vy.DivideY(Ki).DivideY(Ki).MultiplyY(dKdT).SumY
                            'dFdT = dFdT - Vy(i) / (Ki(i) ^ 2) * dKdT(i)
                        End If
                        i = i + 1
                    Loop Until i = n + 1

                    Tant = T
                    deltaT = -fval / dFdT

                    If Abs(deltaT) > 0.1 * T Then
                        T = T + Sign(deltaT) * 0.1 * T
                    Else
                        T = T + deltaT
                    End If

                    WriteDebugInfo("PV Flash [NL]: Iteration #" & ecount & ", T = " & T & ", VF = " & V)

                Loop Until Math.Abs(fval) < etol Or Double.IsNaN(T) = True Or ecount > maxit_e

            Else

                ecount = 0

                Do

                    Ki = PP.DW_CalcKvalue(Vx, Vy, T, P)

                    i = 0
                    Do
                        If Vz(i) <> 0 Then
                            Vy_ant(i) = Vy(i)
                            Vx_ant(i) = Vx(i)
                            Vy(i) = Vz(i) * Ki(i) / ((Ki(i) - 1) * V + 1)
                            Vx(i) = Vy(i) / Ki(i)
                        Else
                            Vy(i) = 0
                            Vx(i) = 0
                        End If
                        i += 1
                    Loop Until i = n + 1

                    Vx = Vx.NormalizeY()
                    Vy = Vy.NormalizeY()

                    If V <= 0.5 Then

                        stmp4 = Ki.MultiplyY(Vx).SumY
                        'i = 0
                        'stmp4 = 0
                        'Do
                        '    stmp4 = stmp4 + Ki(i) * Vx(i)
                        '    i = i + 1
                        'Loop Until i = n + 1

                        Dim K1(n), K2(n), dKdT(n) As Double

                        K1 = PP.DW_CalcKvalue(Vx, Vy, T, P)
                        K2 = PP.DW_CalcKvalue(Vx, Vy, T + 0.1, P)

                        dKdT = K2.SubtractY(K1).MultiplyConstY(1 / 0.1)
                        'For i = 0 To n
                        '    dKdT(i) = (K2(i) - K1(i)) / (0.1)
                        'Next

                        dFdT = Vx.MultiplyY(dKdT).SumY
                        'i = 0
                        'dFdT = 0
                        'Do
                        '    dFdT = dFdT + Vx(i) * dKdT(i)
                        '    i = i + 1
                        'Loop Until i = n + 1

                    Else

                        stmp4 = Vy.DivideY(Ki).SumY
                        'i = 0
                        'stmp4 = 0
                        'Do
                        '    stmp4 = stmp4 + Vy(i) / Ki(i)
                        '    i = i + 1
                        'Loop Until i = n + 1

                        Dim K1(n), K2(n), dKdT(n) As Double

                        K1 = PP.DW_CalcKvalue(Vx, Vy, T, P)
                        K2 = PP.DW_CalcKvalue(Vx, Vy, T + 1, P)

                        dKdT = K2.SubtractY(K1)
                        'For i = 0 To n
                        '    dKdT(i) = (K2(i) - K1(i)) / (1)
                        'Next

                        dFdT = -Vy.DivideY(Ki).DivideY(Ki).MultiplyY(dKdT).SumY
                        'i = 0
                        'dFdT = 0
                        'Do
                        '    dFdT = dFdT - Vy(i) / (Ki(i) ^ 2) * dKdT(i)
                        '    i = i + 1
                        'Loop Until i = n + 1

                    End If

                    ecount += 1

                    fval = stmp4 - 1

                    Tant = T
                    deltaT = -fval / dFdT * AF
                    AF *= 1.01
                    If Abs(deltaT) > 0.1 * T Then
                        T = T + Sign(deltaT) * 0.1 * T
                    Else
                        T = T + deltaT
                    End If

                    e1 = Vx.SubtractY(Vx_ant).AbsSumY + Vy.SubtractY(Vy_ant).AbsSumY + Math.Abs(T - Tant)

                    WriteDebugInfo("PV Flash [NL]: Iteration #" & ecount & ", T = " & T & ", VF = " & V)

                Loop Until (Math.Abs(fval) < etol And e1 < etol) Or Double.IsNaN(T) = True Or ecount > maxit_e

            End If

            d2 = Date.Now

            dt = d2 - d1

            If ecount > maxit_e Then Throw New Exception(DTL.App.GetLocalString("PropPack_FlashMaxIt2"))

            If PP.AUX_CheckTrivial(Ki) Then Throw New Exception("PV Flash [NL]: Invalid result: converged to the trivial solution (T = " & T & " ).")

            WriteDebugInfo("PV Flash [NL]: Converged in " & ecount & " iterations. Time taken: " & dt.TotalMilliseconds & " ms.")

            Return New Object() {L, V, Vx, Vy, T, ecount, Ki, 0.0#, PP.RET_NullVector, 0.0#, PP.RET_NullVector}

        End Function

        Function OBJ_FUNC_PH_FLASH(ByVal Type As String, ByVal X As Double, ByVal P As Double, ByVal Vz() As Double, ByVal PP As PropertyPackages.PropertyPackage) As Object

            Dim n = UBound(Vz)
            Dim L, V, Vx(), Vy(), _Hl, _Hv, T As Double

            If Type = "PT" Then
                Dim tmp = Me.Flash_PT(Vz, P, X, PP)
                L = tmp(0)
                V = tmp(1)
                Vx = tmp(2)
                Vy = tmp(3)
                T = X
            Else
                Dim tmp = Me.Flash_PV(Vz, P, X, 0.0#, PP)
                L = tmp(0)
                V = tmp(1)
                Vx = tmp(2)
                Vy = tmp(3)
                T = tmp(4)
            End If

            _Hv = 0.0#
            _Hl = 0.0#

            Dim mmg, mml As Double
            If V > 0 Then _Hv = PP.DW_CalcEnthalpy(Vy, T, P, State.Vapor)
            If L > 0 Then _Hl = PP.DW_CalcEnthalpy(Vx, T, P, State.Liquid)
            mmg = PP.AUX_MMM(Vy)
            mml = PP.AUX_MMM(Vx)

            Dim herr As Double = Hf - (mmg * V / (mmg * V + mml * L)) * _Hv - (mml * L / (mmg * V + mml * L)) * _Hl
            OBJ_FUNC_PH_FLASH = {herr, T, V, L, Vy, Vx}

            WriteDebugInfo("PH Flash [NL]: Current T = " & T & ", Current H Error = " & herr)

        End Function

        Function OBJ_FUNC_PS_FLASH(ByVal Type As String, ByVal X As Double, ByVal P As Double, ByVal Vz() As Double, ByVal PP As PropertyPackages.PropertyPackage) As Object

            Dim n = UBound(Vz)
            Dim L, V, Vx(), Vy(), _Sl, _Sv, T As Double

            If Type = "PT" Then
                Dim tmp = Me.Flash_PT(Vz, P, X, PP)
                L = tmp(0)
                V = tmp(1)
                Vx = tmp(2)
                Vy = tmp(3)
                T = X
            Else
                Dim tmp = Me.Flash_PV(Vz, P, X, 0.0#, PP)
                L = tmp(0)
                V = tmp(1)
                Vx = tmp(2)
                Vy = tmp(3)
                T = tmp(4)
            End If

            _Sv = 0.0#
            _Sl = 0.0#
            Dim mmg, mml As Double

            If V > 0 Then _Sv = PP.DW_CalcEntropy(Vy, T, P, State.Vapor)
            If L > 0 Then _Sl = PP.DW_CalcEntropy(Vx, T, P, State.Liquid)
            mmg = PP.AUX_MMM(Vy)
            mml = PP.AUX_MMM(Vx)

            Dim serr As Double = Sf - (mmg * V / (mmg * V + mml * L)) * _Sv - (mml * L / (mmg * V + mml * L)) * _Sl
            OBJ_FUNC_PS_FLASH = {serr, T, V, L, Vy, Vx}

            WriteDebugInfo("PS Flash [NL]: Current T = " & T & ", Current S Error = " & serr)

        End Function

        Function Herror(ByVal type As String, ByVal X As Double, ByVal P As Double, ByVal Vz() As Double, ByVal PP As PropertyPackages.PropertyPackage) As Object
            Return OBJ_FUNC_PH_FLASH(type, X, P, Vz, PP)
        End Function

        Function Serror(ByVal type As String, ByVal X As Double, ByVal P As Double, ByVal Vz() As Double, ByVal PP As PropertyPackages.PropertyPackage) As Object
            Return OBJ_FUNC_PS_FLASH(type, X, P, Vz, PP)
        End Function

    End Class

End Namespace