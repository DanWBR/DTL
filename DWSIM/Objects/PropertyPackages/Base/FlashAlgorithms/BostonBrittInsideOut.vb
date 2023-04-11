'    Boston-Britt Inside-Out Flash Algorithms
'    Copyright 2010-2014 Daniel Wagner O. de Medeiros
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
Imports DTL.DTL.MathEx.Common

Namespace DTL.SimulationObjects.PropertyPackages.Auxiliary.FlashAlgorithms

    <Serializable()> Public Class BostonBrittInsideOut

        Inherits FlashAlgorithm

        Dim n, ecount As Integer
        Dim etol As Double = 0.000001
        Dim itol As Double = 0.000001
        Dim maxit_i As Integer = 100
        Dim maxit_e As Integer = 100
        Dim Vn(n) As String
        Dim Vx(n), Vy(n), Vp(n), ui(n), uic(n), pi(n), Ki(n), fi(n), Vt(n), Vpc(n), VTc(n), Vw(n) As Double
        Dim L, Lf, V, Vf, R, Rant, Tant, Pant, T, T_, Tf, P, P_, Pf, T0, P0, A, B, C, D, E, F, Ac, Bc, Cc, Dc, Ec, Fc As Double
        Dim Kb, Kb0, Kb_ As Double
        Dim DHv, DHl, DHv1, DHv2, DHl1, DHl2, Hv0, Hvid, Hlid, Hf, DHlsp, DHvsp As Double
        Dim DSv, DSl, DSv1, DSv2, DSl1, DSl2, Sv0, Svid, Slid, Sf, DSlsp, DSvsp As Double
        Dim Pb, Pd, Pmin, Pmax, Px, sum_x, sum_y, Tmin, Tmax As Double
        Dim proppack As PropertyPackage
        Dim tmpdx, refx, currx As Object

        Public Overrides Function Flash_PT(ByVal Vz As Double(), ByVal P As Double, ByVal T As Double, ByVal PP As PropertyPackage, Optional ByVal ReuseKI As Boolean = False, Optional ByVal PrevKi As Double() = Nothing) As Object

            Dim d1, d2 As Date, dt As TimeSpan
            Dim i, j As Integer

            d1 = Date.Now

            Me.etol = PP.Parameters("PP_PTFELT")
            Me.maxit_e = CInt(PP.Parameters("PP_PTFMEI"))
            Me.itol = PP.Parameters("PP_PTFILT")
            Me.maxit_i = CInt(PP.Parameters("PP_PTFMII"))

            n = Vz.Length - 1

            proppack = PP

            ReDim Vn(n), Vx(n), Vy(n), Vp(n), ui(n), uic(n), pi(n), Ki(n), fi(n)

            Vn = PP.RET_VNAMES()
            fi = Vz.Clone

            Dim VPc(n), VTc(n), Vw(n), g As Double

            VPc = PP.RET_VPC()
            VTc = PP.RET_VTC()
            Vw = PP.RET_VW()

            '--------------------------------------
            ' STEP 1 - Assume u, A, B, C, D, E, F 
            '--------------------------------------


            '----------------------------------------
            ' STEP 1.1 - Estimate K, Vx, Vy, V and L 
            '----------------------------------------

            'Calculate Ki`s

            If Not ReuseKI Then
                i = 0
                Do
                    'Vp(i) = PP.AUX_PVAPi(Vn(i), T)
                    Vp(i) = VPc(i) * Exp(5.37 * (1 + Vw(i)) * (1 - VTc(i) / T))
                    Ki(i) = Vp(i) / P
                    i += 1
                Loop Until i = n + 1
            Else
                For i = 0 To n
                    Vp(i) = PP.AUX_PVAPi(Vn(i), T)
                    Ki(i) = PrevKi(i)
                Next
            End If

            'Estimate V

            If T > Common.Max(proppack.RET_VTC, Vz) Then
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

            Dim Vmin, Vmax As Double
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

            V = Vmin + (Vmax - Vmin) / 4

            L = 1 - V

            If n = 0 Then
                If Vp(0) <= P Then
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
            sum_x = 0
            sum_y = 0
            Do
                sum_x = sum_x + Vx(i)
                sum_y = sum_y + Vy(i)
                i = i + 1
            Loop Until i = n + 1
            i = 0
            Do
                Vx(i) = Vx(i) / sum_x
                Vy(i) = Vy(i) / sum_y
                i = i + 1
            Loop Until i = n + 1

            Kb = 1 'CalcKbj1(Ki)
            Kb0 = 1 'Kb

            For i = 0 To n
                ui(i) = Log(Ki(i))
            Next

            Dim fx(n), x(n), dfdx(n, n), dx(n), xbr(n), fbr(n) As Double
            Dim bo As New BrentOpt.Brent
            Dim bo2 As New BrentOpt.BrentMinimize

            ecount = 0

            R = Kb * V / (Kb * V + Kb0 * L)

            Do

                '--------------------------------------------------------------
                ' STEPS 2, 3, 4, 5, 6, 7 and 8 - Calculate R and Energy Balance
                '--------------------------------------------------------------

                Dim fr As Double
                bo2.DefineFuncDelegate(AddressOf TPErrorFunc)
                Rant = R
                fr = bo2.brentoptimize(0.0#, 1.0#, 0.00000001, R)

                If ecount > 0 Then
                    If Rant = 0.0# And R = 1.0# Then R = 0.0#
                    If Rant = 1.0# And R = 0.0# Then R = 1.0#
                End If

                Me.TPErrorFunc(R)

                If R <= 0 Then
                    R = 0
                    L = 1
                    V = 0
                    For i = 0 To n
                        pi(i) = fi(i) / (1 - R + Kb0 * R * Exp(ui(i)))
                    Next
                    Dim sumpi As Double = 0
                    Dim sumeuipi As Double = 0
                    For i = 0 To n
                        sumpi += pi(i)
                        sumeuipi += Exp(ui(i)) * pi(i)
                    Next
                    For i = 0 To n
                        Vx(i) = pi(i) / sumpi
                        Vy(i) = Exp(ui(i)) * pi(i) / sumeuipi
                    Next
                    Exit Do
                End If

                If R >= 1 Then
                    R = 1
                    L = 0
                    V = 1
                    For i = 0 To n
                        pi(i) = fi(i) / (1 - R + Kb0 * R * Exp(ui(i)))
                    Next
                    Dim sumpi As Double = 0
                    Dim sumeuipi As Double = 0
                    For i = 0 To n
                        sumpi += pi(i)
                        sumeuipi += Exp(ui(i)) * pi(i)
                    Next
                    For i = 0 To n
                        Vx(i) = pi(i) / sumpi
                        Vy(i) = Exp(ui(i)) * pi(i) / sumeuipi
                    Next
                    Exit Do
                End If


                'At this point, we have converged R for the simplified model. Proceed to step 9.

                '----------------------------------------------------------
                ' STEP 9 - Rigorous model Enthalpy and K-values calculation
                '----------------------------------------------------------

                Ki = PP.DW_CalcKvalue(Vx, Vy, T, P)
                'Kb = CalcKbj1(Ki)

                For i = 0 To n
                    uic(i) = Log(Ki(i))
                Next

                '-------------------------------------------
                ' STEP 10 - Update variables using Broyden
                '-------------------------------------------

                For i = 0 To n
                    fx(i) = (ui(i) - uic(i))
                    x(i) = ui(i)
                Next

                If ecount = 0 Then
                    For i = 0 To n
                        For j = 0 To n
                            If i = j Then dfdx(i, j) = 1 Else dfdx(i, j) = 0
                        Next
                    Next
                    Broyden.broydn(n, x, fx, dx, xbr, fbr, dfdx, 0)
                Else
                    Broyden.broydn(n, x, fx, dx, xbr, fbr, dfdx, 1)
                End If

                For i = 0 To n
                    ui(i) = ui(i) + dx(i)
                Next

                ecount += 1

                If Double.IsNaN(V) Then Throw New Exception(App.GetLocalString("PropPack_FlashTPVapFracError"))
                If ecount > maxit_e Then Throw New Exception(App.GetLocalString("PropPack_FlashMaxIt2"))

                WriteDebugInfo("PT Flash [IO]: Iteration #" & ecount & ", VF = " & V)



            Loop Until AbsSum(fx) < etol

            d2 = Date.Now

            dt = d2 - d1

            WriteDebugInfo("PT Flash [IO]: Converged in " & ecount & " iterations. Time taken: " & dt.TotalMilliseconds & " ms. Error function value: " & AbsSum(fx))

out:        Return New Object() {L, V, Vx, Vy, ecount, 0.0#, Vx, 0.0#, PP.RET_NullVector}

        End Function

        Public Overrides Function Flash_PH(ByVal Vz As Double(), ByVal P As Double, ByVal H As Double, ByVal Tref As Double, ByVal PP As PropertyPackage, Optional ByVal ReuseKI As Boolean = False, Optional ByVal PrevKi As Double() = Nothing) As Object

            If PP.Parameters.ContainsKey("PP_ENTH_CP_CALC_METHOD") Then
                If PP.Parameters("PP_ENTH_CP_CALC_METHOD") = 1 Then
                    Throw New Exception("Inside-Out PH Flash doesn't work with 'Ideal' Enthalpy/Cp calculation mode enabled. Please change it to 'Lee-Kesler' or 'Excess', or use another Flash Algorithm.")
                End If
            End If

            Dim d1, d2 As Date, dt As TimeSpan
            Dim i, j As Integer

            d1 = Date.Now

            maxit_i = CInt(PP.Parameters("PP_PHFMII"))
            maxit_e = CInt(PP.Parameters("PP_PHFMEI"))
            Me.itol = PP.Parameters("PP_PHFILT")
            Me.etol = PP.Parameters("PP_PHFELT")

            Me.n = Vz.Length - 1

            proppack = PP
            Hf = H * PP.AUX_MMM(Vz)
            Pf = P

            ReDim Vn(n), Vx(n), Vy(n), Vp(n), ui(n), uic(n), pi(n), Ki(n), fi(n), Vpc(n), VTc(n), Vw(n)

            Vn = PP.RET_VNAMES()
            VTc = PP.RET_VTC
            fi = Vz.Clone

            Tmin = 0
            Tmax = 0
            If Tref = 0.0# Or Double.IsNaN(Tref) Then
                i = 0
                Tref = 0
                Do
                    Tref += 0.8 * Vz(i) * VTc(i)
                    Tmin += 0.1 * Vz(i) * VTc(i)
                    Tmax += 2.0 * Vz(i) * VTc(i)
                    i += 1
                Loop Until i = n + 1
            Else
                Tmin = Tref - 200
                Tmax = Tref + 200
            End If
            If Tmin < 100 Then Tmin = 100

            '--------------------------------------
            ' STEP 1 - Assume u, A, B, C, D, E, F 
            '--------------------------------------

            T = Tref + 1
            T_ = Tref - 1
            T0 = Tref

            '----------------------------------------
            ' STEP 1.1 - Estimate K, Vx, Vy, V and L 
            '----------------------------------------

            'Calculate Ki`s

            Vpc = PP.RET_VPC
            VTc = PP.RET_VTC
            Vw = PP.RET_VW

            If Not ReuseKI Then
                i = 0
                Do
                    Vp(i) = PP.AUX_PVAPi(Vn(i), T)
                    Ki(i) = Vpc(i) / P * Math.Exp(5.373 * (1 + Vw(i)) * (1 - VTc(i) / T))
                    i += 1
                Loop Until i = n + 1
            Else
                For i = 0 To n
                    Vp(i) = PP.AUX_PVAPi(Vn(i), T)
                    Ki(i) = PrevKi(i)
                Next
            End If

            'Estimate V

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

            If P <= Pd Then
                'vapor only
                L = 0
                V = 1
            ElseIf P >= Pb Then
                'liquid only
                L = 1
                V = 0
            Else
                'VLE
                V = 1 - (P - Pd) / (Pb - Pd)
                L = 1 - V
            End If

            If n = 0 Then
                If Vp(0) <= P Then
                    L = 1
                    V = 0
                Else
                    L = 0
                    V = 1
                End If
            End If

            If T > Common.Max(VTc, Vz) Then
                V = 1
                L = 0
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
            sum_x = 0
            sum_y = 0
            Do
                sum_x = sum_x + Vx(i)
                sum_y = sum_y + Vy(i)
                i = i + 1
            Loop Until i = n + 1
            i = 0
            Do
                Vx(i) = Vx(i) / sum_x
                Vy(i) = Vy(i) / sum_y
                i = i + 1
            Loop Until i = n + 1

            Kb_ = CalcKbj1(PP.DW_CalcKvalue(Vx, Vy, T_, P))
            Kb = CalcKbj1(PP.DW_CalcKvalue(Vx, Vy, T, P))
            Kb0 = Kb_

            B = Log(Kb_ / Kb) / (1 / T_ - 1 / T)
            A = Log(Kb) - B * (1 / T - 1 / T_)

            For i = 0 To n
                ui(i) = Log(Ki(i) / Kb)
            Next

            If My.MyApplication._EnableParallelProcessing Then
                My.MyApplication.IsRunningParallelTasks = True
                Try
                    Dim task1 As Task = New Task(Sub()
                                                     DHv1 = PP.DW_CalcEnthalpyDeparture(Vy, T, P, State.Vapor) * PP.AUX_MMM(Vy) / 1000
                                                     DHv2 = PP.DW_CalcEnthalpyDeparture(Vy, Tref, P, State.Vapor) * PP.AUX_MMM(Vy) / 1000
                                                     C = DHv2
                                                     D = (DHv1 - C) / (T - Tref)
                                                 End Sub)
                    Dim task2 As Task = New Task(Sub()
                                                     If T < Common.Max(VTc, Vz) Then
                                                         DHl1 = PP.DW_CalcEnthalpyDeparture(Vx, T, P, State.Liquid) * PP.AUX_MMM(Vx) / 1000
                                                         DHl2 = PP.DW_CalcEnthalpyDeparture(Vx, Tref, P, State.Liquid) * PP.AUX_MMM(Vx) / 1000
                                                         E = DHl2
                                                         F = (DHl1 - E) / (T - Tref)
                                                     Else
                                                         E = 0
                                                         F = 0
                                                     End If
                                                 End Sub)
                    task1.Start()
                    task2.Start()
                    Task.WaitAll(task1, task2)
                Catch ae As AggregateException
                    Throw ae.Flatten().InnerException
                End Try
                My.MyApplication.IsRunningParallelTasks = False
            Else
                DHv1 = PP.DW_CalcEnthalpyDeparture(Vy, T, P, State.Vapor) * PP.AUX_MMM(Vy) / 1000
                DHv2 = PP.DW_CalcEnthalpyDeparture(Vy, Tref, P, State.Vapor) * PP.AUX_MMM(Vy) / 1000
                If T < Common.Max(VTc, Vz) Then
                    DHl1 = PP.DW_CalcEnthalpyDeparture(Vx, T, P, State.Liquid) * PP.AUX_MMM(Vx) / 1000
                    DHl2 = PP.DW_CalcEnthalpyDeparture(Vx, Tref, P, State.Liquid) * PP.AUX_MMM(Vx) / 1000
                    E = DHl2
                    F = (DHl1 - E) / (T - Tref)
                Else
                    E = 0
                    F = 0
                End If
            End If

            Dim fx(n + 6), x(n + 6), dfdx(n + 6, n + 6), dx(n + 6), xbr(n + 6), fbr(n + 6) As Double
            Dim bo As New BrentOpt.Brent
            Dim bo2 As New BrentOpt.BrentMinimize

            Dim fr As Double

            ecount = 0
            R = Kb * V / (Kb * V + Kb0 * L)

restart:    Do

                '--------------------------------------------------------------
                ' STEPS 2, 3, 4, 5, 6, 7 and 8 - Calculate R and Energy Balance
                '--------------------------------------------------------------

                Rant = R

                bo2.DefineFuncDelegate(AddressOf EnergyBalanceAbs)
                fr = bo2.brentoptimize(0.0#, 1.0#, 0.00000001, R)

                If ecount > 0 Then
                    If Rant = 0.0# And R = 1.0# Then R = 0.0#
                    If Rant = 1.0# And R = 0.0# Then R = 1.0#
                End If

                Me.EnergyBalance(R)

                'At this point, we have converged T and R for the simplified model. Proceed to step 9.

                '----------------------------------------------------------
                ' STEP 9 - Rigorous model Enthalpy and K-values calculation
                '----------------------------------------------------------

                Ki = PP.DW_CalcKvalue(Vx, Vy, T, P)

                If Abs(R - Rant) > 0.01 And Abs(T - Tant) > 0.01 Then
                    For i = 0 To n
                        If Ki(i) <> 0 Then
                            uic(i) = Log(Ki(i) / Kb)
                        Else
                            uic(i) = ui(i)
                        End If
                    Next
                End If

                Bc = Log(Kb_ / Kb) / (1 / T_ - 1 / T)
                Ac = Log(Kb) - Bc * (1 / T - 1 / T_)

                If My.MyApplication._EnableParallelProcessing Then
                    My.MyApplication.IsRunningParallelTasks = True
                    Try
                        Dim task1 As Task = New Task(Sub()
                                                         DHv1 = PP.DW_CalcEnthalpyDeparture(Vy, T, P, State.Vapor) * PP.AUX_MMM(Vy) / 1000
                                                         DHv2 = PP.DW_CalcEnthalpyDeparture(Vy, T0, P, State.Vapor) * PP.AUX_MMM(Vy) / 1000
                                                         Cc = DHv2
                                                         Dc = (DHv1 - Cc) / (T - T0)
                                                     End Sub)
                        Dim task2 As Task = New Task(Sub()
                                                         If T < Common.Max(VTc, Vz) Then
                                                             DHl1 = PP.DW_CalcEnthalpyDeparture(Vx, T, P, State.Liquid) * PP.AUX_MMM(Vx) / 1000
                                                             DHl2 = PP.DW_CalcEnthalpyDeparture(Vx, T0, P, State.Liquid) * PP.AUX_MMM(Vx) / 1000
                                                             Ec = DHl2
                                                             Fc = (DHl1 - Ec) / (T - T0)
                                                         Else
                                                             Ec = 0
                                                             Fc = 0
                                                         End If
                                                     End Sub)
                        task1.Start()
                        task2.Start()
                        Task.WaitAll(task1, task2)
                    Catch ae As AggregateException
                        Throw ae.Flatten().InnerException
                    End Try
                    My.MyApplication.IsRunningParallelTasks = False
                Else
                    DHv1 = PP.DW_CalcEnthalpyDeparture(Vy, T, P, State.Vapor) * PP.AUX_MMM(Vy) / 1000
                    DHv2 = PP.DW_CalcEnthalpyDeparture(Vy, T0, P, State.Vapor) * PP.AUX_MMM(Vy) / 1000
                    Cc = DHv2
                    Dc = (DHv1 - Cc) / (T - T0)
                    If T < Common.Max(VTc, Vz) Then
                        DHl1 = PP.DW_CalcEnthalpyDeparture(Vx, T, P, State.Liquid) * PP.AUX_MMM(Vx) / 1000
                        DHl2 = PP.DW_CalcEnthalpyDeparture(Vx, T0, P, State.Liquid) * PP.AUX_MMM(Vx) / 1000
                        Ec = DHl2
                        Fc = (DHl1 - Ec) / (T - T0)
                    Else
                        Ec = 0
                        Fc = 0
                    End If
                End If

                '-------------------------------------------
                ' STEP 10 - Update variables using Broyden
                '-------------------------------------------

                For i = 0 To n
                    fx(i) = (ui(i) - uic(i))
                    x(i) = ui(i)
                Next
                fx(n + 1) = (A - Ac)
                fx(n + 2) = (B - Bc)
                fx(n + 3) = (C - Cc)
                fx(n + 4) = (D - Dc)
                fx(n + 5) = (E - Ec)
                fx(n + 6) = (F - Fc)
                x(n + 1) = A
                x(n + 2) = B
                x(n + 3) = C
                x(n + 4) = D
                x(n + 5) = E
                x(n + 6) = F

                If ecount = 0 Then
                    For i = 0 To n + 6
                        For j = 0 To n + 6
                            If i = j Then dfdx(i, j) = 1 Else dfdx(i, j) = 0
                        Next
                    Next
                    Broyden.broydn(n + 6, x, fx, dx, xbr, fbr, dfdx, 0)
                Else
                    Broyden.broydn(n + 6, x, fx, dx, xbr, fbr, dfdx, 1)
                End If

                Dim bo3 As New BrentOpt.BrentMinimize
                bo3.DefineFuncDelegate(AddressOf MinimizeError)
                Dim alpha As Double = 1.0#, err As Double

                tmpdx = dx.Clone
                currx = x.Clone

                ReDim refx(n + 6)

                For i = 0 To n
                    refx(i) = uic(i)
                Next
                refx(n + 1) = Ac
                refx(n + 2) = Bc
                refx(n + 3) = Cc
                refx(n + 4) = Dc
                refx(n + 5) = Ec
                refx(n + 6) = Fc

                err = 0

                If PP.Parameters("PP_FLASHALGORITHMFASTMODE") = 0 Then err = bo3.brentoptimize(0, 2, 0.0001, alpha)

                For i = 0 To n
                    ui(i) = ui(i) + alpha * dx(i)
                Next
                A += alpha * dx(n + 1)
                B += alpha * dx(n + 2)
                C += alpha * dx(n + 3)
                D += alpha * dx(n + 4)
                E += alpha * dx(n + 5)
                F += alpha * dx(n + 6)

                ecount += 1

                If ecount > maxit_e Then Throw New Exception(App.GetLocalString("PropPack_FlashMaxIt"))
                If Double.IsNaN(AbsSum(fx)) Then
                    Throw New Exception(App.GetLocalString("PropPack_FlashError"))
                End If

                WriteDebugInfo("PH Flash [IO]: Iteration #" & ecount & ", T = " & T)
                WriteDebugInfo("PH Flash [IO]: Iteration #" & ecount & ", VF = " & V)
                WriteDebugInfo("PH Flash [IO]: Iteration #" & ecount & ", H error = " & fr)
                WriteDebugInfo("PH Flash [IO]: Iteration #" & ecount & ", Damping Factor = " & alpha)

            Loop Until SumSqr(fx) < etol

            If Abs(fr) > itol Then
                If V <= 0.01 Then
                    'single phase solution found (liquid only). Obtain T using single phase calculation.
                    Dim x1, fx2, dfdx2 As Double
                    ecount = 0
                    If Tref = 0 Then Tref = 298.15
                    x1 = Tref
                    Do
                        fx2 = EnergyBalanceSPL(x1, Nothing)
                        If Abs(fx2) < etol Then Exit Do
                        dfdx2 = (EnergyBalanceSPL(x1 + 1, Nothing) - fx2)
                        x1 = x1 - fx2 / dfdx2
                        ecount += 1
                    Loop Until ecount > maxit_e Or Double.IsNaN(x1)
                    T = x1
                    Vx = Vz
                Else
                    'single phase solution found (vapor only). Obtain T using single phase calculation.
                    Dim x1, fx2, dfdx2 As Double
                    ecount = 0
                    If Tref = 0 Then Tref = 298.15
                    x1 = Tref
                    Do
                        fx2 = EnergyBalanceSPV(x1, Nothing)
                        If Abs(fx2) < etol Then Exit Do
                        dfdx2 = (EnergyBalanceSPV(x1 + 1, Nothing) - fx2)
                        x1 = x1 - fx2 / dfdx2
                        ecount += 1
                    Loop Until ecount > maxit_e Or Double.IsNaN(x1)
                    T = x1
                    Vy = Vz
                End If
                'confirm single-phase solution with a PT Flash.
                Dim res As Object = Me.Flash_PT(Vz, P, T, PP, False, Nothing)
                If Abs(L - res(0)) > 0.0001 And Abs(V - res(1)) > 0.0001 Then
                    'NOT SP solution. go back to 2-phase loop.
                    GoTo restart
                End If
            End If

            d2 = Date.Now

            dt = d2 - d1

            WriteDebugInfo("PH Flash [IO]: Converged in " & ecount & " iterations. Time taken: " & dt.TotalMilliseconds & " ms. Error function value: " & AbsSum(fx))

            Return New Object() {L, V, Vx, Vy, T, ecount, Ki, 0.0#, PP.RET_NullVector, 0.0#, PP.RET_NullVector}

        End Function

        Public Overrides Function Flash_PS(ByVal Vz As Double(), ByVal P As Double, ByVal S As Double, ByVal Tref As Double, ByVal PP As PropertyPackage, Optional ByVal ReuseKI As Boolean = False, Optional ByVal PrevKi As Double() = Nothing) As Object

            Dim d1, d2 As Date, dt As TimeSpan
            Dim i, j As Integer

            d1 = Date.Now

            maxit_i = CInt(PP.Parameters("PP_PSFMII"))
            maxit_e = CInt(PP.Parameters("PP_PSFMEI"))
            Me.itol = PP.Parameters("PP_PSFILT")
            Me.etol = PP.Parameters("PP_PSFELT")

            Me.n = Vz.Length - 1

            proppack = PP
            Sf = S * PP.AUX_MMM(Vz)
            Pf = P

            ReDim Vn(n), Vx(n), Vy(n), Vp(n), ui(n), uic(n), pi(n), Ki(n), fi(n), Vt(n), Vpc(n), VTc(n), Vw(n)

            Vn = PP.RET_VNAMES()
            VTc = PP.RET_VTC
            fi = Vz.Clone

            Tmin = 0
            Tmax = 0
            If Tref = 0 Or Double.IsNaN(Tref) Then
                i = 0
                Tref = 0
                Do
                    Tref += 0.8 * Vz(i) * VTc(i)
                    Tmin += 0.1 * Vz(i) * VTc(i)
                    Tmax += 2.0 * Vz(i) * VTc(i)
                    i += 1
                Loop Until i = n + 1
            Else
                Tmin = Tref - 200
                Tmax = Tref + 200
            End If
            If Tmin < 100 Then Tmin = 100

            '--------------------------------------
            ' STEP 1 - Assume u, A, B, C, D, E, F 
            '--------------------------------------

            T = Tref - 0.1
            T_ = Tref - 0.2
            T0 = Tref

            '----------------------------------------
            ' STEP 1.1 - Estimate K, Vx, Vy, V and L 
            '----------------------------------------

            'Calculate Ki`s

            Vpc = PP.RET_VPC
            VTc = PP.RET_VTC
            Vw = PP.RET_VW

            If Not ReuseKI Then
                i = 0
                Do
                    Vp(i) = PP.AUX_PVAPi(Vn(i), T)
                    Ki(i) = Vpc(i) / P * Math.Exp(5.373 * (1 + Vw(i)) * (1 - VTc(i) / T))
                    i += 1
                Loop Until i = n + 1
            Else
                For i = 0 To n
                    Vp(i) = PP.AUX_PVAPi(Vn(i), T)
                    Ki(i) = PrevKi(i)
                Next
            End If

            'Estimate V

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

            If P <= Pd Then
                'vapor only
                L = 0
                V = 1
            ElseIf P >= Pb Then
                'liquid only
                L = 1
                V = 0
            Else
                'VLE
                V = 1 - (P - Pd) / (Pb - Pd)
                L = 1 - V
            End If

            If n = 0 Then
                If Vp(0) <= P Then
                    L = 1
                    V = 0
                Else
                    L = 0
                    V = 1
                End If
            End If

            If T > Common.Max(proppack.RET_VTC, Vz) Then
                V = 1
                L = 0
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
            sum_x = 0
            sum_y = 0
            Do
                sum_x = sum_x + Vx(i)
                sum_y = sum_y + Vy(i)
                i = i + 1
            Loop Until i = n + 1
            i = 0
            Do
                Vx(i) = Vx(i) / sum_x
                Vy(i) = Vy(i) / sum_y
                i = i + 1
            Loop Until i = n + 1

            Kb_ = CalcKbj1(PP.DW_CalcKvalue(Vx, Vy, T_, P))
            Kb0 = Kb_

            Kb = CalcKbj1(Ki)

            B = Log(Kb_ / Kb) / (1 / T_ - 1 / T)
            A = Log(Kb) - B * (1 / T - 1 / T_)

            For i = 0 To n
                ui(i) = Log(Ki(i) / Kb)
            Next

            If My.MyApplication._EnableParallelProcessing Then
                My.MyApplication.IsRunningParallelTasks = True
                Try
                    Dim task1 As Task = New Task(Sub()
                                                     DSv1 = PP.DW_CalcEntropyDeparture(Vy, T, P, State.Vapor) * PP.AUX_MMM(Vy)
                                                     DSv2 = PP.DW_CalcEntropyDeparture(Vy, Tref, P, State.Vapor) * PP.AUX_MMM(Vy)
                                                     C = DSv2
                                                     D = (DSv1 - C) / (T - Tref)
                                                 End Sub)
                    Dim task2 As Task = New Task(Sub()
                                                     If T < Common.Max(VTc, Vz) Then
                                                         DSl1 = PP.DW_CalcEntropyDeparture(Vx, T, P, State.Liquid) * PP.AUX_MMM(Vx)
                                                         DSl2 = PP.DW_CalcEntropyDeparture(Vx, Tref, P, State.Liquid) * PP.AUX_MMM(Vx)
                                                         E = DSl2
                                                         F = (DSl1 - E) / (T - Tref)
                                                     Else
                                                         E = 0
                                                         F = 0
                                                     End If
                                                 End Sub)
                    task1.Start()
                    task2.Start()
                    Task.WaitAll(task1, task2)
                Catch ae As AggregateException
                    Throw ae.Flatten().InnerException
                End Try
                My.MyApplication.IsRunningParallelTasks = False
            Else
                DSv1 = PP.DW_CalcEntropyDeparture(Vy, T, P, State.Vapor) * PP.AUX_MMM(Vy)
                DSv2 = PP.DW_CalcEntropyDeparture(Vy, Tref, P, State.Vapor) * PP.AUX_MMM(Vy)
                C = DSv2
                D = (DSv1 - C) / (T - Tref)
                If T < Common.Max(VTc, Vz) Then
                    DSl1 = PP.DW_CalcEntropyDeparture(Vx, T, P, State.Liquid) * PP.AUX_MMM(Vx)
                    DSl2 = PP.DW_CalcEntropyDeparture(Vx, Tref, P, State.Liquid) * PP.AUX_MMM(Vx)
                    E = DSl2
                    F = (DSl1 - E) / (T - Tref)
                Else
                    E = 0
                    F = 0
                End If
            End If

            Dim fx(n + 6), x(n + 6), dfdx(n + 6, n + 6), dx(n + 6), xbr(n + 6), fbr(n + 6) As Double
            Dim bo As New BrentOpt.Brent
            Dim bo2 As New BrentOpt.BrentMinimize

            Dim fr As Double

            ecount = 0
            R = Kb * V / (Kb * V + Kb0 * L)

restart:    Do

                '--------------------------------------------------------------
                ' STEPS 2, 3, 4, 5, 6, 7 and 8 - Calculate R and Entropy Balance
                '--------------------------------------------------------------

                Rant = R

                bo2.DefineFuncDelegate(AddressOf EntropyBalanceAbs)
                fr = bo2.brentoptimize(0.0#, 1.0#, 0.0001, R)

                Me.EntropyBalance(R)

                'At this point, we have converged T and R for the simplified model. Proceed to step 9.

                '----------------------------------------------------------
                ' STEP 9 - Rigorous model Entropy and K-values calculation
                '----------------------------------------------------------

                Ki = PP.DW_CalcKvalue(Vx, Vy, T, P)

                If Abs(R - Rant) > 0.00001 And Abs(T - Tant) > 0.01 Then
                    For i = 0 To n
                        uic(i) = Log(Ki(i) / Kb)
                    Next
                End If

                Bc = Log(Kb_ / Kb) / (1 / T_ - 1 / T)
                Ac = Log(Kb) - Bc * (1 / T - 1 / T_)

                If My.MyApplication._EnableParallelProcessing Then
                    My.MyApplication.IsRunningParallelTasks = True
                    Try
                        Dim task1 As Task = New Task(Sub()
                                                         DSv1 = PP.DW_CalcEntropyDeparture(Vy, T, P, State.Vapor) * PP.AUX_MMM(Vy)
                                                         DSv2 = PP.DW_CalcEntropyDeparture(Vy, T0, P, State.Vapor) * PP.AUX_MMM(Vy)
                                                         Cc = DSv2
                                                         Dc = (DSv1 - Cc) / (T - T0)
                                                     End Sub)
                        Dim task2 As Task = New Task(Sub()
                                                         If T < Common.Max(VTc, Vz) Then
                                                             DSl1 = PP.DW_CalcEntropyDeparture(Vx, T, P, State.Liquid) * PP.AUX_MMM(Vx)
                                                             DSl2 = PP.DW_CalcEntropyDeparture(Vx, T0, P, State.Liquid) * PP.AUX_MMM(Vx)
                                                             Ec = DSl2
                                                             Fc = (DSl1 - Ec) / (T - T0)
                                                         Else
                                                             Ec = 0
                                                             Fc = 0
                                                         End If
                                                     End Sub)
                        task1.Start()
                        task2.Start()
                        Task.WaitAll(task1, task2)
                    Catch ae As AggregateException
                        Throw ae.Flatten().InnerException
                    End Try
                    My.MyApplication.IsRunningParallelTasks = False
                Else
                    DSv1 = PP.DW_CalcEntropyDeparture(Vy, T, P, State.Vapor) * PP.AUX_MMM(Vy)
                    DSv2 = PP.DW_CalcEntropyDeparture(Vy, T0, P, State.Vapor) * PP.AUX_MMM(Vy)
                    Cc = DSv2
                    Dc = (DSv1 - Cc) / (T - T0)
                    If T < Common.Max(VTc, Vz) Then
                        DSl1 = PP.DW_CalcEntropyDeparture(Vx, T, P, State.Liquid) * PP.AUX_MMM(Vx)
                        DSl2 = PP.DW_CalcEntropyDeparture(Vx, T0, P, State.Liquid) * PP.AUX_MMM(Vx)
                        Ec = DSl2
                        Fc = (DSl1 - Ec) / (T - T0)
                    Else
                        Ec = 0
                        Fc = 0
                    End If
                End If

                '-------------------------------------------
                ' STEP 10 - Update variables using Broyden
                '-------------------------------------------

                For i = 0 To n
                    fx(i) = (ui(i) - uic(i))
                    x(i) = ui(i)
                Next
                fx(n + 1) = (A - Ac)
                fx(n + 2) = (B - Bc)
                fx(n + 3) = (C - Cc)
                fx(n + 4) = (D - Dc)
                fx(n + 5) = (E - Ec)
                fx(n + 6) = (F - Fc)
                x(n + 1) = A
                x(n + 2) = B
                x(n + 3) = C
                x(n + 4) = D
                x(n + 5) = E
                x(n + 6) = F

                If ecount = 0 Then
                    For i = 0 To n + 6
                        For j = 0 To n + 6
                            If i = j Then dfdx(i, j) = 1 Else dfdx(i, j) = 0
                        Next
                    Next
                    Broyden.broydn(n + 6, x, fx, dx, xbr, fbr, dfdx, 0)
                Else
                    Broyden.broydn(n + 6, x, fx, dx, xbr, fbr, dfdx, 1)
                End If

                Dim bo3 As New BrentOpt.BrentMinimize
                bo3.DefineFuncDelegate(AddressOf MinimizeError)
                Dim alpha As Double = 1.0#, err As Double

                tmpdx = dx.Clone
                currx = x.Clone

                ReDim refx(n + 6)

                For i = 0 To n
                    refx(i) = uic(i)
                Next
                refx(n + 1) = Ac
                refx(n + 2) = Bc
                refx(n + 3) = Cc
                refx(n + 4) = Dc
                refx(n + 5) = Ec
                refx(n + 6) = Fc

                err = 0
                If PP.Parameters("PP_FLASHALGORITHMFASTMODE") = 0 Then err = bo3.brentoptimize(0, 2, 0.0001, alpha)

                For i = 0 To n
                    ui(i) = ui(i) + alpha * dx(i)
                Next
                A += alpha * dx(n + 1)
                B += alpha * dx(n + 2)
                C += alpha * dx(n + 3)
                D += alpha * dx(n + 4)
                E += alpha * dx(n + 5)
                F += alpha * dx(n + 6)

                ecount += 1

                If ecount > maxit_e Then Throw New Exception(App.GetLocalString("PropPack_FlashMaxIt"))
                If Double.IsNaN(AbsSum(fx)) Then Throw New Exception(App.GetLocalString("PropPack_FlashError"))

                WriteDebugInfo("PS Flash [IO]: Iteration #" & ecount & ", T = " & T)
                WriteDebugInfo("PS Flash [IO]: Iteration #" & ecount & ", VF = " & V)
                WriteDebugInfo("PS Flash [IO]: Iteration #" & ecount & ", H error = " & fr)
                WriteDebugInfo("PS Flash [IO]: Iteration #" & ecount & ", Damping Factor = " & alpha)

            Loop Until AbsSum(fx) < etol

            If Abs(fr) > itol Then
                If V <= 0.01 Then
                    'single phase solution found (liquid only). Obtain T using single phase calculation.
                    Dim x1, fx2, dfdx2 As Double
                    ecount = 0
                    If Tref = 0 Then Tref = 298.15
                    x1 = Tref
                    Do
                        fx2 = EntropyBalanceSPL(x1, Nothing)
                        If Abs(fx2) < etol Then Exit Do
                        dfdx2 = (EntropyBalanceSPL(x1 + 1, Nothing) - fx2)
                        x1 = x1 - fx2 / dfdx2
                        ecount += 1
                    Loop Until ecount > maxit_e Or Double.IsNaN(x1)
                    T = x1
                    Vx = Vz
                Else
                    'single phase solution found (vapor only). Obtain T using single phase calculation.
                    Dim x1, fx2, dfdx2 As Double
                    ecount = 0
                    If Tref = 0 Then Tref = 298.15
                    x1 = Tref
                    Do
                        fx2 = EntropyBalanceSPV(x1, Nothing)
                        If Abs(fx2) < etol Then Exit Do
                        dfdx2 = (EntropyBalanceSPV(x1 + 1, Nothing) - fx2)
                        x1 = x1 - fx2 / dfdx2
                        ecount += 1
                    Loop Until ecount > maxit_e Or Double.IsNaN(x1)
                    T = x1
                    Vy = Vz
                End If
                'confirm single-phase solution with a PT Flash.
                Dim res As Object = Me.Flash_PT(Vz, P, T, PP, False, Nothing)
                If Abs(L - res(0)) > 0.0001 And Abs(V - res(1)) > 0.0001 Then
                    'NOT single-phase solution. go back to 2-phase loop.
                    GoTo restart
                End If
            End If

            d2 = Date.Now

            dt = d2 - d1

            WriteDebugInfo("PS Flash [IO]: Converged in " & ecount & " iterations. Time taken: " & dt.TotalMilliseconds & " ms. Error function value: " & AbsSum(fx))

            Return New Object() {L, V, Vx, Vy, T, ecount, Ki, 0.0#, PP.RET_NullVector, 0.0#, PP.RET_NullVector}

        End Function
        Public Overrides Function Flash_PV(ByVal Vz As Double(), ByVal P As Double, ByVal V As Double, ByVal Tref As Double, ByVal PP As PropertyPackage, Optional ByVal ReuseKI As Boolean = False, Optional ByVal PrevKi As Double() = Nothing) As Object

            Dim d1, d2 As Date, dt As TimeSpan
            Dim i, j As Integer

            d1 = Date.Now

            Me.etol = PP.Parameters("PP_PTFELT")
            Me.maxit_e = CInt(PP.Parameters("PP_PTFMEI"))
            Me.itol = PP.Parameters("PP_PTFILT")
            Me.maxit_i = CInt(PP.Parameters("PP_PTFMII"))

            n = Vz.Length - 1

            proppack = PP
            Vf = V
            L = 1 - V
            Lf = 1 - Vf
            Pf = P

            ReDim Vn(n), Vx(n), Vy(n), Vp(n), ui(n), uic(n), pi(n), Ki(n), fi(n)
            Dim Vt(n), VTc(n), Tmin, Tmax As Double

            Vn = PP.RET_VNAMES()
            VTc = PP.RET_VTC
            fi = Vz.Clone

            '--------------------------------------
            ' STEP 1 - Assume u, A, B, C, D, E, F 
            '--------------------------------------

            Tmin = 0
            Tmax = 0
            If Tref = 0 Then
                i = 0
                Tref = 0
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

            Dim fx(n + 2), x(n + 2), dfdx(n + 2, n + 2), dx(n + 2), xbr(n + 2), fbr(n + 2) As Double

            T = Tref
            T_ = T - 2
            T0 = T - 5

            If PP.AUX_IS_SINGLECOMP(Vz) Then
                i = 0
                T = 0
                Do
                    If Vz(i) > 0 Then
                        T += Vz(i) * PP.AUX_TSATi(P, i)
                        Vp(i) = PP.AUX_PVAPi(Vn(i), T)
                        Ki(i) = Vp(i) / P
                    End If
                    i += 1
                Loop Until i = n + 1
                Vx = Vz
                Vy = Vz
                GoTo final
            End If

            '----------------------------------------
            ' STEP 1.1 - Estimate K, Vx, Vy, V and L 
            '----------------------------------------

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
            sum_x = 0
            sum_y = 0
            Do
                sum_x = sum_x + Vx(i)
                sum_y = sum_y + Vy(i)
                i = i + 1
            Loop Until i = n + 1
            i = 0
            Do
                Vx(i) = Vx(i) / sum_x
                Vy(i) = Vy(i) / sum_y
                i = i + 1
            Loop Until i = n + 1

            Kb_ = CalcKbj1(PP.DW_CalcKvalue(Vx, Vy, T_, P))
            Kb0 = CalcKbj1(PP.DW_CalcKvalue(Vx, Vy, T0, P))
            Kb = CalcKbj1(PP.DW_CalcKvalue(Vx, Vy, T, P))

            B = Log(Kb_ / Kb) / (1 / T_ - 1 / T)
            A = Log(Kb) - B * (1 / T - 1 / T0)

            For i = 0 To n
                ui(i) = Log(Ki(i) / Kb)
            Next

            Dim RLoop As Boolean = True

            If V = 0.0# Or V = 1.0# Then RLoop = False

            ecount = 0

            Do

                '--------------------------------------------------------------
                ' STEPS 2, 3, 4, 5, 6, 7 and 8 - Calculate R and Energy Balance
                '--------------------------------------------------------------

                R = Kb * V / (Kb * V + Kb0 * L)

                If RLoop Then

                    Dim fr, dfr, R0, R1 As Double
                    Dim icount As Integer = 0

                    Do
                        R1 = R + 0.001
                        fr = Me.LiquidFractionBalance(R)
                        dfr = (fr - Me.LiquidFractionBalance(R1)) / -0.001
                        R0 = R
                        R += -fr / dfr
                        If R < 0 Then R = 0
                        If R > 1 Then R = 1
                        icount += 1
                    Loop Until Abs(fr) < itol Or icount > maxit_i Or Abs(R - R0) < 0.000001

                Else

                    For i = 0 To n
                        pi(i) = fi(i) / (1 - R + Kb0 * R * Exp(ui(i)))
                    Next

                    Dim sumpi As Double = 0
                    Dim sumeuipi As Double = 0

                    For i = 0 To n
                        sumpi += pi(i)
                        sumeuipi += Exp(ui(i)) * pi(i)
                    Next

                    Kb = sumpi / sumeuipi

                    Dim Tant As Double = T
                    T = 1 / T0 + (Log(Kb) - A) / B
                    T = 1 / T

                    For i = 0 To n
                        Vx(i) = pi(i) / sumpi
                        Vy(i) = Exp(ui(i)) * pi(i) / sumeuipi
                    Next

                End If

                'At this point, we have converged T for the simplified model. Proceed to step 9.

                '----------------------------------------------------------
                ' STEP 9 - Rigorous model Enthalpy and K-values calculation
                '----------------------------------------------------------

                Ki = PP.DW_CalcKvalue(Vx, Vy, T, P)
                'Kb_ = CalcKbj1(PP.DW_CalcKvalue(Vx, Vy, T_, P))

                For i = 0 To n
                    uic(i) = Log(Ki(i) / Kb)
                Next

                Bc = Log(Kb_ / Kb) / (1 / T_ - 1 / T)
                Ac = Log(Kb) - Bc * (1 / T - 1 / T0)

                '-------------------------------------------
                ' STEP 10 - Update variables using Broyden
                '-------------------------------------------

                For i = 0 To n
                    fx(i) = (ui(i) - uic(i))
                    x(i) = ui(i)
                Next
                fx(n + 1) = (A - Ac)
                fx(n + 2) = (B - Bc)
                x(n + 1) = A
                x(n + 2) = B

                If ecount = 0 Then
                    For i = 0 To n + 2
                        For j = 0 To n + 2
                            If i = j Then dfdx(i, j) = 1 Else dfdx(i, j) = 0
                        Next
                    Next
                    Broyden.broydn(n + 2, x, fx, dx, xbr, fbr, dfdx, 0)
                Else
                    Broyden.broydn(n + 2, x, fx, dx, xbr, fbr, dfdx, 1)
                End If

                Dim bo2 As New BrentOpt.BrentMinimize
                bo2.DefineFuncDelegate(AddressOf MinimizeError)
                Dim alpha As Double = 1.0#, err As Double

                tmpdx = dx.Clone
                currx = x.Clone

                ReDim refx(n + 2)

                For i = 0 To n
                    refx(i) = uic(i)
                Next
                refx(n + 1) = Ac
                refx(n + 2) = Bc

                err = 0
                If PP.Parameters("PP_FLASHALGORITHMFASTMODE") = 0 Then err = bo2.brentoptimize(0, 2, 0.0001, alpha)

                For i = 0 To n
                    ui(i) = ui(i) + alpha * dx(i)
                Next
                A += alpha * dx(n + 1)
                B += alpha * dx(n + 2)

                ecount += 1

                If ecount > maxit_e Then
                    Throw New Exception(App.GetLocalString("PropPack_FlashMaxIt"))
                End If
                If Double.IsNaN(AbsSum(fx)) Then
                    Throw New Exception(App.GetLocalString("PropPack_FlashError"))
                End If

                WriteDebugInfo("PV Flash [IO]: Iteration #" & ecount & ", T = " & T & ", VF = " & V)
                WriteDebugInfo("PV Flash [IO]: Iteration #" & ecount & ", Damping Factor = " & alpha)

            Loop Until AbsSum(fx) < etol * (n + 2)

final:      d2 = Date.Now

            dt = d2 - d1

            If PP.AUX_CheckTrivial(Ki) Then Throw New Exception("PV Flash [IO]: Invalid result: converged to the trivial solution (T = " & T & " ).")

            WriteDebugInfo("PV Flash [IO]: Converged in " & ecount & " iterations. Time taken: " & dt.TotalMilliseconds & " ms. Error function value: " & AbsSum(fx))

            Return New Object() {L, V, Vx, Vy, T, ecount, Ki, 0.0#, PP.RET_NullVector, 0.0#, PP.RET_NullVector}

        End Function
        Public Overrides Function Flash_TV(ByVal Vz As Double(), ByVal T As Double, ByVal V As Double, ByVal Pref As Double, ByVal PP As PropertyPackage, Optional ByVal ReuseKI As Boolean = False, Optional ByVal PrevKi As Double() = Nothing) As Object

            Dim d1, d2 As Date, dt As TimeSpan
            Dim i, j As Integer

            d1 = Date.Now

            Me.etol = PP.Parameters("PP_PTFELT")
            Me.maxit_e = CInt(PP.Parameters("PP_PTFMEI"))

            n = Vz.Length - 1

            proppack = PP
            Vf = V
            L = 1 - V
            Lf = 1 - Vf
            Tf = T

            ReDim Vn(n), Vx(n), Vy(n), Vp(n), ui(n), uic(n), pi(n), Ki(n), fi(n)

            Dim VTc = PP.RET_VTC()

            Vn = PP.RET_VNAMES()
            fi = Vz.Clone

            If Pref = 0 Then
                i = 0
                Do
                    If T / VTc(i) <= 0.9 Then
                        Vp(i) = PP.AUX_PVAPi(Vn(i), T)
                    End If
                    i += 1
                Loop Until i = n + 1
                Pmin = Common.Min(Vp)
                Pmax = Common.Max(Vp)
                Pref = Pmin + (1 - V) * (Pmax - Pmin)
            Else
                Pmin = Pref * 0.8
                Pmax = Pref * 1.2
            End If

            Dim fx(n + 2), x(n + 2), dfdx(n + 2, n + 2), dx(n + 2), xbr(n + 2), fbr(n + 2) As Double

            P = Pref
            P_ = Pref * 1.05
            P0 = Pref * 0.95

            If PP.AUX_IS_SINGLECOMP(Vz) Then
                i = 0
                P = 0
                Do
                    Vp(i) = PP.AUX_PVAPi(Vn(i), T)
                    P += Vz(i) * Vp(i)
                    Ki(i) = Vz(i)
                    i += 1
                Loop Until i = n + 1
                Vx = Vz
                Vy = Vz
                GoTo final
            End If

            '----------------------------------------
            ' STEP 1.1 - Estimate K, Vx, Vy, V and L 
            '----------------------------------------

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
                    Vx(i) = Vy(i) / Ki(i)
                Else
                    Vy(i) = 0
                    Vx(i) = 0
                End If
                i += 1
            Loop Until i = n + 1

            i = 0
            sum_x = 0
            sum_y = 0
            Do
                sum_x = sum_x + Vx(i)
                sum_y = sum_y + Vy(i)
                i = i + 1
            Loop Until i = n + 1
            i = 0
            Do
                Vx(i) = Vx(i) / sum_x
                Vy(i) = Vy(i) / sum_y
                i = i + 1
            Loop Until i = n + 1

            Kb0 = CalcKbj1(PP.DW_CalcKvalue(Vx, Vy, T, P0))
            Kb_ = CalcKbj1(PP.DW_CalcKvalue(Vx, Vy, T, P_))
            Kb = CalcKbj1(PP.DW_CalcKvalue(Vx, Vy, T, P))

            B = Log(Kb_ * P_ / (Kb0 * P0)) / Log(P_ / P0)
            A = Log(Kb * P) - B * Log(P / P0)

            For i = 0 To n
                ui(i) = Log(Ki(i) / Kb)
            Next

            Dim RLoop As Boolean = True

            If V = 0.0# Or V = 1.0# Then RLoop = False

            ecount = 0

            Do

                '--------------------------------------------------------------
                ' STEPS 2, 3, 4, 5, 6, 7 and 8 - Calculate R and Energy Balance
                '--------------------------------------------------------------

                R = Kb * V / (Kb * V + Kb0 * L)

                If RLoop Then

                    Dim fr, dfr, R0, R1 As Double
                    Dim icount As Integer = 0

                    Do
                        R1 = R + 0.001
                        fr = Me.LiquidFractionBalanceP(R)
                        dfr = (fr - Me.LiquidFractionBalanceP(R1)) / -0.001
                        R0 = R
                        R += -fr / dfr
                        If R < 0 Then R = 0
                        If R > 1 Then R = 1
                        icount += 1
                    Loop Until Abs(fr) < itol Or icount > maxit_i

                Else

                    For i = 0 To n
                        pi(i) = fi(i) / (1 - R + Kb0 * R * Exp(ui(i)))
                    Next

                    Dim sumpi As Double = 0
                    Dim sumeuipi As Double = 0

                    For i = 0 To n
                        sumpi += pi(i)
                        sumeuipi += Exp(ui(i)) * pi(i)
                    Next

                    Kb = sumpi / sumeuipi

                    P = Exp((A - Log(Kb) - B * Log(P0)) / (1 - B))

                    For i = 0 To n
                        Vx(i) = pi(i) / sumpi
                        Vy(i) = Exp(ui(i)) * pi(i) / sumeuipi
                    Next

                End If

                'At this point, we have converged T for the simplified model. Proceed to step 9.

                '----------------------------------------------------------
                ' STEP 9 - Rigorous model Enthalpy and K-values calculation
                '----------------------------------------------------------

                Ki = PP.DW_CalcKvalue(Vx, Vy, T, P)
                'Kb_ = CalcKbj1(PP.DW_CalcKvalue(Vx, Vy, T, P_))

                For i = 0 To n
                    uic(i) = Log(Ki(i) / Kb)
                Next

                Bc = Log(Kb_ * P_ / (Kb0 * P0)) / Log(P_ / P0)
                Ac = Log(Kb * P) - Bc * Log(P / P0)

                '-------------------------------------------
                ' STEP 10 - Update variables using Broyden
                '-------------------------------------------

                For i = 0 To n
                    fx(i) = (ui(i) - uic(i))
                    x(i) = ui(i)
                Next
                fx(n + 1) = (A - Ac)
                fx(n + 2) = (B - Bc)
                x(n + 1) = A
                x(n + 2) = B

                If ecount = 0 Then
                    For i = 0 To n + 2
                        For j = 0 To n + 2
                            If i = j Then dfdx(i, j) = 1 Else dfdx(i, j) = 0
                        Next
                    Next
                    Broyden.broydn(n + 2, x, fx, dx, xbr, fbr, dfdx, 0)
                Else
                    Broyden.broydn(n + 2, x, fx, dx, xbr, fbr, dfdx, 1)
                End If

                Dim bo2 As New BrentOpt.BrentMinimize
                bo2.DefineFuncDelegate(AddressOf MinimizeError)
                Dim alpha As Double = 1.0#, err As Double

                tmpdx = dx.Clone
                currx = x.Clone

                ReDim refx(n + 2)

                For i = 0 To n
                    refx(i) = uic(i)
                Next
                refx(n + 1) = Ac
                refx(n + 2) = Bc

                err = 0
                If PP.Parameters("PP_FLASHALGORITHMFASTMODE") = 0 Then err = bo2.brentoptimize(0, 2, 0.0001, alpha)

                For i = 0 To n
                    ui(i) = ui(i) + alpha * dx(i)
                Next
                A += alpha * dx(n + 1)
                B += alpha * dx(n + 2)

                ecount += 1

                If ecount > maxit_e Then Throw New Exception(App.GetLocalString("PropPack_FlashMaxIt"))
                If Double.IsNaN(AbsSum(fx)) Then Throw New Exception(App.GetLocalString("PropPack_FlashError"))

                WriteDebugInfo("TV Flash [IO]: Iteration #" & ecount & ", P = " & P & ", VF = " & V)
                WriteDebugInfo("TV Flash [IO]: Iteration #" & ecount & ", Damping Factor = " & alpha)



            Loop Until AbsSum(fx) < etol * (n + 2)

final:      d2 = Date.Now

            dt = d2 - d1

            If PP.AUX_CheckTrivial(Ki) Then Throw New Exception("TV Flash [IO]: Invalid result: converged to the trivial solution (P = " & P & " ).")

            WriteDebugInfo("TV Flash [IO]: Converged in " & ecount & " iterations. Time taken: " & dt.TotalMilliseconds & " ms. Error function value: " & AbsSum(fx))

            Return New Object() {L, V, Vx, Vy, P, ecount, Ki, 0.0#, PP.RET_NullVector, 0.0#, PP.RET_NullVector}

        End Function
        Private Function LiquidFractionBalance(ByVal R As Double) As Double

            Dim i As Integer

            For i = 0 To n
                pi(i) = fi(i) / (1 - R + Kb0 * R * Exp(ui(i)))
            Next

            Dim sumpi As Double = 0
            Dim sumeuipi As Double = 0

            For i = 0 To n
                sumpi += pi(i)
                sumeuipi += Exp(ui(i)) * pi(i)
            Next

            Kb = sumpi / sumeuipi

            Dim Tant As Double = T
            T = 1 / T0 + (Log(Kb) - A) / B
            T = 1 / T

            For i = 0 To n
                Vx(i) = pi(i) / sumpi
                Vy(i) = Exp(ui(i)) * pi(i) / sumeuipi
            Next

            L = (1 - R) * sumpi
            V = 1 - L

            Dim eberror As Double = L - Lf

            Return eberror

        End Function

        Private Function LiquidFractionBalanceP(ByVal R As Double) As Double

            Dim i As Integer

            For i = 0 To n
                pi(i) = fi(i) / (1 - R + Kb0 * R * Exp(ui(i)))
            Next

            Dim sumpi As Double = 0
            Dim sumeuipi As Double = 0

            For i = 0 To n
                sumpi += pi(i)
                sumeuipi += Exp(ui(i)) * pi(i)
            Next

            Kb = sumpi / sumeuipi

            P = Exp((A - Log(Kb) - B * Log(P0)) / (1 - B))

            For i = 0 To n
                Vx(i) = pi(i) / sumpi
                Vy(i) = Exp(ui(i)) * pi(i) / sumeuipi
            Next

            L = (1 - R) * sumpi
            V = 1 - L

            Dim eberror As Double = L - Lf

            Return eberror

        End Function

        Private Function EnergyBalance(ByVal R As Double) As Double

            Dim i As Integer

            For i = 0 To n
                pi(i) = fi(i) / (1 - R + Kb0 * R * Exp(ui(i)))
            Next

            Dim sumpi As Double = 0
            Dim sumeuipi As Double = 0

            For i = 0 To n
                sumpi += pi(i)
                sumeuipi += Exp(ui(i)) * pi(i)
            Next

            Kb = sumpi / sumeuipi

            T = 1 / T_ + (Log(Kb) - A) / B
            T = 1 / T

            For i = 0 To n
                Vx(i) = pi(i) / sumpi
                Vy(i) = Exp(ui(i)) * pi(i) / sumeuipi
            Next

            L = (1 - R) * sumpi
            V = 1 - L

            DHv = C + D * (T - T0)
            DHl = E + F * (T - T0)

            If Double.IsNaN(DHl) Then DHl = 0

            Hvid = proppack.RET_Hid(298.15, T, Vy) * proppack.AUX_MMM(Vy) / 1000
            Hlid = proppack.RET_Hid(298.15, T, Vx) * proppack.AUX_MMM(Vx) / 1000

            Dim eberror As Double = Hf / 1000 - L * (DHl + Hlid) - V * (DHv + Hvid)

            Return eberror

        End Function

        Private Function EnergyBalanceAbs(ByVal R As Double) As Double

            Dim i As Integer

            For i = 0 To n
                pi(i) = fi(i) / (1 - R + Kb0 * R * Exp(ui(i)))
            Next

            Dim sumpi As Double = 0
            Dim sumeuipi As Double = 0

            For i = 0 To n
                sumpi += pi(i)
                sumeuipi += Exp(ui(i)) * pi(i)
            Next

            Kb = sumpi / sumeuipi

            T = 1 / T_ + (Log(Kb) - A) / B
            T = 1 / T

            For i = 0 To n
                Vx(i) = pi(i) / sumpi
                Vy(i) = Exp(ui(i)) * pi(i) / sumeuipi
            Next

            L = (1 - R) * sumpi
            V = 1 - L

            DHv = C + D * (T - T0)
            DHl = E + F * (T - T0)

            If Double.IsNaN(DHl) Then DHl = 0

            Hvid = proppack.RET_Hid(298.15, T, Vy) * proppack.AUX_MMM(Vy) / 1000
            Hlid = proppack.RET_Hid(298.15, T, Vx) * proppack.AUX_MMM(Vx) / 1000

            Dim eberror As Double = Hf / 1000 - L * (DHl + Hlid) - V * (DHv + Hvid)

            Return Abs(eberror)

        End Function

        Private Function EnergyBalanceSPL(ByVal T As Double, ByVal otherargs As Object) As Double

            Dim HL, balerror As Double

            HL = proppack.DW_CalcEnthalpy(fi, T, Pf, State.Liquid) * proppack.AUX_MMM(fi)

            balerror = Hf - HL

            WriteDebugInfo("PH Flash [IO]: Iteration #" & ecount & ", T = " & T)
            WriteDebugInfo("PH Flash [IO]: Iteration #" & ecount & ", VF = 0 (SP)")
            WriteDebugInfo("PH Flash [IO]: Iteration #" & ecount & ", H error = " & balerror)

            Return balerror

        End Function

        Private Function EnergyBalanceSPV(ByVal T As Double, ByVal otherargs As Object) As Double

            Dim HV, balerror As Double

            HV = proppack.DW_CalcEnthalpy(fi, T, Pf, State.Vapor) * proppack.AUX_MMM(fi)

            balerror = Hf - HV

            WriteDebugInfo("PH Flash [IO]: Iteration #" & ecount & ", T = " & T)
            WriteDebugInfo("PH Flash [IO]: Iteration #" & ecount & ", VF = 1 (SP)")
            WriteDebugInfo("PH Flash [IO]: Iteration #" & ecount & ", H error = " & balerror)

            Return balerror

        End Function

        Private Function EntropyBalance(ByVal R As Double) As Double

            Dim i As Integer

            For i = 0 To n
                pi(i) = fi(i) / (1 - R + Kb0 * R * Exp(ui(i)))
            Next

            Dim sumpi As Double = 0
            Dim sumeuipi As Double = 0

            For i = 0 To n
                sumpi += pi(i)
                sumeuipi += Exp(ui(i)) * pi(i)
            Next

            Kb = sumpi / sumeuipi

            T = 1 / T_ + (Log(Kb) - A) / B
            T = 1 / T

            For i = 0 To n
                Vx(i) = pi(i) / sumpi
                Vy(i) = Exp(ui(i)) * pi(i) / sumeuipi
            Next

            L = (1 - R) * sumpi
            V = 1 - L

            DSv = C + D * (T - T0)
            DSl = E + F * (T - T0)

            Svid = proppack.RET_Sid(298.15, T, Pf, Vy) * proppack.AUX_MMM(Vy)
            Slid = proppack.RET_Sid(298.15, T, Pf, Vx) * proppack.AUX_MMM(Vx)

            Dim eberror As Double = Sf - L * (DSl + Slid) - V * (DSv + Svid)


            Return eberror

        End Function

        Private Function EntropyBalanceAbs(ByVal R As Double) As Double

            Dim i As Integer

            For i = 0 To n
                pi(i) = fi(i) / (1 - R + Kb0 * R * Exp(ui(i)))
            Next

            Dim sumpi As Double = 0
            Dim sumeuipi As Double = 0

            For i = 0 To n
                sumpi += pi(i)
                sumeuipi += Exp(ui(i)) * pi(i)
            Next

            Kb = sumpi / sumeuipi

            T = 1 / T_ + (Log(Kb) - A) / B
            T = 1 / T

            For i = 0 To n
                Vx(i) = pi(i) / sumpi
                Vy(i) = Exp(ui(i)) * pi(i) / sumeuipi
            Next

            L = (1 - R) * sumpi
            V = 1 - L

            DSv = C + D * (T - T0)
            DSl = E + F * (T - T0)

            Svid = proppack.RET_Sid(298.15, T, Pf, Vy) * proppack.AUX_MMM(Vy)
            Slid = proppack.RET_Sid(298.15, T, Pf, Vx) * proppack.AUX_MMM(Vx)

            Dim eberror As Double = Sf - L * (DSl + Slid) - V * (DSv + Svid)

            Return Abs(eberror)

        End Function

        Private Function EntropyBalanceSPL(ByVal T As Double, ByVal otherargs As Object) As Double

            Dim SL, balerror As Double

            SL = proppack.DW_CalcEntropy(fi, T, Pf, State.Liquid) * proppack.AUX_MMM(fi)

            balerror = Sf - SL

            WriteDebugInfo("PS Flash [IO]: Iteration #" & ecount & ", T = " & T)
            WriteDebugInfo("PS Flash [IO]: Iteration #" & ecount & ", VF = 0 (SP)")
            WriteDebugInfo("PS Flash [IO]: Iteration #" & ecount & ", S error = " & balerror)

            Return balerror


        End Function

        Private Function EntropyBalanceSPV(ByVal T As Double, ByVal otherargs As Object) As Double

            Dim SV, balerror As Double

            SV = proppack.DW_CalcEntropy(fi, T, Pf, State.Vapor) * proppack.AUX_MMM(fi)

            balerror = Sf - SV

            WriteDebugInfo("PS Flash [IO]: Iteration #" & ecount & ", T = " & T)
            WriteDebugInfo("PS Flash [IO]: Iteration #" & ecount & ", VF = 1 (SP)")
            WriteDebugInfo("PS Flash [IO]: Iteration #" & ecount & ", S error = " & balerror)

            Return balerror

        End Function

        Private Function CalcKbj1(ByVal K() As Double) As Double

            Dim i As Integer
            Dim n As Integer = UBound(K) - 1

            Dim Kbj1 As Object

            Kbj1 = K(0)
            For i = 1 To n
                If Abs(K(i) - 1) < Abs(Kbj1 - 1) Then Kbj1 = K(i)
            Next

            Return Kbj1

        End Function

        Private Function CalcKbj2(ByVal K() As Double, ByVal K2() As Double, ByVal T1 As Double, ByVal T2 As Double) As Double

            Dim i, j As Integer
            Dim n As Integer = UBound(K) - 1

            Dim Kbj1 As Object
            Dim Kw11(n), Kw21(n) As Object
            Dim wi(n), ti(n), sumti As Double

            For i = 0 To n
                Kw11(i) = K(i)

                Kw21(i) = K2(i)
            Next

            sumti = 0
            For i = 0 To n
                ti(i) = Vy(i) * (Log(Kw21(i)) - Log(Kw11(i))) / (1 / T2 - 1 / T1)
                sumti += ti(i)
            Next

            For i = 0 To n
                wi(i) = ti(i) / sumti
            Next

            Kbj1 = 0
            For i = 0 To n
                Kbj1 += wi(i) * Log(K(i))
                Kbj1 = Exp(Kbj1)
                If Kbj1 < 0 Then
                    Kbj1 = K(i)
                    For j = 1 To n
                        If Abs(K(j) - 1) < Abs(Kbj1 - 1) Then Kbj1 = K(j)
                    Next
                End If
            Next

            Return Kbj1

        End Function

        Private Function TPErrorFunc(ByVal Rt As Double) As Double

            Dim i As Integer

            For i = 0 To n
                pi(i) = fi(i) / (1 - Rt + Kb0 * Rt * Exp(ui(i)))
            Next

            Dim sumpi As Double = 0
            Dim sumeuipi As Double = 0

            For i = 0 To n
                sumpi += pi(i)
                sumeuipi += Exp(ui(i)) * pi(i)
            Next

            For i = 0 To n
                Vx(i) = pi(i) / sumpi
                Vy(i) = Exp(ui(i)) * pi(i) / sumeuipi
            Next

            L = (1 - Rt) * sumpi
            V = 1 - L

            Dim eberror As Double = sumpi / sumeuipi - 1

            Return eberror ^ 2

        End Function

        Private Function MinimizeError(ByVal alpha As Double) As Double

            Dim n As Integer = UBound(tmpdx)
            Dim i As Integer
            Dim errors(n) As Double

            For i = 0 To n
                errors(i) = (refx(i) - (currx(i) + alpha * tmpdx(i))) ^ 2
            Next

            Return errors.Sum()

        End Function

    End Class

End Namespace
