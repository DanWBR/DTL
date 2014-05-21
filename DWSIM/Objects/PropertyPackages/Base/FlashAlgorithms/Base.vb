﻿'    Flash Algorithm Abstract Base Class
'    Copyright 2010 Daniel Wagner O. de Medeiros
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

Imports System.Math
Imports System.Threading.Tasks

Namespace DTL.SimulationObjects.PropertyPackages.Auxiliary.FlashAlgorithms

    ''' <summary>
    ''' This is the base class for the flash algorithms.
    ''' </summary>
    ''' <remarks></remarks>
    <System.Serializable()> Public MustInherit Class FlashAlgorithm

        Public StabSearchSeverity As Integer = 0
        Public StabSearchCompIDs As String() = New String() {}

        Private _P As Double, _Vz, _Vx1est, _Vx2est As Double(), _pp As PropertyPackage

        Sub New()

        End Sub

        Public MustOverride Function Flash_PT(ByVal Vz As Double(), ByVal P As Double, ByVal T As Double, ByVal PP As PropertyPackages.PropertyPackage, Optional ByVal ReuseKI As Boolean = False, Optional ByVal PrevKi As Double() = Nothing) As Object

        Public MustOverride Function Flash_PH(ByVal Vz As Double(), ByVal P As Double, ByVal H As Double, ByVal Tref As Double, ByVal PP As PropertyPackages.PropertyPackage, Optional ByVal ReuseKI As Boolean = False, Optional ByVal PrevKi As Double() = Nothing) As Object

        Public MustOverride Function Flash_PS(ByVal Vz As Double(), ByVal P As Double, ByVal S As Double, ByVal Tref As Double, ByVal PP As PropertyPackages.PropertyPackage, Optional ByVal ReuseKI As Boolean = False, Optional ByVal PrevKi As Double() = Nothing) As Object

        Public MustOverride Function Flash_PV(ByVal Vz As Double(), ByVal P As Double, ByVal V As Double, ByVal Tref As Double, ByVal PP As PropertyPackages.PropertyPackage, Optional ByVal ReuseKI As Boolean = False, Optional ByVal PrevKi As Double() = Nothing) As Object

        Public MustOverride Function Flash_TV(ByVal Vz As Double(), ByVal T As Double, ByVal V As Double, ByVal Pref As Double, ByVal PP As PropertyPackages.PropertyPackage, Optional ByVal ReuseKI As Boolean = False, Optional ByVal PrevKi As Double() = Nothing) As Object

        Public Function BubbleTemperature_LLE(ByVal Vz As Double(), ByVal Vx1est As Double(), ByVal Vx2est As Double(), ByVal P As Double, ByVal Tmin As Double, ByVal Tmax As Double, ByVal PP As PropertyPackages.PropertyPackage) As Double

            _P = P
            _pp = PP
            _Vz = Vz
            _Vx1est = Vx1est
            _Vx2est = Vx2est

            Dim T, err As Double

            Dim bm As New MathEx.BrentOpt.BrentMinimize
            bm.DefineFuncDelegate(AddressOf BubbleTemperature_LLEPerror)

            err = bm.brentoptimize(Tmin, Tmax, 0.0001, T)

            err = BubbleTemperature_LLEPerror(T)

            Return T

        End Function

        Private Function BubbleTemperature_LLEPerror(ByVal x As Double) As Double

            Dim n As Integer = UBound(_Vz)

            Dim Vp(n), fi1(n), fi2(n), act1(n), act2(n), Vx1(n), Vx2(n) As Double

            Dim result As Object = New SimpleLLE() With {.UseInitialEstimatesForPhase1 = True, .UseInitialEstimatesForPhase2 = True,
                                                          .InitialEstimatesForPhase1 = _Vx1est, .InitialEstimatesForPhase2 = _Vx2est}.Flash_PT(_Vz, _P, x, _pp)

            'Dim result As Object = New GibbsMinimization3P() With {.ForceTwoPhaseOnly = False, .StabSearchSeverity = 0, .StabSearchCompIDs = _pp.RET_VNAMES}.Flash_PT(_Vz, _P, x, _pp)

            Vx1 = result(2)
            Vx2 = result(6)
            fi1 = _pp.DW_CalcFugCoeff(Vx1, x, _P, State.Liquid)
            fi2 = _pp.DW_CalcFugCoeff(Vx2, x, _P, State.Liquid)

            Dim i As Integer

            For i = 0 To n
                Vp(i) = _pp.AUX_PVAPi(i, x)
                act1(i) = _P / Vp(i) * fi1(i)
                act2(i) = _P / Vp(i) * fi2(i)
            Next

            Dim err As Double = _P
            For i = 0 To n
                err -= Vx2(i) * act2(i) * Vp(i)
            Next

            Return Math.Abs(err)

        End Function

        Public Function BubblePressure_LLE(ByVal Vz As Double(), ByVal Vx1est As Double(), ByVal Vx2est As Double(), ByVal P As Double, ByVal T As Double, ByVal PP As PropertyPackages.PropertyPackage) As Double

            Dim n As Integer = UBound(_Vz)

            Dim Vp(n), fi1(n), fi2(n), act1(n), act2(n), Vx1(n), Vx2(n) As Double

            Dim result As Object = New GibbsMinimization3P() With {.ForceTwoPhaseOnly = False,
                                                                   .StabSearchCompIDs = _pp.RET_VNAMES,
                                                                   .StabSearchSeverity = 0}.Flash_PT(_Vz, P, T, PP)

            Vx1 = result(2)
            Vx2 = result(6)
            fi1 = _pp.DW_CalcFugCoeff(Vx1, T, P, State.Liquid)
            fi2 = _pp.DW_CalcFugCoeff(Vx2, T, P, State.Liquid)

            Dim i As Integer

            For i = 0 To n
                Vp(i) = _pp.AUX_PVAPi(i, T)
                act1(i) = P / Vp(i) * fi1(i)
                act2(i) = P / Vp(i) * fi2(i)
            Next

            _P = 0.0#
            For i = 0 To n
                _P += Vx2(i) * act2(i) * Vp(i)
            Next

            Return _P

        End Function

        Public Function StabTest(ByVal T As Double, ByVal P As Double, ByVal Vz As Array, ByVal pp As PropertyPackage, Optional ByVal VzArray(,) As Double = Nothing, Optional ByVal searchseverity As Integer = 0)

            Dim i, j, c, n, o, l, nt, maxits As Integer
            n = UBound(Vz)
            nt = UBound(VzArray, 1)

            Dim Y, K As Double(,), tol As Double

            Select Case searchseverity
                Case 0
                    ReDim Y(nt, n)
                    tol = 0.01
                    maxits = 100
                Case 1
                    ReDim Y(nt + 1, n)
                    tol = 0.001
                    maxits = 100
                Case Else
                    ReDim Y(nt + 2, n)
                    tol = 0.0001
                    maxits = 200
            End Select

            For i = 0 To nt
                For j = 0 To n
                    Y(i, j) = VzArray(i, j)
                Next
            Next

            ReDim K(0, n)

            Dim m As Integer = UBound(Y, 1)

            Dim h(n), lnfi_z(n), Y_ant(m, n) As Double

            Dim gl, hl, sl, gv, hv, sv As Double

            If My.MyApplication._EnableParallelProcessing Then
                My.MyApplication.IsRunningParallelTasks = True
                If My.MyApplication._EnableGPUProcessing Then
                    My.MyApplication.gpu.EnableMultithreading()
                End If
                Try
                    Dim task1 As Task = New Task(Sub()
                                                     hl = pp.DW_CalcEnthalpy(Vz, T, P, State.Liquid)
                                                 End Sub)
                    Dim task2 As Task = New Task(Sub()
                                                     sl = pp.DW_CalcEntropy(Vz, T, P, State.Liquid)
                                                 End Sub)
                    Dim task3 As Task = New Task(Sub()
                                                     hv = pp.DW_CalcEnthalpy(Vz, T, P, State.Vapor)
                                                 End Sub)
                    Dim task4 As Task = New Task(Sub()
                                                     sv = pp.DW_CalcEntropy(Vz, T, P, State.Vapor)
                                                 End Sub)
                    task1.Start()
                    task2.Start()
                    task3.Start()
                    task4.Start()
                    Task.WaitAll(task1, task2, task3, task4)
                Catch ae As AggregateException
                    For Each ex As Exception In ae.InnerExceptions
                        Throw ex
                    Next
                Finally
                    If My.MyApplication._EnableGPUProcessing Then
                        My.MyApplication.gpu.DisableMultithreading()
                        My.MyApplication.gpu.FreeAll()
                    End If
                End Try
                My.MyApplication.IsRunningParallelTasks = False
            Else
                hl = pp.DW_CalcEnthalpy(Vz, T, P, State.Liquid)
                sl = pp.DW_CalcEntropy(Vz, T, P, State.Liquid)
                hv = pp.DW_CalcEnthalpy(Vz, T, P, State.Vapor)
                sv = pp.DW_CalcEntropy(Vz, T, P, State.Vapor)
            End If

            gl = hl - T * sl
            gv = hv - T * sv

            If gl <= gv Then
                lnfi_z = pp.DW_CalcFugCoeff(Vz, T, P, State.Liquid)
            Else
                lnfi_z = pp.DW_CalcFugCoeff(Vz, T, P, State.Vapor)
            End If

            For i = 0 To n
                lnfi_z(i) = Log(lnfi_z(i))
            Next

            i = 0
            Do
                h(i) = Log(Vz(i)) + lnfi_z(i)
                i = i + 1
            Loop Until i = n + 1

            If Not VzArray Is Nothing Then
                If searchseverity = 1 Then
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
                        Y(nt + 1, i) = sum0(i) / UBound(VzArray, 1)
                        i = i + 1
                    Loop Until i = n + 1
                End If
                If searchseverity = 2 Then
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
                        Y(nt + 1, i) = sum0(i) / UBound(VzArray, 1)
                        Y(nt + 2, i) = Exp(h(i))
                        i = i + 1
                    Loop Until i = n + 1
                End If
            Else
                i = 0
                Do
                    Y(n + 1, i) = Exp(h(i))
                    i = i + 1
                Loop Until i = n + 1
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

                        If My.MyApplication._EnableParallelProcessing Then
                            My.MyApplication.IsRunningParallelTasks = True
                            If My.MyApplication._EnableGPUProcessing Then
                                My.MyApplication.gpu.EnableMultithreading()
                            End If
                            Try
                                Dim task1 As Task = New Task(Sub()
                                                                 hl = pp.DW_CalcEnthalpy(currcomp, T, P, State.Liquid)
                                                             End Sub)
                                Dim task2 As Task = New Task(Sub()
                                                                 sl = pp.DW_CalcEntropy(currcomp, T, P, State.Liquid)
                                                             End Sub)
                                Dim task3 As Task = New Task(Sub()
                                                                 hv = pp.DW_CalcEnthalpy(currcomp, T, P, State.Vapor)
                                                             End Sub)
                                Dim task4 As Task = New Task(Sub()
                                                                 sv = pp.DW_CalcEntropy(currcomp, T, P, State.Vapor)
                                                             End Sub)
                                task1.Start()
                                task2.Start()
                                task3.Start()
                                task4.Start()
                                Task.WaitAll(task1, task2, task3, task4)
                            Catch ae As AggregateException
                                For Each ex As Exception In ae.InnerExceptions
                                    Throw ex
                                Next
                            Finally
                                If My.MyApplication._EnableGPUProcessing Then
                                    My.MyApplication.gpu.DisableMultithreading()
                                    My.MyApplication.gpu.FreeAll()
                                End If
                            End Try
                            My.MyApplication.IsRunningParallelTasks = False
                        Else
                            hl = pp.DW_CalcEnthalpy(currcomp, T, P, State.Liquid)
                            sl = pp.DW_CalcEntropy(currcomp, T, P, State.Liquid)
                            hv = pp.DW_CalcEnthalpy(currcomp, T, P, State.Vapor)
                            sv = pp.DW_CalcEntropy(currcomp, T, P, State.Vapor)
                        End If

                        gl = hl - T * sl
                        gv = hv - T * sv

                        If gl <= gv Then
                            tmpfug = pp.DW_CalcFugCoeff(currcomp, T, P, State.Liquid)
                        Else
                            tmpfug = pp.DW_CalcFugCoeff(currcomp, T, P, State.Vapor)
                        End If

                        j = 0
                        Do
                            lnfi(i, j) = Log(tmpfug(j))
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
                            If Abs(dY(i, j)) > tol Then finish = False
                            j = j + 1
                        Loop Until j = n + 1
                    End If
                    i = i + 1
                Loop Until i = m + 1

                c = c + 1

                If c > maxits Then Throw New Exception("Stability Test: Maximum Iterations Reached.")

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

    End Class

End Namespace
