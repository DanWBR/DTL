﻿Imports DTL.DTL.SimulationObjects.PropertyPackages

Module example2

    Sub Main()

        RunTest()

    End Sub

    Sub RunTest()

        Dim dtlc As New DTL.Thermodynamics.Calculator()

        dtlc.Initialize()

        dtlc.SetDebugLevel(0)

        Dim proppacks As String() = dtlc.GetPropPackList()

        'Peng-Robinson Property Package
        Dim prpp As PropertyPackage = dtlc.GetPropPackInstance(proppacks(0))

        'Configures the property package to handle liquid phase instability.
        prpp.StabilityTestKeyCompounds = New String() {"Water"}
        prpp.SetParameterValue(Parameter.FlashAlgorithmFastMode, 0)
        prpp.StabilityTestSeverity = 0

        Dim P As Double = 25 'bar
        Dim T As Double = 94 'C

        Dim compounds As String() = New String() {"Water", "Ethane", "Propane", "Isobutane", "N-butane", "1-butene", "N-pentane"}

        Dim molefractions As Double() = New Double() {0.2, 0.2, 0.2, 0.1, 0.1, 0.1, 0.1}

        'do a three-phase flash calculation using the new generic function

        Dim result As Auxiliary.FlashAlgorithms.FlashCalculationResult = dtlc.CalcEquilibrium(DTL.Thermodynamics.Calculator.FlashCalculationType.PressureTemperature,
                                                                        6, P * 101325, T + 273.15, prpp, compounds, molefractions, Nothing, 0)

        'display the results

        If result.ResultException Is Nothing Then

            Console.Write(vbCrLf)
            Console.WriteLine("Flash calculation results at P = " & P & " bar and T = " & T & " C:")
            Console.Write(vbCrLf)
            Console.WriteLine("Time taken: " & result.TimeTaken.TotalMilliseconds & " ms")
            Console.WriteLine("Iterations: " & result.IterationsTaken)
            Console.Write(vbCrLf)
            Console.WriteLine("Compounds: " & compounds.ToArrayString())
            Console.WriteLine("Mixture mole fractions: " & molefractions.ToArrayString())
            Console.WriteLine("Vapor Phase mole fraction: " & result.GetVaporPhaseMoleFraction.ToString())
            Console.WriteLine("Vapor Phase compound mole fractions: " & result.GetVaporPhaseMoleFractions.ToArrayString())
            Console.WriteLine("Liquid Phase 1 mole fraction: " & result.GetLiquidPhase1MoleFraction.ToString())
            Console.WriteLine("Liquid Phase 1 compound mole fractions: " & result.GetLiquidPhase1MoleFractions.ToArrayString())
            Console.WriteLine("Liquid Phase 2 mole fraction: " & result.GetLiquidPhase2MoleFraction.ToString())
            Console.WriteLine("Liquid Phase 2 compound mole fractions: " & result.GetLiquidPhase2MoleFractions.ToArrayString())

        Else

            Console.Write(vbCrLf)
            Console.WriteLine("Error calculating flash: " & result.ResultException.Message.ToString)
            Console.Write(vbCrLf)

        End If

        Console.Write(vbCrLf)
        Console.WriteLine("Press any key to continue...")
        Console.ReadKey(True)

    End Sub

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToArrayString(vector As Double()) As String

        Dim retstr As String = "{ "
        For Each d In vector
            retstr += d.ToString + ", "
        Next
        retstr.TrimEnd(",")
        retstr += "}"

        Return retstr

    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToArrayString(vector As String()) As String

        Dim retstr As String = "{ "
        For Each s In vector
            retstr += s + ", "
        Next
        retstr.TrimEnd(",")
        retstr += "}"

        Return retstr

    End Function

End Module
