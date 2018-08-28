Imports DTL.DTL.SimulationObjects.PropertyPackages
Imports System.Linq
Module test
    Sub Main()
        Dim dtlc As New DTL.Thermodynamics.Calculator()
        dtlc.Initialize()
        dtlc.SetDebugLevel(0)
        'Peng-Robinson Property Package
        Dim prpp As PropertyPackage = dtlc.GetPropPackInstance("Peng-Robinson (PR)")
        Dim FlashAlg As Integer = 2 'Nested loops VLE
        Dim compounds As String() = New String() {"Methane", "Isobutane", "N-decane"}
        Dim molefractions As Double() = New Double() {0.8, 0.1, 0.1}
        Dim P As Double = 80 'bar
        Dim T As Double = 30 'C
        prpp.SetParameterValue(Parameter.FlashAlgorithmFastMode, 0)
        Dim result As Auxiliary.FlashAlgorithms.FlashCalculationResult
        result = dtlc.CalcEquilibrium(DTL.Thermodynamics.Calculator.FlashCalculationType.PressureTemperature,
        FlashAlg, P * 101325, T + 273.15, prpp, compounds, molefractions, Nothing, 0)
        If result.ResultException Is Nothing Then
            Dim vmolefractions() As Double
            Dim lmolefractions() As Double
            Dim fv As Double
            Console.Write("Vapor mole fraction ")
            fv = result.GetVaporPhaseMoleFraction
            Console.Write(fv)
            Console.Write(vbCrLf)
            'Vapor Phase compound mole fractions:
            Console.Write("Vapor compound mole fractions ")
            vmolefractions = result.GetVaporPhaseMoleFractions
            For count = 0 To UBound(vmolefractions)
                Console.Write(vmolefractions(count) & " ")
            Next
            Console.Write(vbCrLf)
            'Liquid Phase 1 compound mole fractions:
            Console.Write("Liquid compound mole fractions ")
            lmolefractions = result.GetLiquidPhase1MoleFractions
            For count = 0 To UBound(lmolefractions)
                Console.Write(lmolefractions(count) & " ")
            Next
            Console.Write(vbCrLf)
            Dim PropValues As Object
            PropValues = dtlc.CalcProp(prpp, "viscosity", "mole", "Vapor", compounds, T + 273.15, P * 101325, vmolefractions)
            Console.Write("Vapor viscosity ")
            Console.Write(PropValues(0).ToString & vbTab)
            Console.Write(vbCrLf)
            PropValues = dtlc.CalcProp(prpp, "viscosity", "mole", "Liquid", compounds, T + 273.15, P * 101325, lmolefractions)
            Console.Write("Liquid viscosity ")
            Console.Write(PropValues(0).ToString & vbTab)
            Console.Write(vbCrLf)
            Console.Write("Surface Tension ")
            PropValues = dtlc.CalcTwoPhaseProp(prpp, "surfaceTension", "UNDEFINED", "Vapor", "Liquid", compounds, T + 273.15, P * 101325, vmolefractions, lmolefractions)
            Console.Write(PropValues(0).ToString & vbTab)
            Console.Write(vbCrLf)
        Else
            Console.Write(vbCrLf)
            Console.WriteLine("Error calculating flash: " & result.ResultException.Message.ToString)
            Console.Write(vbCrLf)
        End If
        Console.Write(vbCrLf)
        Console.WriteLine("Press any key to continue...")
        Console.ReadKey(True)
    End Sub
End Module
