Imports DTL.DTL.SimulationObjects.PropertyPackages

Module Module1

    Sub Main()

        RunTest()

    End Sub

    Sub RunTest()

        Dim dtlc As New DTL.Thermodynamics.Calculator()
        dtlc.Initialize()

        Dim proppacks As String() = dtlc.GetPropPackList()
        Dim prpp As PropertyPackage = dtlc.GetPropPackInstance(proppacks(3))

        prpp.StabilityTestKeyCompounds = New String() {"Water"}
        prpp.StabilityTestSeverity = 0
        prpp._ioquick = True

        Dim P As Double = 10
        Dim T As Double = 25

        Dim result2 As Object(,) = dtlc.PTFlash(prpp, 0, P * 101325, T + 273.15,
                                                New String() {"Methane", "Ethane", "Propane", "N-butane", "N-pentane"},
                                                New Double() {0.2, 0.2, 0.2, 0.2, 0.2})

        Console.Write(vbCrLf)
        Console.WriteLine("Flash calculation results at P = " & P & " bar and T = " & T & " C:")
        Console.Write(vbCrLf)

        Dim i, j As Integer, line As String
        For i = 0 To result2.GetLength(0) - 1
            Select Case i
                Case 0
                    line = "Phase Name".PadRight(24) & "Mixture" & vbTab
                Case 1
                    line = "Phase Mole Fraction".PadRight(24) & vbTab
                Case Else
                    line = ""
            End Select
            For j = 0 To result2.GetLength(1) - 1
                If Double.TryParse(result2(i, j).ToString, New Double) Then
                    line += Double.Parse(result2(i, j).ToString).ToString("0.0000").PadRight(10)
                Else
                    line += result2(i, j).ToString.PadRight(10)
                End If
            Next
            Console.WriteLine(line)
        Next

        Console.Write(vbCrLf)
        Console.WriteLine("Press any key to continue...")
        Console.ReadKey(True)

    End Sub


End Module
