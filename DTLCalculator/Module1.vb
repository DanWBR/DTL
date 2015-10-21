'    DWSIM Thermodynamics Library Calculator Executable
'    Copyright 2015 Daniel Wagner O. de Medeiros
'
'    This file is part of DTL/DWSIM.
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

Imports System.Reflection

Module Module1

    Sub Main()

        Console.WriteLine("DTL - DWSIM Thermodynamics Library")
        Console.WriteLine("Copyright 2015 Daniel Wagner O. de Medeiros")
        Dim dt As DateTime = CType("01/01/2000", DateTime).AddDays(My.Application.Info.Version.Build).AddSeconds(My.Application.Info.Version.Revision * 2)
        Console.WriteLine("Version " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & _
        ", Build " & My.Application.Info.Version.Build & " (" & Format(dt, "dd/MM/yyyy HH:mm") & ")")
        If Type.GetType("Mono.Runtime") Is Nothing Then
            Console.WriteLine("Microsoft .NET Framework Runtime Version " & System.Runtime.InteropServices.RuntimeEnvironment.GetSystemVersion.ToString())
        Else
            Dim displayName As MethodInfo = Type.GetType("Mono.Runtime").GetMethod("GetDisplayName", BindingFlags.NonPublic Or BindingFlags.[Static])
            If displayName IsNot Nothing Then
                Console.WriteLine("Mono " + displayName.Invoke(Nothing, Nothing) + " / " + System.Runtime.InteropServices.RuntimeEnvironment.GetSystemVersion.ToString())
            Else
                Console.WriteLine(System.Runtime.InteropServices.RuntimeEnvironment.GetSystemVersion.ToString())
            End If
        End If
        Console.WriteLine()

        Dim argcount As Integer = My.Application.CommandLineArgs.Count
        Dim args(My.Application.CommandLineArgs.Count - 1) As String
        My.Application.CommandLineArgs.CopyTo(args, 0)

        'DTLCalculator "input.xml" "output.xml" 

        Dim inputfile As String = ""
        Dim outputfile As String = ""
        Dim path1 As String = My.Computer.FileSystem.CurrentDirectory + IO.Path.DirectorySeparatorChar
   
        inputfile = args(0)
        outputfile = args(1)

        Console.WriteLine("Working directory: " & path1)
        Console.WriteLine("Input XML file: " & inputfile)
        Console.WriteLine("Output XML file: " & outputfile)
        Console.WriteLine()

        Console.WriteLine("Parsing input XML file...")
        Console.WriteLine()
        Console.ReadKey()

    End Sub

End Module
