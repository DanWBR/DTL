Imports System.IO
Imports Cudafy.Translator
Imports Cudafy
Imports Cudafy.Host

'    Shared Functions
'    Copyright 2008 Daniel Wagner O. de Medeiros
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

Namespace DTL

    <System.Serializable()> Public Class App

        Public Shared Sub WriteToConsole(text As String, minlevel As Integer)

            If My.MyApplication._DebugLevel >= minlevel Then Console.WriteLine(text)

        End Sub

        Public Shared Function GetLocalString(ByVal id As String) As String
            Return id
        End Function

        Public Shared Function GetPropertyName(ByVal PropID As String) As String
            Return PropID
        End Function

        Public Shared Function GetComponentName(ByVal UniqueName As String) As String
            Return UniqueName
        End Function

        Public Shared Function GetComponentType(ByRef comp As DTL.ClassesBasicasTermodinamica.ConstantProperties) As String
            If comp.IsHYPO Then
                Return GetLocalString("CompHypo")
            ElseIf comp.IsPF Then
                Return GetLocalString("CompPseudo")
            Else
                Return GetLocalString("CompNormal")
            End If
        End Function

        Public Shared Function IsRunningOnMono() As Boolean
            Return Not Type.GetType("Mono.Runtime") Is Nothing
        End Function

        Shared Sub InitComputeDevice(Optional ByVal CudafyTarget As Cudafy.eLanguage = eLanguage.OpenCL, Optional ByVal DeviceID As Integer = 0)

            If My.MyApplication.gpu Is Nothing Then

                'set target language

                CudafyTranslator.Language = CudafyTarget
  
                'get the gpu instance

                Dim gputype As eGPUType

                Select Case CudafyTarget
                    Case eLanguage.Cuda
                        gputype = eGPUType.Cuda
                    Case eLanguage.OpenCL
                        gputype = eGPUType.OpenCL
                    Case Else
                        gputype = eGPUType.Emulator
                End Select

                My.MyApplication.gpu = CudafyHost.GetDevice(gputype, DeviceID)

                'cudafy all classes that contain a gpu function

                If My.MyApplication.gpumod Is Nothing Then
                    Select Case CudafyTarget
                        Case eLanguage.Cuda
                            My.MyApplication.gpumod = CudafyModule.TryDeserialize("cudacode.cdfy")
                        Case eLanguage.OpenCL
                            'My.MyApplication.gpumod = CudafyModule.TryDeserialize("openclcode.cdfy")
                    End Select
                    If My.MyApplication.gpumod Is Nothing OrElse Not My.MyApplication.gpumod.TryVerifyChecksums() Then
                        Select Case CudafyTarget
                            Case eLanguage.Cuda
                                Dim cp As New Cudafy.CompileProperties()
                                With cp
                                    .Architecture = eArchitecture.sm_20
                                    .CompileMode = eCudafyCompileMode.Default
                                    .Platform = ePlatform.x86
                                    .WorkingDirectory = My.Computer.FileSystem.SpecialDirectories.Temp
                                    'CUDA SDK v5.0 path
                                    .CompilerPath = "C:\Program Files\NVIDIA GPU Computing Toolkit\CUDA\v5.0\bin\nvcc.exe"
                                    .IncludeDirectoryPath = "C:\Program Files\NVIDIA GPU Computing Toolkit\CUDA\v5.0\include"
                                End With
                                My.MyApplication.gpumod = CudafyTranslator.Cudafy(cp, GetType(DTL.SimulationObjects.PropertyPackages.Auxiliary.LeeKeslerPlocker), _
                                            GetType(DTL.SimulationObjects.PropertyPackages.ThermoPlugs.PR))
                                My.MyApplication.gpumod.Serialize("cudacode.cdfy")
                            Case eLanguage.OpenCL
                                My.MyApplication.gpumod = CudafyTranslator.Cudafy(GetType(DTL.SimulationObjects.PropertyPackages.Auxiliary.LeeKeslerPlocker), _
                                           GetType(DTL.SimulationObjects.PropertyPackages.ThermoPlugs.PR))
                                'My.MyApplication.gpumod.Serialize("openclcode.cdfy")
                        End Select
                    End If
                End If

                'load cudafy module

                If Not My.MyApplication.gpu.IsModuleLoaded(My.MyApplication.gpumod.Name) Then My.MyApplication.gpu.LoadModule(My.MyApplication.gpumod)

            End If

        End Sub


    End Class

End Namespace
