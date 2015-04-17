﻿'    Stream Classes
'    Copyright 2008-2011 Daniel Wagner O. de Medeiros
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

Imports DTL.DTL.ClassesBasicasTermodinamica
Imports CapeOpen = CAPEOPEN110
Imports DTL.DTL.SimulationObjects.PropertyPackages
Imports System.Runtime.InteropServices
Imports CAPEOPEN110

Namespace DTL.SimulationObjects.Streams

    <System.Serializable()> Friend Class MaterialStream

        Inherits SimulationObjects_BaseClass

        'CAPE-OPEN 1.0
        Implements ICapeIdentification, ICapeThermoMaterialObject, ICapeThermoCalculationRoutine, ICapeThermoEquilibriumServer, ICapeThermoPropertyPackage

        'CAPE-OPEN 1.1
        Implements ICapeThermoMaterial, ICapeThermoCompounds, ICapeThermoPhases, ICapeThermoUniversalConstant, ICapeThermoPropertyRoutine, ICapeThermoEquilibriumRoutine, ICapeThermoMaterialContext

        'CAPE-OPEN Error Interfaces
        Implements ECapeUser, ECapeUnknown, ECapeRoot

        Friend _pp As DTL.SimulationObjects.PropertyPackages.PropertyPackage
        Private _ppid As String = ""

        Protected m_compositionbasis As CompBasis = CompBasis.Molar_Fractions
        Protected m_Phases As New Dictionary(Of String, DTL.ClassesBasicasTermodinamica.Fase)

        Private _inequilibrium As Boolean = False

#Region "    DWSIM Specific"

        Public Enum Flashspec
            Temperature_and_Pressure = 0
            Pressure_and_Enthalpy = 1
            Pressure_and_Entropy = 2
            Pressure_and_VaporFraction = 3
            Temperature_and_VaporFraction = 4
        End Enum

        Private _flashspec As Flashspec = Flashspec.Temperature_and_Pressure

        Public Property SpecType() As Flashspec
            Get
                Return _flashspec
            End Get
            Set(ByVal value As Flashspec)
                _flashspec = value
            End Set
        End Property

        Public Property AtEquilibrium() As Boolean
            Get
                Return _inequilibrium
            End Get
            Set(ByVal value As Boolean)
                _inequilibrium = value
            End Set
        End Property

        Public Property PropertyPackage() As DTL.SimulationObjects.PropertyPackages.PropertyPackage
            Get
                Return _pp
            End Get
            Set(ByVal value As DTL.SimulationObjects.PropertyPackages.PropertyPackage)
                If value IsNot Nothing Then _ppid = value.UniqueID
            End Set
        End Property

        Public Enum CompBasis
            Molar_Fractions
            Mass_Fractions
            Volumetric_Fractions
            Molar_Flows
            Mass_Flows
            Volumetric_Flows
        End Enum

        Public Property CompositionBasis() As CompBasis
            Get
                Return m_compositionbasis
            End Get
            Set(ByVal value As CompBasis)
                m_compositionbasis = value
            End Set
        End Property

        Public Sub New(ByVal nome As String, ByVal descricao As String)

            MyBase.CreateNew()

            Me.m_ComponentName = nome
            Me.m_ComponentDescription = descricao

            Me.Fases.Add("0", New DTL.ClassesBasicasTermodinamica.Fase(DTL.App.GetLocalString("Mistura"), ""))
            Me.Fases.Add("1", New DTL.ClassesBasicasTermodinamica.Fase(DTL.App.GetLocalString("OverallLiquid"), ""))
            Me.Fases.Add("2", New DTL.ClassesBasicasTermodinamica.Fase(DTL.App.GetLocalString("Vapor"), ""))
            Me.Fases.Add("3", New DTL.ClassesBasicasTermodinamica.Fase(DTL.App.GetLocalString("Liquid1"), ""))
            Me.Fases.Add("4", New DTL.ClassesBasicasTermodinamica.Fase(DTL.App.GetLocalString("Liquid2"), ""))
            Me.Fases.Add("5", New DTL.ClassesBasicasTermodinamica.Fase(DTL.App.GetLocalString("Liquid3"), ""))
            Me.Fases.Add("6", New DTL.ClassesBasicasTermodinamica.Fase(DTL.App.GetLocalString("Aqueous"), ""))
            Me.Fases.Add("7", New DTL.ClassesBasicasTermodinamica.Fase(DTL.App.GetLocalString("Solid"), ""))

        End Sub

        Public ReadOnly Property Fases() As Dictionary(Of String, DTL.ClassesBasicasTermodinamica.Fase)
            Get
                Return m_Phases
            End Get
        End Property

        Public Sub Assign(ByVal ASource As MaterialStream)

            Me.AtEquilibrium = ASource.AtEquilibrium

            'Copy properties from the ASource stream.

            Dim i As Integer

            For i = 0 To 7

                If ASource.Fases.ContainsKey(i) Then

                    Fases(i).SPMProperties.temperature = ASource.Fases(i).SPMProperties.temperature
                    Fases(i).SPMProperties.pressure = ASource.Fases(i).SPMProperties.pressure
                    Fases(i).SPMProperties.enthalpy = ASource.Fases(i).SPMProperties.enthalpy

                    'Copy component properties.
                    Dim comp As DTL.ClassesBasicasTermodinamica.Substancia

                    For Each comp In Fases(i).Componentes.Values
                        comp.FracaoMolar = ASource.Fases(i).Componentes(comp.Nome).FracaoMolar
                        comp.FracaoMassica = ASource.Fases(i).Componentes(comp.Nome).FracaoMassica
                    Next

                    'Should be defined after concentrations?!?! [yes, no, maybe... whatever]
                    Fases(i).SPMProperties.massflow = ASource.Fases(i).SPMProperties.massflow
                    Fases(i).SPMProperties.molarflow = ASource.Fases(i).SPMProperties.molarflow

                    Fases(i).SPMProperties.massfraction = ASource.Fases(i).SPMProperties.massfraction
                    Fases(i).SPMProperties.molarfraction = ASource.Fases(i).SPMProperties.molarfraction

                End If

            Next

        End Sub

        Public Sub AssignProps(ByVal ASource As MaterialStream)

            'Copy properties from the ASource stream.

            Dim i As Integer

            For i = 0 To 7

                If ASource.Fases.ContainsKey(i) Then

                    Fases(i).SPMProperties.temperature = ASource.Fases(i).SPMProperties.temperature
                    Fases(i).SPMProperties.pressure = ASource.Fases(i).SPMProperties.pressure
                    Fases(i).SPMProperties.density = ASource.Fases(i).SPMProperties.density
                    Fases(i).SPMProperties.enthalpy = ASource.Fases(i).SPMProperties.enthalpy
                    Fases(i).SPMProperties.entropy = ASource.Fases(i).SPMProperties.entropy
                    Fases(i).SPMProperties.molar_enthalpy = ASource.Fases(i).SPMProperties.molar_enthalpy
                    Fases(i).SPMProperties.molar_entropy = ASource.Fases(i).SPMProperties.molar_entropy
                    Fases(i).SPMProperties.compressibilityFactor = ASource.Fases(i).SPMProperties.compressibilityFactor
                    Fases(i).SPMProperties.heatCapacityCp = ASource.Fases(i).SPMProperties.heatCapacityCp
                    Fases(i).SPMProperties.heatCapacityCv = ASource.Fases(i).SPMProperties.heatCapacityCv
                    Fases(i).SPMProperties.molecularWeight = ASource.Fases(i).SPMProperties.molecularWeight
                    Fases(i).SPMProperties.thermalConductivity = ASource.Fases(i).SPMProperties.thermalConductivity
                    Fases(i).SPMProperties.speedOfSound = ASource.Fases(i).SPMProperties.speedOfSound
                    Fases(i).SPMProperties.volumetric_flow = ASource.Fases(i).SPMProperties.volumetric_flow
                    Fases(i).SPMProperties.jouleThomsonCoefficient = ASource.Fases(i).SPMProperties.jouleThomsonCoefficient
                    Fases(i).SPMProperties.excessEnthalpy = ASource.Fases(i).SPMProperties.excessEnthalpy
                    Fases(i).SPMProperties.excessEntropy = ASource.Fases(i).SPMProperties.excessEntropy
                    Fases(i).SPMProperties.compressibility = ASource.Fases(i).SPMProperties.compressibility
                    Fases(i).SPMProperties.bubbleTemperature = ASource.Fases(i).SPMProperties.bubbleTemperature
                    Fases(i).SPMProperties.bubblePressure = ASource.Fases(i).SPMProperties.bubblePressure
                    Fases(i).SPMProperties.dewTemperature = ASource.Fases(i).SPMProperties.dewTemperature
                    Fases(i).SPMProperties.dewPressure = ASource.Fases(i).SPMProperties.dewPressure
                    Fases(i).SPMProperties.viscosity = ASource.Fases(i).SPMProperties.viscosity
                    Fases(i).SPMProperties.kinematic_viscosity = ASource.Fases(i).SPMProperties.kinematic_viscosity
                    Fases(i).SPMProperties.molarflow = ASource.Fases(i).SPMProperties.molarflow
                    Fases(i).SPMProperties.massflow = ASource.Fases(i).SPMProperties.massflow
                    Fases(i).SPMProperties.massfraction = ASource.Fases(i).SPMProperties.massfraction
                    Fases(i).SPMProperties.molarfraction = ASource.Fases(i).SPMProperties.molarfraction

                End If

            Next

        End Sub

        Public Sub Clear()

            Dim i As Integer

            For i = 0 To Fases.Count - 1

                Fases(i).SPMProperties.temperature = Nothing
                Fases(i).SPMProperties.pressure = Nothing
                Fases(i).SPMProperties.enthalpy = Nothing
                Fases(i).SPMProperties.molarfraction = Nothing
                Fases(i).SPMProperties.massfraction = Nothing

                'Copy component properties.
                Dim comp As DTL.ClassesBasicasTermodinamica.Substancia

                For Each comp In Fases(i).Componentes.Values
                    comp.FracaoMolar = Nothing
                    comp.FracaoMassica = Nothing
                Next

                'Should be define after concentrations?!?!
                Fases(i).SPMProperties.massflow = Nothing
                Fases(i).SPMProperties.molarflow = Nothing

            Next

        End Sub

        Public Sub SetOverallComposition(ByVal Vx As Array)

            Dim i As Integer = 0
            For Each c As Substancia In Me.Fases(0).Componentes.Values
                c.FracaoMolar = Vx(i)
                i += 1
            Next

        End Sub

        Public Sub SetPhaseComposition(ByVal Vx As Array, ByVal phase As PropertyPackages.Fase)

            Dim i As Integer = 0, idx As Integer = 0
            Select Case phase
                Case PropertyPackages.Fase.Aqueous
                    idx = 2
                Case PropertyPackages.Fase.Liquid
                    idx = 1
                Case PropertyPackages.Fase.Liquid1
                    idx = 3
                Case PropertyPackages.Fase.Liquid2
                    idx = 4
                Case PropertyPackages.Fase.Liquid3
                    idx = 5
                Case PropertyPackages.Fase.Mixture
                    idx = 0
                Case PropertyPackages.Fase.Solid
                    idx = 7
                Case PropertyPackages.Fase.Vapor
                    idx = 2
            End Select
            For Each c As Substancia In Me.Fases(idx).Componentes.Values
                c.FracaoMolar = Vx(i)
                i += 1
            Next

        End Sub

        Public Overrides Function GetPropertyValue(ByVal prop As String, Optional ByVal su As SistemasDeUnidades.Unidades = Nothing) As Object

            If su Is Nothing Then su = New DTL.SistemasDeUnidades.UnidadesSI
            Dim cv As New DTL.SistemasDeUnidades.Conversor
            Dim value As String = ""
            Dim sname As String = ""
            Dim propidx As Integer = CInt(prop.Split(",")(0).Split("_")(2))
            If prop.Split(",").Length = 2 Then
                sname = prop.Split(",")(1)
            End If

            Select Case propidx

                Case 0
                    'PROP_MS_0 Temperature
                    value = cv.ConverterDoSI(su.spmp_temperature, Me.Fases(0).SPMProperties.temperature.GetValueOrDefault)
                Case 1
                    'PROP_MS_1 Pressure
                    value = cv.ConverterDoSI(su.spmp_pressure, Me.Fases(0).SPMProperties.pressure.GetValueOrDefault)
                Case 2
                    'PROP_MS_2	Mass Flow
                    value = cv.ConverterDoSI(su.spmp_massflow, Me.Fases(0).SPMProperties.massflow.GetValueOrDefault)
                Case 3
                    'PROP_MS_3	Molar Flow
                    value = cv.ConverterDoSI(su.spmp_molarflow, Me.Fases(0).SPMProperties.molarflow.GetValueOrDefault)
                Case 4
                    'PROP_MS_4	Volumetric Flow
                    value = cv.ConverterDoSI(su.spmp_volumetricFlow, Me.Fases(0).SPMProperties.volumetric_flow.GetValueOrDefault)
                Case 5
                    'PROP_MS_5	Mixture Density
                    value = cv.ConverterDoSI(su.spmp_density, Me.Fases(0).SPMProperties.density.GetValueOrDefault)
                Case 6
                    'PROP_MS_6	Mixture Molar Weight
                    value = cv.ConverterDoSI(su.spmp_molecularWeight, Me.Fases(0).SPMProperties.molecularWeight.GetValueOrDefault)
                Case 7
                    'PROP_MS_7	Mixture Specific Enthalpy
                    value = cv.ConverterDoSI(su.spmp_enthalpy, Me.Fases(0).SPMProperties.enthalpy.GetValueOrDefault)
                Case 8
                    'PROP_MS_8	Mixture Specific Entropy
                    value = cv.ConverterDoSI(su.spmp_entropy, Me.Fases(0).SPMProperties.entropy.GetValueOrDefault)
                Case 9
                    'PROP_MS_9	Mixture Molar Enthalpy
                    value = cv.ConverterDoSI(su.molar_enthalpy, Me.Fases(0).SPMProperties.molar_enthalpy.GetValueOrDefault)
                Case 10
                    'PROP_MS_10	Mixture Molar Entropy
                    value = cv.ConverterDoSI(su.molar_entropy, Me.Fases(0).SPMProperties.molar_entropy.GetValueOrDefault)
                Case 11
                    'PROP_MS_11	Mixture Thermal Conductivity
                    value = cv.ConverterDoSI(su.spmp_thermalConductivity, Me.Fases(0).SPMProperties.thermalConductivity.GetValueOrDefault)
                Case 12
                    'PROP_MS_12	Vapor Phase Density
                    value = cv.ConverterDoSI(su.spmp_density, Me.Fases(2).SPMProperties.density.GetValueOrDefault)
                Case 13
                    'PROP_MS_13	Vapor Phase Molar Weight
                    value = cv.ConverterDoSI(su.spmp_molecularWeight, Me.Fases(2).SPMProperties.molecularWeight.GetValueOrDefault)
                Case 14
                    'PROP_MS_14	Vapor Phase Specific Enthalpy
                    value = cv.ConverterDoSI(su.spmp_enthalpy, Me.Fases(2).SPMProperties.enthalpy.GetValueOrDefault)
                Case 15
                    'PROP_MS_15	Vapor Phase Specific Entropy
                    value = cv.ConverterDoSI(su.spmp_entropy, Me.Fases(2).SPMProperties.entropy.GetValueOrDefault)
                Case 16
                    'PROP_MS_16	Vapor Phase Molar Enthalpy
                    value = cv.ConverterDoSI(su.molar_enthalpy, Me.Fases(2).SPMProperties.molar_enthalpy.GetValueOrDefault)
                Case 17
                    'PROP_MS_17	Vapor Phase Molar Entropy
                    value = cv.ConverterDoSI(su.molar_entropy, Me.Fases(2).SPMProperties.molar_entropy.GetValueOrDefault)
                Case 18
                    'PROP_MS_18	Vapor Phase Thermal Conductivity
                    value = cv.ConverterDoSI(su.spmp_thermalConductivity, Me.Fases(2).SPMProperties.thermalConductivity.GetValueOrDefault)
                Case 19
                    'PROP_MS_19	Vapor Phase Kinematic Viscosity
                    value = cv.ConverterDoSI(su.spmp_cinematic_viscosity, Me.Fases(2).SPMProperties.kinematic_viscosity.GetValueOrDefault)
                Case 20
                    'PROP_MS_20	Vapor Phase Dynamic Viscosity
                    value = cv.ConverterDoSI(su.spmp_viscosity, Me.Fases(2).SPMProperties.viscosity.GetValueOrDefault)
                Case 21
                    'PROP_MS_21	Vapor Phase Heat Capacity (Cp)
                    value = cv.ConverterDoSI(su.spmp_heatCapacityCp, Me.Fases(2).SPMProperties.heatCapacityCp.GetValueOrDefault)
                Case 22
                    'PROP_MS_22	Vapor Phase Heat Capacity Ratio (Cp/Cv)
                    value = Me.Fases(2).SPMProperties.heatCapacityCp.GetValueOrDefault / Me.Fases(2).SPMProperties.heatCapacityCv.GetValueOrDefault
                Case 23
                    'PROP_MS_23	Vapor Phase Mass Flow
                    value = cv.ConverterDoSI(su.spmp_massflow, Me.Fases(2).SPMProperties.massflow.GetValueOrDefault)
                Case 24
                    'PROP_MS_24	Vapor Phase Molar Flow
                    value = cv.ConverterDoSI(su.spmp_molarflow, Me.Fases(2).SPMProperties.molarflow.GetValueOrDefault)
                Case 25
                    'PROP_MS_25	Vapor Phase Volumetric Flow
                    value = cv.ConverterDoSI(su.spmp_volumetricFlow, Me.Fases(2).SPMProperties.volumetric_flow.GetValueOrDefault)
                Case 26
                    'PROP_MS_26	Vapor Phase Compressibility Factor
                    value = Me.Fases(2).SPMProperties.compressibilityFactor.GetValueOrDefault
                Case 27
                    'PROP_MS_27	Vapor Phase Molar Fraction
                    value = Me.Fases(2).SPMProperties.molarfraction.GetValueOrDefault
                Case 28
                    'PROP_MS_28	Vapor Phase Mass Fraction
                    value = Me.Fases(2).SPMProperties.massfraction.GetValueOrDefault
                Case 29
                    'PROP_MS_29	Vapor Phase Volumetric Fraction
                    value = Me.Fases(2).SPMProperties.volumetric_flow.GetValueOrDefault / Me.Fases(0).SPMProperties.volumetric_flow.GetValueOrDefault
                Case 30
                    'PROP_MS_30	Liquid Phase (Mixture) Density
                    value = cv.ConverterDoSI(su.spmp_density, Me.Fases(1).SPMProperties.density.GetValueOrDefault)
                Case 31
                    'PROP_MS_31	Liquid Phase (Mixture) Molar Weight
                    value = cv.ConverterDoSI(su.spmp_molecularWeight, Me.Fases(1).SPMProperties.molecularWeight.GetValueOrDefault)
                Case 32
                    'PROP_MS_32	Liquid Phase (Mixture) Specific Enthalpy
                    value = cv.ConverterDoSI(su.spmp_enthalpy, Me.Fases(1).SPMProperties.enthalpy.GetValueOrDefault)
                Case 33
                    'PROP_MS_33	Liquid Phase (Mixture) Specific Entropy
                    value = cv.ConverterDoSI(su.spmp_entropy, Me.Fases(1).SPMProperties.entropy.GetValueOrDefault)
                Case 34
                    'PROP_MS_34	Liquid Phase (Mixture) Molar Enthalpy
                    value = cv.ConverterDoSI(su.molar_enthalpy, Me.Fases(1).SPMProperties.molar_enthalpy.GetValueOrDefault)
                Case 35
                    'PROP_MS_35	Liquid Phase (Mixture) Molar Entropy
                    value = cv.ConverterDoSI(su.molar_entropy, Me.Fases(1).SPMProperties.molar_entropy.GetValueOrDefault)
                Case 36
                    'PROP_MS_36	Liquid Phase (Mixture) Thermal Conductivity
                    value = cv.ConverterDoSI(su.spmp_thermalConductivity, Me.Fases(1).SPMProperties.thermalConductivity.GetValueOrDefault)
                Case 37
                    'PROP_MS_37	Liquid Phase (Mixture) Kinematic Viscosity
                    value = cv.ConverterDoSI(su.spmp_cinematic_viscosity, Me.Fases(1).SPMProperties.kinematic_viscosity.GetValueOrDefault)
                Case 38
                    'PROP_MS_38	Liquid Phase (Mixture) Dynamic Viscosity
                    value = cv.ConverterDoSI(su.spmp_viscosity, Me.Fases(1).SPMProperties.viscosity.GetValueOrDefault)
                Case 39
                    'PROP_MS_39	Liquid Phase (Mixture) Heat Capacity (Cp)
                    value = cv.ConverterDoSI(su.spmp_heatCapacityCp, Me.Fases(1).SPMProperties.heatCapacityCp.GetValueOrDefault)
                Case 40
                    'PROP_MS_40	Liquid Phase (Mixture) Heat Capacity Ratio (Cp/Cv)
                    value = Me.Fases(1).SPMProperties.heatCapacityCp.GetValueOrDefault / Me.Fases(1).SPMProperties.heatCapacityCv.GetValueOrDefault
                Case 41
                    'PROP_MS_41	Liquid Phase (Mixture) Mass Flow
                    value = cv.ConverterDoSI(su.spmp_massflow, Me.Fases(1).SPMProperties.massflow.GetValueOrDefault)
                Case 42
                    'PROP_MS_42	Liquid Phase (Mixture) Molar Flow
                    value = cv.ConverterDoSI(su.spmp_molarflow, Me.Fases(1).SPMProperties.molarflow.GetValueOrDefault)
                Case 43
                    'PROP_MS_43	Liquid Phase (Mixture) Volumetric Flow
                    value = cv.ConverterDoSI(su.spmp_volumetricFlow, Me.Fases(1).SPMProperties.volumetric_flow.GetValueOrDefault)
                Case 44
                    'PROP_MS_44	Liquid Phase (Mixture) Compressibility Factor
                    value = Me.Fases(1).SPMProperties.compressibilityFactor.GetValueOrDefault
                Case 45
                    'PROP_MS_45	Liquid Phase (Mixture) Molar Fraction
                    value = Me.Fases(1).SPMProperties.molarfraction.GetValueOrDefault
                Case 46
                    'PROP_MS_46	Liquid Phase (Mixture) Mass Fraction
                    value = Me.Fases(1).SPMProperties.massfraction.GetValueOrDefault
                Case 47
                    'PROP_MS_47	Liquid Phase (Mixture) Volumetric Fraction
                    value = Me.Fases(1).SPMProperties.volumetric_flow.GetValueOrDefault / Me.Fases(0).SPMProperties.volumetric_flow.GetValueOrDefault
                Case 48
                    'PROP_MS_48	Liquid Phase (1) Density
                    value = cv.ConverterDoSI(su.spmp_density, Me.Fases(3).SPMProperties.density.GetValueOrDefault)
                Case 49
                    'PROP_MS_49	Liquid Phase (1) Molar Weight
                    value = cv.ConverterDoSI(su.spmp_molecularWeight, Me.Fases(3).SPMProperties.molecularWeight.GetValueOrDefault)
                Case 50
                    'PROP_MS_50	Liquid Phase (1) Specific Enthalpy
                    value = cv.ConverterDoSI(su.spmp_enthalpy, Me.Fases(3).SPMProperties.enthalpy.GetValueOrDefault)
                Case 51
                    'PROP_MS_51	Liquid Phase (1) Specific Entropy
                    value = cv.ConverterDoSI(su.spmp_entropy, Me.Fases(3).SPMProperties.entropy.GetValueOrDefault)
                Case 52
                    'PROP_MS_52	Liquid Phase (1) Molar Enthalpy
                    value = cv.ConverterDoSI(su.molar_enthalpy, Me.Fases(3).SPMProperties.molar_enthalpy.GetValueOrDefault)
                Case 53
                    'PROP_MS_53	Liquid Phase (1) Molar Entropy
                    value = cv.ConverterDoSI(su.molar_entropy, Me.Fases(3).SPMProperties.molar_entropy.GetValueOrDefault)
                Case 54
                    'PROP_MS_54	Liquid Phase (1) Thermal Conductivity
                    value = cv.ConverterDoSI(su.spmp_thermalConductivity, Me.Fases(3).SPMProperties.thermalConductivity.GetValueOrDefault)
                Case 55
                    'PROP_MS_55	Liquid Phase (1) Kinematic Viscosity
                    value = cv.ConverterDoSI(su.spmp_cinematic_viscosity, Me.Fases(3).SPMProperties.kinematic_viscosity.GetValueOrDefault)
                Case 56
                    'PROP_MS_56	Liquid Phase (1) Dynamic Viscosity
                    value = cv.ConverterDoSI(su.spmp_viscosity, Me.Fases(3).SPMProperties.viscosity.GetValueOrDefault)
                Case 57
                    'PROP_MS_57	Liquid Phase (1) Heat Capacity (Cp)
                    value = cv.ConverterDoSI(su.spmp_heatCapacityCp, Me.Fases(3).SPMProperties.heatCapacityCp.GetValueOrDefault)
                Case 58
                    'PROP_MS_58	Liquid Phase (1) Heat Capacity Ratio (Cp/Cv)
                    value = Me.Fases(3).SPMProperties.heatCapacityCp.GetValueOrDefault / Me.Fases(3).SPMProperties.heatCapacityCv.GetValueOrDefault
                Case 59
                    'PROP_MS_59	Liquid Phase (1) Mass Flow
                    value = cv.ConverterDoSI(su.spmp_massflow, Me.Fases(3).SPMProperties.massflow.GetValueOrDefault)
                Case 60
                    'PROP_MS_60	Liquid Phase (1) Molar Flow
                    value = cv.ConverterDoSI(su.spmp_molarflow, Me.Fases(3).SPMProperties.molarflow.GetValueOrDefault)
                Case 61
                    'PROP_MS_61	Liquid Phase (1) Volumetric Flow
                    value = cv.ConverterDoSI(su.spmp_volumetricFlow, Me.Fases(3).SPMProperties.volumetric_flow.GetValueOrDefault)
                Case 62
                    'PROP_MS_62	Liquid Phase (1) Compressibility Factor
                    value = Me.Fases(3).SPMProperties.compressibilityFactor.GetValueOrDefault
                Case 63
                    'PROP_MS_63	Liquid Phase (1) Molar Fraction
                    value = Me.Fases(3).SPMProperties.molarfraction.GetValueOrDefault
                Case 64
                    'PROP_MS_64	Liquid Phase (1) Mass Fraction
                    value = Me.Fases(3).SPMProperties.massfraction.GetValueOrDefault
                Case 65
                    'PROP_MS_65	Liquid Phase (1) Volumetric Fraction
                    value = Me.Fases(3).SPMProperties.volumetric_flow.GetValueOrDefault / Me.Fases(0).SPMProperties.volumetric_flow.GetValueOrDefault
                Case 66
                    'PROP_MS_66	Liquid Phase (2) Density
                    value = cv.ConverterDoSI(su.spmp_density, Me.Fases(4).SPMProperties.density.GetValueOrDefault)
                Case 67
                    'PROP_MS_67	Liquid Phase (2) Molar Weight
                    value = cv.ConverterDoSI(su.spmp_molecularWeight, Me.Fases(4).SPMProperties.molecularWeight.GetValueOrDefault)
                Case 68
                    'PROP_MS_68	Liquid Phase (2) Specific Enthalpy
                    value = cv.ConverterDoSI(su.spmp_enthalpy, Me.Fases(4).SPMProperties.enthalpy.GetValueOrDefault)
                Case 69
                    'PROP_MS_69	Liquid Phase (2) Specific Entropy
                    value = cv.ConverterDoSI(su.spmp_entropy, Me.Fases(4).SPMProperties.entropy.GetValueOrDefault)
                Case 70
                    'PROP_MS_70	Liquid Phase (2) Molar Enthalpy
                    value = cv.ConverterDoSI(su.molar_enthalpy, Me.Fases(4).SPMProperties.molar_enthalpy.GetValueOrDefault)
                Case 71
                    'PROP_MS_71	Liquid Phase (2) Molar Entropy
                    value = cv.ConverterDoSI(su.molar_entropy, Me.Fases(4).SPMProperties.molar_entropy.GetValueOrDefault)
                Case 72
                    'PROP_MS_72	Liquid Phase (2) Thermal Conductivity
                    value = cv.ConverterDoSI(su.spmp_thermalConductivity, Me.Fases(4).SPMProperties.thermalConductivity.GetValueOrDefault)
                Case 73
                    'PROP_MS_73	Liquid Phase (2) Kinematic Viscosity
                    value = cv.ConverterDoSI(su.spmp_cinematic_viscosity, Me.Fases(4).SPMProperties.kinematic_viscosity.GetValueOrDefault)
                Case 74
                    'PROP_MS_74	Liquid Phase (2) Dynamic Viscosity
                    value = cv.ConverterDoSI(su.spmp_viscosity, Me.Fases(4).SPMProperties.viscosity.GetValueOrDefault)
                Case 75
                    'PROP_MS_75	Liquid Phase (2) Heat Capacity (Cp)
                    value = cv.ConverterDoSI(su.spmp_heatCapacityCp, Me.Fases(4).SPMProperties.heatCapacityCp.GetValueOrDefault)
                Case 76
                    'PROP_MS_76	Liquid Phase (2) Heat Capacity Ratio (Cp/Cv)
                    value = Me.Fases(4).SPMProperties.heatCapacityCp.GetValueOrDefault / Me.Fases(4).SPMProperties.heatCapacityCv.GetValueOrDefault
                Case 77
                    'PROP_MS_77	Liquid Phase (2) Mass Flow
                    value = cv.ConverterDoSI(su.spmp_massflow, Me.Fases(4).SPMProperties.massflow.GetValueOrDefault)
                Case 78
                    'PROP_MS_78	Liquid Phase (2) Molar Flow
                    value = cv.ConverterDoSI(su.spmp_molarflow, Me.Fases(4).SPMProperties.molarflow.GetValueOrDefault)
                Case 79
                    'PROP_MS_79	Liquid Phase (2) Volumetric Flow
                    value = cv.ConverterDoSI(su.spmp_volumetricFlow, Me.Fases(4).SPMProperties.volumetric_flow.GetValueOrDefault)
                Case 80
                    'PROP_MS_80	Liquid Phase (2) Compressibility Factor
                    value = Me.Fases(4).SPMProperties.compressibilityFactor.GetValueOrDefault
                Case 81
                    'PROP_MS_81	Liquid Phase (2) Molar Fraction
                    value = Me.Fases(4).SPMProperties.molarfraction.GetValueOrDefault
                Case 82
                    'PROP_MS_82	Liquid Phase (2) Mass Fraction
                    value = Me.Fases(4).SPMProperties.massfraction.GetValueOrDefault
                Case 83
                    'PROP_MS_83	Liquid Phase (2) Volumetric Fraction
                    value = Me.Fases(4).SPMProperties.volumetric_flow.GetValueOrDefault / Me.Fases(0).SPMProperties.volumetric_flow.GetValueOrDefault
                Case 84
                    'PROP_MS_84	Aqueous Phase Density
                    If Me.Fases.ContainsKey(6) Then
                        value = cv.ConverterDoSI(su.spmp_density, Me.Fases(6).SPMProperties.density.GetValueOrDefault)
                    Else
                        value = 0
                    End If
                Case 85
                    'PROP_MS_85	Aqueous Phase Molar Weight
                    If Me.Fases.ContainsKey(6) Then
                        value = cv.ConverterDoSI(su.spmp_molecularWeight, Me.Fases(6).SPMProperties.molecularWeight.GetValueOrDefault)
                    Else
                        value = 0
                    End If
                Case 86
                    'PROP_MS_86	Aqueous Phase Specific Enthalpy
                    If Me.Fases.ContainsKey(6) Then
                        value = cv.ConverterDoSI(su.spmp_enthalpy, Me.Fases(6).SPMProperties.enthalpy.GetValueOrDefault)
                    Else
                        value = 0
                    End If
                Case 87
                    'PROP_MS_87	Aqueous Phase Specific Entropy
                    If Me.Fases.ContainsKey(6) Then
                        value = cv.ConverterDoSI(su.spmp_entropy, Me.Fases(6).SPMProperties.entropy.GetValueOrDefault)
                    Else
                        value = 0
                    End If
                Case 88
                    'PROP_MS_88	Aqueous Phase Molar Enthalpy
                    If Me.Fases.ContainsKey(6) Then
                        value = cv.ConverterDoSI(su.molar_enthalpy, Me.Fases(6).SPMProperties.molar_enthalpy.GetValueOrDefault)
                    Else
                        value = 0
                    End If
                Case 89
                    'PROP_MS_89	Aqueous Phase Molar Entropy
                    If Me.Fases.ContainsKey(6) Then
                        value = cv.ConverterDoSI(su.molar_entropy, Me.Fases(6).SPMProperties.molar_entropy.GetValueOrDefault)
                    Else
                        value = 0
                    End If
                Case 90
                    'PROP_MS_90	Aqueous Phase Thermal Conductivity
                    If Me.Fases.ContainsKey(6) Then
                        value = cv.ConverterDoSI(su.spmp_thermalConductivity, Me.Fases(6).SPMProperties.thermalConductivity.GetValueOrDefault)
                    Else
                        value = 0
                    End If
                Case 91
                    'PROP_MS_91	Aqueous Phase Kinematic Viscosity
                    If Me.Fases.ContainsKey(6) Then
                        value = cv.ConverterDoSI(su.spmp_cinematic_viscosity, Me.Fases(6).SPMProperties.kinematic_viscosity.GetValueOrDefault)
                    Else
                        value = 0
                    End If
                Case 92
                    'PROP_MS_92	Aqueous Phase Dynamic Viscosity
                    If Me.Fases.ContainsKey(6) Then
                        value = cv.ConverterDoSI(su.spmp_viscosity, Me.Fases(6).SPMProperties.viscosity.GetValueOrDefault)
                    Else
                        value = 0
                    End If
                Case 93
                    'PROP_MS_93	Aqueous Phase Heat Capacity (Cp)
                    If Me.Fases.ContainsKey(6) Then
                        value = cv.ConverterDoSI(su.spmp_heatCapacityCp, Me.Fases(6).SPMProperties.heatCapacityCp.GetValueOrDefault)
                    Else
                        value = 0
                    End If
                Case 94
                    'PROP_MS_94	Aqueous Phase Heat Capacity Ratio (Cp/Cv)
                    If Me.Fases.ContainsKey(6) Then
                        value = Me.Fases(6).SPMProperties.heatCapacityCp.GetValueOrDefault / Me.Fases(6).SPMProperties.heatCapacityCv.GetValueOrDefault
                    Else
                        value = 0
                    End If
                Case 95
                    'PROP_MS_95	Aqueous Phase Mass Flow
                    If Me.Fases.ContainsKey(6) Then
                        value = cv.ConverterDoSI(su.spmp_massflow, Me.Fases(6).SPMProperties.massflow.GetValueOrDefault)
                    Else
                        value = 0
                    End If
                Case 96
                    'PROP_MS_96	Aqueous Phase Molar Flow
                    If Me.Fases.ContainsKey(6) Then
                        value = cv.ConverterDoSI(su.spmp_molarflow, Me.Fases(6).SPMProperties.molarflow.GetValueOrDefault)
                    Else
                        value = 0
                    End If
                Case 97
                    'PROP_MS_97	Aqueous Phase Volumetric Flow
                    If Me.Fases.ContainsKey(6) Then
                        value = cv.ConverterDoSI(su.spmp_volumetricFlow, Me.Fases(6).SPMProperties.volumetric_flow.GetValueOrDefault)
                    Else
                        value = 0
                    End If
                Case 98
                    'PROP_MS_98	Aqueous Phase Compressibility Factor
                    If Me.Fases.ContainsKey(6) Then
                        value = Me.Fases(6).SPMProperties.compressibilityFactor.GetValueOrDefault
                    Else
                        value = 0
                    End If
                Case 99
                    'PROP_MS_99	Aqueous Phase Molar Fraction
                    If Me.Fases.ContainsKey(6) Then
                        value = Me.Fases(6).SPMProperties.molarfraction.GetValueOrDefault
                    Else
                        value = 0
                    End If
                Case 100
                    'PROP_MS_100	Aqueous Phase Mass Fraction
                    If Me.Fases.ContainsKey(6) Then
                        value = Me.Fases(6).SPMProperties.massfraction.GetValueOrDefault
                    Else
                        value = 0
                    End If
                Case 101
                    'PROP_MS_101	Aqueous Phase Volumetric Fraction
                    If Me.Fases.ContainsKey(6) Then
                        value = Me.Fases(6).SPMProperties.volumetric_flow.GetValueOrDefault / Me.Fases(0).SPMProperties.volumetric_flow.GetValueOrDefault
                    Else
                        value = 0
                    End If
                Case 103, 111, 112, 113, 114, 115
                    If Me.Fases(0).Componentes.ContainsKey(sname) Then
                        If propidx = 103 Then
                            value = Me.Fases(0).Componentes(sname).FracaoMassica.GetValueOrDefault
                        ElseIf propidx = 111 Then
                            value = Me.Fases(2).Componentes(sname).FracaoMassica.GetValueOrDefault
                        ElseIf propidx = 112 Then
                            value = Me.Fases(1).Componentes(sname).FracaoMassica.GetValueOrDefault
                        ElseIf propidx = 113 Then
                            value = Me.Fases(3).Componentes(sname).FracaoMassica.GetValueOrDefault
                        ElseIf propidx = 114 Then
                            value = Me.Fases(4).Componentes(sname).FracaoMassica.GetValueOrDefault
                        ElseIf propidx = 115 Then
                            value = Me.Fases(5).Componentes(sname).FracaoMassica.GetValueOrDefault
                        End If
                    Else
                        value = 0
                    End If
                Case 102, 106, 107, 108, 109, 110
                    If Me.Fases(0).Componentes.ContainsKey(sname) Then
                        If propidx = 102 Then
                            value = Me.Fases(0).Componentes(sname).FracaoMolar.GetValueOrDefault
                        ElseIf propidx = 106 Then
                            value = Me.Fases(2).Componentes(sname).FracaoMolar.GetValueOrDefault
                        ElseIf propidx = 107 Then
                            value = Me.Fases(1).Componentes(sname).FracaoMolar.GetValueOrDefault
                        ElseIf propidx = 108 Then
                            value = Me.Fases(3).Componentes(sname).FracaoMolar.GetValueOrDefault
                        ElseIf propidx = 109 Then
                            value = Me.Fases(4).Componentes(sname).FracaoMolar.GetValueOrDefault
                        ElseIf propidx = 110 Then
                            value = Me.Fases(5).Componentes(sname).FracaoMolar.GetValueOrDefault
                        End If
                    Else
                        value = 0
                    End If
                Case 104, 116, 117, 118, 119, 120
                    If Me.Fases(0).Componentes.ContainsKey(sname) Then
                        If propidx = 104 Then
                            value = Me.Fases(0).Componentes(sname).MolarFlow.GetValueOrDefault
                        ElseIf propidx = 116 Then
                            value = Me.Fases(2).Componentes(sname).MolarFlow.GetValueOrDefault
                        ElseIf propidx = 117 Then
                            value = Me.Fases(1).Componentes(sname).MolarFlow.GetValueOrDefault
                        ElseIf propidx = 118 Then
                            value = Me.Fases(3).Componentes(sname).MolarFlow.GetValueOrDefault
                        ElseIf propidx = 119 Then
                            value = Me.Fases(4).Componentes(sname).MolarFlow.GetValueOrDefault
                        ElseIf propidx = 120 Then
                            value = Me.Fases(5).Componentes(sname).MolarFlow.GetValueOrDefault
                        End If
                    Else
                        value = 0
                    End If
                Case 105, 121, 122, 123, 124, 125
                    If Me.Fases(0).Componentes.ContainsKey(sname) Then
                        If propidx = 105 Then
                            value = Me.Fases(0).Componentes(sname).MassFlow.GetValueOrDefault
                        ElseIf propidx = 121 Then
                            value = Me.Fases(2).Componentes(sname).MassFlow.GetValueOrDefault
                        ElseIf propidx = 122 Then
                            value = Me.Fases(1).Componentes(sname).MassFlow.GetValueOrDefault
                        ElseIf propidx = 123 Then
                            value = Me.Fases(3).Componentes(sname).MassFlow.GetValueOrDefault
                        ElseIf propidx = 124 Then
                            value = Me.Fases(4).Componentes(sname).MassFlow.GetValueOrDefault
                        ElseIf propidx = 125 Then
                            value = Me.Fases(5).Componentes(sname).MassFlow.GetValueOrDefault
                        End If
                    Else
                        value = 0
                    End If
                Case 126
                    value = cv.ConverterDoSI(su.spmp_pressure, Me.Fases(0).SPMProperties.bubblePressure.GetValueOrDefault)
                Case 127
                    value = cv.ConverterDoSI(su.spmp_pressure, Me.Fases(0).SPMProperties.dewPressure.GetValueOrDefault)
                Case 128
                    value = cv.ConverterDoSI(su.spmp_temperature, Me.Fases(0).SPMProperties.bubbleTemperature.GetValueOrDefault)
                Case 129
                    value = cv.ConverterDoSI(su.spmp_temperature, Me.Fases(0).SPMProperties.dewTemperature.GetValueOrDefault)
                Case 130
                    If Me.Fases(1).SPMProperties.molarfraction.GetValueOrDefault = 1.0# Then
                        value = "Liquid Only"
                    ElseIf Me.Fases(2).SPMProperties.molarfraction.GetValueOrDefault = 1.0# Then
                        value = "Vapor Only"
                    Else
                        value = "Mixed"
                    End If
            End Select

            Return value

        End Function

        Public Overloads Overrides Function GetProperties(ByVal proptype As SimulationObjects_BaseClass.PropertyType) As String()
            Dim i As Integer = 0
            Dim proplist As New ArrayList
            Return proplist.ToArray(GetType(System.String))
            proplist = Nothing
        End Function

        Public Overrides Function SetPropertyValue(ByVal prop As String, ByVal propval As Object, Optional ByVal su As DTL.SistemasDeUnidades.Unidades = Nothing) As Object

            If su Is Nothing Then su = New DTL.SistemasDeUnidades.UnidadesSI
            Dim cv As New DTL.SistemasDeUnidades.Conversor
            Dim propidx As Integer = CInt(prop.Split(",")(0).Split("_")(2))
            Dim sname As String = ""
            If prop.Split(",").Length = 2 Then
                sname = prop.Split(",")(1)
            End If

            Me.PropertyPackage.CurrentMaterialStream = Me

            Select Case propidx
                Case 0
                    'PROP_MS_0 Temperature
                    Me.Fases(0).SPMProperties.temperature = cv.ConverterParaSI(su.spmp_temperature, propval)
                Case 1
                    'PROP_MS_1 Pressure
                    Me.Fases(0).SPMProperties.pressure = cv.ConverterParaSI(su.spmp_pressure, propval)
                Case 2
                    'PROP_MS_2	Mass Flow
                    Me.Fases(0).SPMProperties.massflow = cv.ConverterParaSI(su.spmp_massflow, propval)
                    Me.PropertyPackage.DW_CalcVazaoMolar()
                    Me.PropertyPackage.DW_CalcVazaoVolumetrica()
                Case 3
                    'PROP_MS_3	Molar Flow
                    Me.Fases(0).SPMProperties.molarflow = cv.ConverterParaSI(su.spmp_molarflow, propval)
                    Me.PropertyPackage.DW_CalcVazaoMassica()
                    Me.PropertyPackage.DW_CalcVazaoVolumetrica()
                Case 4
                    'PROP_MS_4	Volumetric Flow
                    Me.Fases(0).SPMProperties.volumetric_flow = cv.ConverterParaSI(su.spmp_volumetricFlow, propval)
                    Me.Fases(0).SPMProperties.massflow = Me.Fases(0).SPMProperties.volumetric_flow * Me.Fases(0).SPMProperties.density.GetValueOrDefault
                    Me.PropertyPackage.DW_CalcVazaoMolar()
                Case 102
                    If Me.Fases(0).Componentes.ContainsKey(sname) Then
                        Me.Fases(0).Componentes(sname).FracaoMolar = propval
                        Dim sumfm As Double = 0
                        For Each comp As Substancia In Me.Fases(0).Componentes.Values
                            sumfm += comp.FracaoMolar
                        Next
                        For Each comp As Substancia In Me.Fases(0).Componentes.Values
                            comp.FracaoMolar /= sumfm
                        Next
                        Dim mtotal As Double = 0
                        Me.PropertyPackage.DW_CalcCompMolarFlow(0)
                        For Each comp As Substancia In Me.Fases(0).Componentes.Values
                            mtotal += comp.FracaoMolar.GetValueOrDefault * comp.ConstantProperties.Molar_Weight
                        Next
                        For Each comp As Substancia In Me.Fases(0).Componentes.Values
                            comp.FracaoMassica = comp.FracaoMolar.GetValueOrDefault * comp.ConstantProperties.Molar_Weight / mtotal
                        Next
                        Me.PropertyPackage.DW_CalcCompMassFlow(0)
                    End If
                Case 103
                    If Me.Fases(0).Componentes.ContainsKey(sname) Then
                        Me.Fases(0).Componentes(sname).FracaoMassica = propval
                        Dim sumfm As Double = 0
                        For Each comp As Substancia In Me.Fases(0).Componentes.Values
                            sumfm += comp.FracaoMassica
                        Next
                        For Each comp As Substancia In Me.Fases(0).Componentes.Values
                            comp.FracaoMassica /= sumfm
                        Next
                        Dim mtotal As Double = 0
                        Me.PropertyPackage.DW_CalcCompMassFlow(0)
                        For Each comp As Substancia In Me.Fases(0).Componentes.Values
                            mtotal += comp.FracaoMassica.GetValueOrDefault / comp.ConstantProperties.Molar_Weight
                        Next
                        For Each comp As Substancia In Me.Fases(0).Componentes.Values
                            comp.FracaoMolar = comp.FracaoMassica.GetValueOrDefault / comp.ConstantProperties.Molar_Weight / mtotal
                        Next
                        Me.PropertyPackage.DW_CalcCompMolarFlow(0)
                    End If
                Case 104
                    If Me.Fases(0).Componentes.ContainsKey(sname) Then
                        Me.Fases(0).Componentes(sname).MolarFlow = cv.ConverterParaSI(su.spmp_molarflow, propval)
                        Me.Fases(0).Componentes(sname).MassFlow = cv.ConverterParaSI(su.spmp_molarflow, propval) / 1000 * Me.Fases(0).Componentes(sname).ConstantProperties.Molar_Weight
                        Dim summ As Double = 0
                        For Each comp As Substancia In Me.Fases(0).Componentes.Values
                            summ += comp.MolarFlow
                        Next
                        Me.Fases(0).SPMProperties.molarflow = summ
                        For Each comp As Substancia In Me.Fases(0).Componentes.Values
                            comp.FracaoMolar = comp.MolarFlow / summ
                        Next
                        Dim mtotal As Double = 0
                        For Each comp As Substancia In Me.Fases(0).Componentes.Values
                            mtotal += comp.FracaoMolar.GetValueOrDefault * comp.ConstantProperties.Molar_Weight
                        Next
                        Me.Fases(0).SPMProperties.massflow = mtotal * summ / 1000
                        For Each comp As Substancia In Me.Fases(0).Componentes.Values
                            comp.FracaoMassica = comp.MolarFlow.GetValueOrDefault * Me.Fases(0).SPMProperties.massflow.GetValueOrDefault
                        Next
                    End If
                Case 105
                    If Me.Fases(0).Componentes.ContainsKey(sname) Then
                        Me.Fases(0).Componentes(sname).MassFlow = cv.ConverterParaSI(su.spmp_massflow, propval)
                        Me.Fases(0).Componentes(sname).MolarFlow = cv.ConverterParaSI(su.spmp_massflow, propval) / Me.Fases(0).Componentes(sname).ConstantProperties.Molar_Weight * 1000
                        Dim mtotal As Double = 0
                        For Each comp As Substancia In Me.Fases(0).Componentes.Values
                            mtotal += comp.MassFlow
                        Next
                        Me.Fases(0).SPMProperties.massflow = mtotal
                        For Each comp As Substancia In Me.Fases(0).Componentes.Values
                            comp.FracaoMassica = comp.MassFlow / mtotal
                        Next
                        Dim summ As Double = 0
                        For Each comp As Substancia In Me.Fases(0).Componentes.Values
                            summ += comp.FracaoMassica.GetValueOrDefault / comp.ConstantProperties.Molar_Weight / 1000
                        Next
                        Me.Fases(0).SPMProperties.molarflow = mtotal / summ
                        For Each comp As Substancia In Me.Fases(0).Componentes.Values
                            comp.FracaoMolar = comp.MolarFlow.GetValueOrDefault * Me.Fases(0).SPMProperties.molarflow.GetValueOrDefault
                        Next
                    End If
            End Select
            Return 1
        End Function

        Public Overrides Function GetPropertyUnit(ByVal prop As String, Optional ByVal su As SistemasDeUnidades.Unidades = Nothing) As Object

            If su Is Nothing Then su = New DTL.SistemasDeUnidades.UnidadesSI
            Dim value As String = ""
            Dim propidx As Integer = CInt(prop.Split(",")(0).Split("_")(2))

            Select Case propidx

                Case 0
                    'PROP_MS_0 Temperature
                    value = su.spmp_temperature
                Case 1
                    'PROP_MS_1 Pressure
                    value = su.spmp_pressure
                Case 2
                    'PROP_MS_2	Mass Flow
                    value = su.spmp_massflow
                Case 3
                    'PROP_MS_3	Molar Flow
                    value = su.spmp_molarflow
                Case 4
                    'PROP_MS_4	Volumetric Flow
                    value = su.spmp_volumetricFlow
                Case 5
                    'PROP_MS_5	Mixture Density
                    value = su.spmp_density
                Case 6
                    'PROP_MS_6	Mixture Molar Weight
                    value = su.spmp_molecularWeight
                Case 7
                    'PROP_MS_7	Mixture Specific Enthalpy
                    value = su.spmp_enthalpy
                Case 8
                    'PROP_MS_8	Mixture Specific Entropy
                    value = su.spmp_entropy
                Case 9
                    'PROP_MS_9	Mixture Molar Enthalpy
                    value = su.molar_enthalpy
                Case 10
                    'PROP_MS_10	Mixture Molar Entropy
                    value = su.molar_entropy
                Case 11
                    'PROP_MS_11	Mixture Thermal Conductivity
                    value = su.spmp_thermalConductivity
                Case 12
                    'PROP_MS_12	Vapor Phase Density
                    value = su.spmp_density
                Case 13
                    'PROP_MS_13	Vapor Phase Molar Weight
                    value = su.spmp_molecularWeight
                Case 14
                    'PROP_MS_14	Vapor Phase Specific Enthalpy
                    value = su.spmp_enthalpy
                Case 15
                    'PROP_MS_15	Vapor Phase Specific Entropy
                    value = su.spmp_entropy
                Case 16
                    'PROP_MS_16	Vapor Phase Molar Enthalpy
                    value = su.molar_enthalpy
                Case 17
                    'PROP_MS_17	Vapor Phase Molar Entropy
                    value = su.molar_entropy
                Case 18
                    'PROP_MS_18	Vapor Phase Thermal Conductivity
                    value = su.spmp_thermalConductivity
                Case 19
                    'PROP_MS_19	Vapor Phase Kinematic Viscosity
                    value = su.spmp_cinematic_viscosity
                Case 20
                    'PROP_MS_20	Vapor Phase Dynamic Viscosity
                    value = su.spmp_viscosity
                Case 21
                    'PROP_MS_21	Vapor Phase Heat Capacity (Cp)
                    value = su.spmp_heatCapacityCp
                Case 22
                    'PROP_MS_22	Vapor Phase Heat Capacity Ratio (Cp/Cv)
                    value = ""
                Case 23
                    'PROP_MS_23	Vapor Phase Mass Flow
                    value = su.spmp_massflow
                Case 24
                    'PROP_MS_24	Vapor Phase Molar Flow
                    value = su.spmp_molarflow
                Case 25
                    'PROP_MS_25	Vapor Phase Volumetric Flow
                    value = su.spmp_volumetricFlow
                Case 26
                    'PROP_MS_26	Vapor Phase Compressibility Factor
                    value = ""
                Case 27
                    'PROP_MS_27	Vapor Phase Molar Fraction
                    value = ""
                Case 28
                    'PROP_MS_28	Vapor Phase Mass Fraction
                    value = ""
                Case 29
                    'PROP_MS_29	Vapor Phase Volumetric Fraction
                    value = ""
                Case 30
                    'PROP_MS_30	Liquid Phase (Mixture) Density
                    value = su.spmp_density
                Case 31
                    'PROP_MS_31	Liquid Phase (Mixture) Molar Weight
                    value = su.spmp_molecularWeight
                Case 32
                    'PROP_MS_32	Liquid Phase (Mixture) Specific Enthalpy
                    value = su.spmp_enthalpy
                Case 33
                    'PROP_MS_33	Liquid Phase (Mixture) Specific Entropy
                    value = su.spmp_entropy
                Case 34
                    'PROP_MS_34	Liquid Phase (Mixture) Molar Enthalpy
                    value = su.molar_enthalpy
                Case 35
                    'PROP_MS_35	Liquid Phase (Mixture) Molar Entropy
                    value = su.molar_entropy
                Case 36
                    'PROP_MS_36	Liquid Phase (Mixture) Thermal Conductivity
                    value = su.spmp_thermalConductivity
                Case 37
                    'PROP_MS_37	Liquid Phase (Mixture) Kinematic Viscosity
                    value = su.spmp_cinematic_viscosity
                Case 38
                    'PROP_MS_38	Liquid Phase (Mixture) Dynamic Viscosity
                    value = su.spmp_viscosity
                Case 39
                    'PROP_MS_39	Liquid Phase (Mixture) Heat Capacity (Cp)
                    value = su.spmp_heatCapacityCp
                Case 40
                    'PROP_MS_40	Liquid Phase (Mixture) Heat Capacity Ratio (Cp/Cv)
                    value = ""
                Case 41
                    'PROP_MS_41	Liquid Phase (Mixture) Mass Flow
                    value = su.spmp_massflow
                Case 42
                    'PROP_MS_42	Liquid Phase (Mixture) Molar Flow
                    value = su.spmp_molarflow
                Case 43
                    'PROP_MS_43	Liquid Phase (Mixture) Volumetric Flow
                    value = su.spmp_volumetricFlow
                Case 44
                    'PROP_MS_44	Liquid Phase (Mixture) Compressibility Factor
                    value = ""
                Case 45
                    'PROP_MS_45	Liquid Phase (Mixture) Molar Fraction
                    value = ""
                Case 46
                    'PROP_MS_46	Liquid Phase (Mixture) Mass Fraction
                    value = ""
                Case 47
                    'PROP_MS_47	Liquid Phase (Mixture) Volumetric Fraction
                    value = ""
                Case 48
                    'PROP_MS_48	Liquid Phase (1) Density
                    value = su.spmp_density
                Case 49
                    'PROP_MS_49	Liquid Phase (1) Molar Weight
                    value = su.spmp_molecularWeight
                Case 50
                    'PROP_MS_50	Liquid Phase (1) Specific Enthalpy
                    value = su.spmp_enthalpy
                Case 51
                    'PROP_MS_51	Liquid Phase (1) Specific Entropy
                    value = su.spmp_entropy
                Case 52
                    'PROP_MS_52	Liquid Phase (1) Molar Enthalpy
                    value = su.molar_enthalpy
                Case 53
                    'PROP_MS_53	Liquid Phase (1) Molar Entropy
                    value = su.molar_entropy
                Case 54
                    'PROP_MS_54	Liquid Phase (1) Thermal Conductivity
                    value = su.spmp_thermalConductivity
                Case 55
                    'PROP_MS_55	Liquid Phase (1) Kinematic Viscosity
                    value = su.spmp_cinematic_viscosity
                Case 56
                    'PROP_MS_56	Liquid Phase (1) Dynamic Viscosity
                    value = su.spmp_viscosity
                Case 57
                    'PROP_MS_57	Liquid Phase (1) Heat Capacity (Cp)
                    value = su.spmp_heatCapacityCp
                Case 58
                    'PROP_MS_58	Liquid Phase (1) Heat Capacity Ratio (Cp/Cv)
                    value = ""
                Case 59
                    'PROP_MS_59	Liquid Phase (1) Mass Flow
                    value = su.spmp_massflow
                Case 60
                    'PROP_MS_60	Liquid Phase (1) Molar Flow
                    value = su.spmp_molarflow
                Case 61
                    'PROP_MS_61	Liquid Phase (1) Volumetric Flow
                    value = su.spmp_volumetricFlow
                Case 62
                    'PROP_MS_62	Liquid Phase (1) Compressibility Factor
                    value = ""
                Case 63
                    'PROP_MS_63	Liquid Phase (1) Molar Fraction
                    value = ""
                Case 64
                    'PROP_MS_64	Liquid Phase (1) Mass Fraction
                    value = ""
                Case 65
                    'PROP_MS_65	Liquid Phase (1) Volumetric Fraction
                    value = ""
                Case 66
                    'PROP_MS_66	Liquid Phase (2) Density
                    value = su.spmp_density
                Case 67
                    'PROP_MS_67	Liquid Phase (2) Molar Weight
                    value = su.spmp_molecularWeight
                Case 68
                    'PROP_MS_68	Liquid Phase (2) Specific Enthalpy
                    value = su.spmp_enthalpy
                Case 69
                    'PROP_MS_69	Liquid Phase (2) Specific Entropy
                    value = su.spmp_entropy
                Case 70
                    'PROP_MS_70	Liquid Phase (2) Molar Enthalpy
                    value = su.molar_enthalpy
                Case 71
                    'PROP_MS_71	Liquid Phase (2) Molar Entropy
                    value = su.molar_entropy
                Case 72
                    'PROP_MS_72	Liquid Phase (2) Thermal Conductivity
                    value = su.spmp_thermalConductivity
                Case 73
                    'PROP_MS_73	Liquid Phase (2) Kinematic Viscosity
                    value = su.spmp_cinematic_viscosity
                Case 74
                    'PROP_MS_74	Liquid Phase (2) Dynamic Viscosity
                    value = su.spmp_viscosity
                Case 75
                    'PROP_MS_75	Liquid Phase (2) Heat Capacity (Cp)
                    value = su.spmp_heatCapacityCp
                Case 76
                    'PROP_MS_76	Liquid Phase (2) Heat Capacity Ratio (Cp/Cv)
                    value = ""
                Case 77
                    'PROP_MS_77	Liquid Phase (2) Mass Flow
                    value = su.spmp_massflow
                Case 78
                    'PROP_MS_78	Liquid Phase (2) Molar Flow
                    value = su.spmp_molarflow
                Case 79
                    'PROP_MS_79	Liquid Phase (2) Volumetric Flow
                    value = su.spmp_volumetricFlow
                Case 80
                    'PROP_MS_80	Liquid Phase (2) Compressibility Factor
                    value = ""
                Case 81
                    'PROP_MS_81	Liquid Phase (2) Molar Fraction
                    value = ""
                Case 82
                    'PROP_MS_82	Liquid Phase (2) Mass Fraction
                    value = ""
                Case 83
                    'PROP_MS_83	Liquid Phase (2) Volumetric Fraction
                    value = ""
                Case 84
                    'PROP_MS_84	Aqueous Phase Density
                    value = su.spmp_density
                Case 85
                    'PROP_MS_85	Aqueous Phase Molar Weight
                    value = su.spmp_molecularWeight
                Case 86
                    'PROP_MS_86	Aqueous Phase Specific Enthalpy
                    value = su.spmp_enthalpy
                Case 87
                    'PROP_MS_87	Aqueous Phase Specific Entropy
                    value = su.spmp_entropy
                Case 88
                    'PROP_MS_88	Aqueous Phase Molar Enthalpy
                    value = su.molar_enthalpy
                Case 89
                    'PROP_MS_89	Aqueous Phase Molar Entropy
                    value = su.molar_entropy
                Case 90
                    'PROP_MS_90	Aqueous Phase Thermal Conductivity
                    value = su.spmp_thermalConductivity
                Case 91
                    'PROP_MS_91	Aqueous Phase Kinematic Viscosity
                    value = su.spmp_cinematic_viscosity
                Case 92
                    'PROP_MS_92	Aqueous Phase Dynamic Viscosity
                    value = su.spmp_viscosity
                Case 93
                    'PROP_MS_93	Aqueous Phase Heat Capacity (Cp)
                    value = su.spmp_heatCapacityCp
                Case 94
                    'PROP_MS_94	Aqueous Phase Heat Capacity Ratio (Cp/Cv)
                    value = ""
                Case 95
                    'PROP_MS_95	Aqueous Phase Mass Flow
                    value = su.spmp_massflow
                Case 96
                    'PROP_MS_96	Aqueous Phase Molar Flow
                    value = su.spmp_molarflow
                Case 97
                    'PROP_MS_97	Aqueous Phase Volumetric Flow
                    value = su.spmp_volumetricFlow
                Case 98
                    'PROP_MS_98	Aqueous Phase Compressibility Factor
                    value = ""
                Case 99
                    'PROP_MS_99	Aqueous Phase Molar Fraction
                    value = ""
                Case 100
                    'PROP_MS_100	Aqueous Phase Mass Fraction
                    value = ""
                Case 101
                    'PROP_MS_101	Aqueous Phase Volumetric Fraction
                    value = ""
                Case 102, 103, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115
                    value = ""
                Case 104, 116, 117, 118, 119, 120
                    value = su.spmp_molarflow
                Case 105, 121, 122, 123, 124, 125
                    value = su.spmp_massflow
                Case 126, 127
                    value = su.spmp_pressure
                Case 128, 129
                    value = su.spmp_temperature
                Case 130
                    value = ""
            End Select

            Return value
        End Function

#End Region

#Region "    CAPE-OPEN 1.0 Methods and Properties"

        Public Sub PropList1(ByRef props As Object, ByRef phases As Object, ByRef calcType As Object) Implements CAPEOPEN110.ICapeThermoCalculationRoutine.PropList
            Throw New NotImplementedException
        End Sub

        Public Function PropCheck2(ByVal materialObject As Object, ByVal flashType As String, ByVal props As Object) As Object Implements CAPEOPEN110.ICapeThermoEquilibriumServer.PropCheck
            Throw New NotImplementedException
        End Function

        Public Function ValidityCheck2(ByVal materialObject As Object, ByVal props As Object) As Object Implements CAPEOPEN110.ICapeThermoEquilibriumServer.ValidityCheck
            Throw New NotImplementedException
        End Function

        ''' <summary>
        ''' Gets the name of the component.
        ''' </summary>
        ''' <value></value>
        ''' <returns>CapeString</returns>
        ''' <remarks>Implements CapeOpen.ICapeIdentification.ComponentDescription</remarks>
        Public Overridable Property ComponentDescription() As String Implements CapeOpen.ICapeIdentification.ComponentDescription
            Get
                Return Me.m_ComponentName
            End Get
            Set(ByVal value As String)
                Me.m_ComponentName = value
            End Set
        End Property

        ''' <summary>
        ''' Gets the description of the component.
        ''' </summary>
        ''' <value></value>
        ''' <returns>CapeString</returns>
        ''' <remarks>Implements CapeOpen.ICapeIdentification.ComponentName</remarks>
        Public Property ComponentName() As String Implements CapeOpen.ICapeIdentification.ComponentName
            Get
                Return "temporary stream"
            End Get
            Set(ByVal value As String)

            End Set
        End Property

        ''' <summary>
        ''' Gets a list of properties that have been calculated.
        ''' </summary>
        ''' <returns>Properties for which results are available.</returns>
        ''' <remarks>Not implemented in DTL.</remarks>
        Public Function AvailableProps() As Object Implements CapeOpen.ICapeThermoMaterialObject.AvailableProps
            Throw New NotImplementedException
        End Function

        ''' <summary>
        ''' This method is responsible for calculating a flash or delegating flash calculations to the associated Property Package or Equilibrium Server.
        ''' </summary>
        ''' <param name="flashType">Flash calculation type.</param>
        ''' <param name="props">Properties to be calculated at equilibrium. UNDEFINED for no properties. If a list, then the 
        ''' property values should be set for each phase present at equilibrium (not including the overall phase).</param>
        ''' <remarks>The CalcEquilibrium method must set on the Material Object the amounts (phaseFraction) and compositions
        ''' (fraction) for all phases present at equilibrium, as well as the temperature and pressure for the overall
        ''' mixture, if not set as part of the calculation specifications. The CalcEquilibrium method must not set on the
        ''' Material Object any other value - in particular it must not set any values for phases that do not exist. See
        ''' 5.2.1 for more information.
        ''' The available list of flashes is given in section 5.6.1.
        ''' It is advised not to combine a flash calculation with a property calculation. Although by the returned error
        ''' one cannot see which has failed, plus the additional arguments to CalcProp (such as calculation type) cannot
        ''' be specified. Advice is to perform a CalcEquilibrium, get the phaseIDs and perform a CalcProp on the
        ''' existing phases.
        ''' The Material Object may or may not delegate this call to a Property Package.</remarks>
        Public Sub CalcEquilibrium(ByVal flashType As String, ByVal props As Object) Implements CapeOpen.ICapeThermoMaterialObject.CalcEquilibrium
            Me.PropertyPackage.CurrentMaterialStream = Me
            Try
                Me.PropertyPackage.CalcEquilibrium(Me, flashType, props)
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoMaterialObject", ex.Source, ex.StackTrace, "CalcEquilibrium", hcode)
            End Try
        End Sub

        ''' <summary>
        ''' This method is responsible for doing all property calculations or delegating these calculations to the
        ''' associated Property Package.
        ''' </summary>
        ''' <param name="props">The List of Properties to be calculated.</param>
        ''' <param name="phases">List of phases for which the Properties are to be calculated.</param>
        ''' <param name="calcType">Type of calculation: Mixture Property or Pure Compound Property. For
        ''' partial property, such as fugacity coefficients of compounds in a
        ''' mixture, use “Mixture” CalcType. For pure compound fugacity
        ''' coefficients, use “Pure” CalcType.</param>
        ''' <remarks>"Pure" calctype is not implemented in DTL.</remarks>
        Public Sub CalcProp(ByVal props As Object, ByVal phases As Object, ByVal calcType As String) Implements CapeOpen.ICapeThermoMaterialObject.CalcProp
            Me.PropertyPackage.CurrentMaterialStream = Me
            Try
                Me.PropertyPackage.CalcProp(Me, props, phases, calcType)
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoMaterialObject", ex.Source, ex.StackTrace, "CalcProp", hcode)
            End Try
        End Sub

        ''' <summary>
        ''' Returns the list of compound IDs of a given Material Object.
        ''' </summary>
        ''' <value></value>
        ''' <returns>Compound IDs</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ComponentIds() As Object Implements CapeOpen.ICapeThermoMaterialObject.ComponentIds
            Get
                Dim compids As Object = Nothing
                Dim formulas As Object = Nothing
                Dim mols As Object = Nothing
                Dim cas As Object = Nothing
                Dim nbps As Object = Nothing
                Dim names As Object = Nothing
                Me.PropertyPackage.CurrentMaterialStream = Me
                Try
                    Me.PropertyPackage.GetComponentList(compids, formulas, names, nbps, mols, cas)
                Catch ex As Exception
                    Dim hcode As Integer = 0
                    Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                    If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                    ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoMaterialObject", ex.Source, ex.StackTrace, "ComponentIds", hcode)
                End Try
                Return compids
            End Get
        End Property

        ''' <summary>
        ''' Create a Material Object from the parent Material Template of the current Material Object.
        ''' </summary>
        ''' <returns>The created and initialized Material Object.</returns>
        ''' <remarks></remarks>
        Public Function CreateMaterialObject() As Object Implements CapeOpen.ICapeThermoMaterialObject.CreateMaterialObject
            Dim mat As New Streams.MaterialStream("temporary stream", "temporary stream")
            Return mat
        End Function

        ''' <summary>
        ''' Creates a duplicate of the current Material Object.
        ''' </summary>
        ''' <returns>The duplicated Material Object.</returns>
        ''' <remarks></remarks>
        Public Function Duplicate() As Object Implements CapeOpen.ICapeThermoMaterialObject.Duplicate
            Dim newmat As MaterialStream = Me.Clone
            Return newmat
        End Function

        ''' <summary>
        ''' Retrieve pure compound constants from the Property Package.
        ''' </summary>
        ''' <param name="props">List of pure compound constants</param>
        ''' <param name="compIds">List of compound IDs for which constants are to be retrieved.
        '''UNDEFINED is to be used when the call applied to all compounds in
        '''the Material Object.</param>
        ''' <returns>Compound Constant values returned from the Property Package for the
        ''' specified compounds.</returns>
        ''' <remarks></remarks>
        Public Function GetComponentConstant(ByVal props As Object, ByVal compIds As Object) As Object Implements CapeOpen.ICapeThermoMaterialObject.GetComponentConstant
            Me.PropertyPackage.CurrentMaterialStream = Me
            Dim obj As Object = Nothing
            Try
                obj = Me.PropertyPackage.GetCompoundConstant(props, compIds)
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoMaterialObject", ex.Source, ex.StackTrace, "ComponentIds", hcode)
            End Try
            Return obj
        End Function

        ''' <summary>
        ''' Returns the independent variables of a Material Object. This method is deprecated.
        ''' </summary>
        ''' <param name="indVars">Independent variables to be set</param>
        ''' <returns>Values of independent variables.</returns>
        ''' <remarks>This method should not be used.</remarks>
        Public Function GetIndependentVar(ByVal indVars As Object) As Object Implements CapeOpen.ICapeThermoMaterialObject.GetIndependentVar
            Throw New NotImplementedException
        End Function

        ''' <summary>
        ''' Returns number of chemical compounds in Material Object.
        ''' </summary>
        ''' <returns>Number of compounds in the Material Object.</returns>
        ''' <remarks></remarks>
        Public Function GetNumComponents() As Integer Implements CapeOpen.ICapeThermoMaterialObject.GetNumComponents
            Me.PropertyPackage.CurrentMaterialStream = Me
            Return Me.Fases(0).Componentes.Count
        End Function

        ''' <summary>
        ''' This method is responsible for retrieving the results from calculations from the Material Object.
        ''' </summary>
        ''' <param name="property">The Property for which results are requested from the Material Object.</param>
        ''' <param name="phase">The qualified phase for the results.</param>
        ''' <param name="compIds">The qualified compounds for the results. UNDEFINED to specify all
        ''' compounds in the Material Object. For scalar mixture properties such
        ''' as liquid enthalpy, this qualifier must not be specified. Use
        ''' UNDEFINED as place holder.</param>
        ''' <param name="calcType">The qualified type of calculation for the results. (valid Calculation Types: Pure and Mixture)</param>
        ''' <param name="basis">Qualifies the basis of the result (i.e., mass /mole). Use UNDEFINED
        ''' for default or as place holder for property for which basis does not apply (see also 3.3.1).</param>
        ''' <returns>Results vector containing property values in SI units arranged by the defined qualifiers.</returns>
        ''' <remarks></remarks>
        Public Function GetProp(ByVal [property] As String, ByVal phase As String, ByVal compIds As Object, ByVal calcType As String, ByVal basis As String) As Object Implements CapeOpen.ICapeThermoMaterialObject.GetProp

            Me.PropertyPackage.CurrentMaterialStream = Me

            If Not calcType Is Nothing Then
                If calcType = "Pure" Then Throw New NotImplementedException
            End If
            Dim res As New ArrayList
            Dim comps As New ArrayList
            For Each c As Substancia In Me.Fases(0).Componentes.Values
                comps.Add(c.Nome)
            Next
            Dim f As Integer = 0
            Dim phs As DTL.SimulationObjects.PropertyPackages.Fase
            Select Case phase.ToLower
                Case "overall"
                    f = 0
                    phs = PropertyPackages.Fase.Mixture
                Case Else
                    For Each pi As PhaseInfo In Me.PropertyPackage.PhaseMappings.Values
                        If phase = pi.PhaseLabel Then
                            f = pi.DWPhaseIndex
                            phs = pi.DWPhaseID
                            Exit For
                        End If
                    Next
            End Select
            Select Case [property].ToLower
                Case "compressibilityfactor"
                    res.Add(Me.Fases(f).SPMProperties.compressibilityFactor.GetValueOrDefault)
                Case "heatofvaporization"
                Case "heatcapacity", "heatcapacitycp"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            res.Add(Me.Fases(f).SPMProperties.heatCapacityCp * Me.PropertyPackage.AUX_MMM(phs))
                        Case "Mass", "mass"
                            res.Add(Me.Fases(f).SPMProperties.heatCapacityCp * 1000)
                    End Select
                Case "heatcapacitycv"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            res.Add(Me.Fases(f).SPMProperties.heatCapacityCv * Me.PropertyPackage.AUX_MMM(phs))
                        Case "Mass", "mass"
                            res.Add(Me.Fases(f).SPMProperties.heatCapacityCv * 1000)
                    End Select
                Case "idealgasheatcapacity"
                    If f = 1 Then
                        res.Add(Me.PropertyPackage.AUX_CPm(PropertyPackages.Fase.Liquid, Me.Fases(0).SPMProperties.temperature * 1000))
                    Else
                        res.Add(Me.PropertyPackage.AUX_CPm(PropertyPackages.Fase.Vapor, Me.Fases(0).SPMProperties.temperature * 1000))
                    End If
                Case "idealgasenthalpy"
                    If f = 1 Then
                        res.Add(Me.PropertyPackage.RET_Hid(298.15, Me.Fases(0).SPMProperties.temperature.GetValueOrDefault * 1000, PropertyPackages.Fase.Liquid))
                    Else
                        res.Add(Me.PropertyPackage.RET_Hid(298.15, Me.Fases(0).SPMProperties.temperature.GetValueOrDefault * 1000, PropertyPackages.Fase.Vapor))
                    End If
                Case "excessenthalpy"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            res.Add(Me.Fases(f).SPMProperties.excessEnthalpy.GetValueOrDefault * Me.PropertyPackage.AUX_MMM(phs))
                        Case "Mass", "mass"
                            res.Add(Me.Fases(f).SPMProperties.excessEnthalpy.GetValueOrDefault * 1000)
                    End Select
                Case "excessentropy"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            res.Add(Me.Fases(f).SPMProperties.excessEntropy.GetValueOrDefault * Me.PropertyPackage.AUX_MMM(phs))
                        Case "Mass", "mass"
                            res.Add(Me.Fases(f).SPMProperties.excessEntropy.GetValueOrDefault * 1000)
                    End Select
                Case "viscosity"
                    res.Add(Me.Fases(f).SPMProperties.viscosity.GetValueOrDefault)
                Case "thermalconductivity"
                    res.Add(Me.Fases(f).SPMProperties.thermalConductivity.GetValueOrDefault)
                Case "fugacity"
                    For Each c As String In comps
                        res.Add(Me.Fases(f).Componentes(c).FracaoMolar.GetValueOrDefault * Me.Fases(f).Componentes(c).FugacityCoeff.GetValueOrDefault * Me.Fases(0).SPMProperties.pressure.GetValueOrDefault)
                    Next
                Case "activity"
                    For Each c As String In comps
                        res.Add(Me.Fases(f).Componentes(c).ActivityCoeff.GetValueOrDefault * Me.Fases(f).Componentes(c).FracaoMolar.GetValueOrDefault)
                    Next
                Case "fugacitycoefficient"
                    For Each c As String In comps
                        res.Add(Me.Fases(f).Componentes(c).FugacityCoeff)
                    Next
                Case "activitycoefficient"
                    For Each c As String In comps
                        res.Add(Me.Fases(f).Componentes(c).ActivityCoeff)
                    Next
                Case "logfugacitycoefficient"
                    For Each c As String In comps
                        res.Add(Math.Log(Me.Fases(f).Componentes(c).FugacityCoeff))
                    Next
                Case "volume"
                    res.Add(Me.Fases(f).SPMProperties.molecularWeight / Me.Fases(f).SPMProperties.density / 1000)
                Case "density"
                    res.Add(Me.Fases(f).SPMProperties.density)
                Case "enthalpy", "enthalpynf"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Dim val = Me.Fases(f).SPMProperties.molecularWeight.GetValueOrDefault
                            If val = 0.0# Then
                                res.Add(Me.Fases(f).SPMProperties.molar_enthalpy.GetValueOrDefault)
                            Else
                                res.Add(Me.Fases(f).SPMProperties.enthalpy.GetValueOrDefault * val)
                            End If
                        Case "Mass", "mass"
                            res.Add(Me.Fases(f).SPMProperties.enthalpy.GetValueOrDefault * 1000)
                    End Select
                Case "entropy", "entropynf"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Dim val = Me.Fases(f).SPMProperties.molecularWeight.GetValueOrDefault
                            If val = 0.0# Then
                                res.Add(Me.Fases(f).SPMProperties.molar_entropy.GetValueOrDefault)
                            Else
                                res.Add(Me.Fases(f).SPMProperties.entropy.GetValueOrDefault * val)
                            End If
                        Case "Mass", "mass"
                            res.Add(Me.Fases(f).SPMProperties.entropy.GetValueOrDefault * 1000)
                    End Select
                Case "enthalpyf"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Dim val = Me.Fases(f).SPMProperties.molecularWeight.GetValueOrDefault
                            If val = 0.0# Then
                                res.Add(Me.Fases(f).SPMProperties.molar_enthalpyF.GetValueOrDefault)
                            Else
                                res.Add(Me.Fases(f).SPMProperties.enthalpyF.GetValueOrDefault * val)
                            End If
                        Case "Mass", "mass"
                            res.Add(Me.Fases(f).SPMProperties.enthalpyF.GetValueOrDefault * 1000)
                    End Select
                Case "entropyf"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Dim val = Me.Fases(f).SPMProperties.molecularWeight.GetValueOrDefault
                            If val = 0.0# Then
                                res.Add(Me.Fases(f).SPMProperties.molar_entropyF.GetValueOrDefault)
                            Else
                                res.Add(Me.Fases(f).SPMProperties.entropyF.GetValueOrDefault * val)
                            End If
                        Case "Mass", "mass"
                            res.Add(Me.Fases(f).SPMProperties.entropy.GetValueOrDefault * 1000)
                    End Select
                Case "moles"
                    res.Add(Me.Fases(f).SPMProperties.molarflow.GetValueOrDefault)
                Case "mass"
                    res.Add(Me.Fases(f).SPMProperties.massflow.GetValueOrDefault)
                Case "molecularweight"
                    res.Add(Me.Fases(f).SPMProperties.molecularWeight)
                Case "temperature"
                    res.Add(Me.Fases(0).SPMProperties.temperature.GetValueOrDefault)
                Case "pressure"
                    res.Add(Me.Fases(0).SPMProperties.pressure.GetValueOrDefault)
                Case "flow"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            For Each c As String In comps
                                res.Add(Me.Fases(f).Componentes(c).MolarFlow.GetValueOrDefault)
                            Next
                        Case "Mass", "mass"
                            For Each c As String In comps
                                res.Add(Me.Fases(f).Componentes(c).MassFlow.GetValueOrDefault)
                            Next
                    End Select
                Case "fraction", "massfraction", "molarfraction"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole", ""
                            For Each c As String In comps
                                res.Add(Me.Fases(f).Componentes(c).FracaoMolar.GetValueOrDefault)
                            Next
                        Case "Mass", "mass"
                            For Each c As String In comps
                                res.Add(Me.Fases(f).Componentes(c).FracaoMassica.GetValueOrDefault)
                            Next
                        Case ""
                            If [property].ToLower.Contains("mole") Then
                                For Each c As String In comps
                                    res.Add(Me.Fases(f).Componentes(c).FracaoMolar.GetValueOrDefault)
                                Next
                            ElseIf [property].ToLower.Contains("mass") Then
                                For Each c As String In comps
                                    res.Add(Me.Fases(f).Componentes(c).FracaoMassica.GetValueOrDefault)
                                Next
                            End If
                    End Select
                Case "concentration"
                    For Each c As String In comps
                        res.Add(Me.Fases(f).Componentes(c).MassFlow.GetValueOrDefault / Me.Fases(f).SPMProperties.volumetric_flow.GetValueOrDefault)
                    Next
                Case "molarity"
                    For Each c As String In comps
                        res.Add(Me.Fases(f).Componentes(c).MolarFlow.GetValueOrDefault / Me.Fases(f).SPMProperties.volumetric_flow.GetValueOrDefault)
                    Next
                Case "phasefraction"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            res.Add(Me.Fases(f).SPMProperties.molarfraction.GetValueOrDefault)
                        Case "Mass", "mass"
                            res.Add(Me.Fases(f).SPMProperties.massfraction.GetValueOrDefault)
                    End Select
                Case "totalflow"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            res.Add(Me.Fases(f).SPMProperties.molarflow.GetValueOrDefault)
                        Case "Mass", "mass"
                            res.Add(Me.Fases(f).SPMProperties.massflow.GetValueOrDefault)
                    End Select
                Case "kvalue"
                    For Each c As String In comps
                        res.Add(Me.Fases(0).Componentes(c).Kvalue)
                    Next
                Case "logkvalue"
                    For Each c As String In comps
                        res.Add(Me.Fases(0).Componentes(c).lnKvalue)
                    Next
                Case "surfacetension"
                    res.Add(Me.Fases(1).TPMProperties.surfaceTension)
                Case Else
                    Dim ex As New Exception
                    Dim hcode As Integer = 0
                    ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoMaterialObject", ex.Source, ex.StackTrace, "GetProp", hcode)
            End Select

            Dim arr(res.Count - 1) As Double
            Array.Copy(res.ToArray, arr, res.Count)
            Return arr

        End Function

        ''' <summary>
        ''' Returns list of properties that can be calculated by the Material Object.
        ''' </summary>
        ''' <returns>List of all supported properties of the Material Object.</returns>
        ''' <remarks>DWSIM passes this call to the Property Package currently associated to the stream.</remarks>
        Public Function GetPropList() As Object Implements CapeOpen.ICapeThermoMaterialObject.GetPropList
            Dim mylist As Object = Me.PropertyPackage.GetPropList
            Return mylist
        End Function

        ''' <summary>
        ''' Retrieves values of universal constants from the Property Package.
        ''' </summary>
        ''' <param name="props">List of universal constants to be retrieved</param>
        ''' <returns>Values of universal constants</returns>
        ''' <remarks>DWSIM passes this call to the Property Package currently associated to the stream.</remarks>
        Public Function GetUniversalConstant(ByVal props As Object) As Object Implements CapeOpen.ICapeThermoMaterialObject.GetUniversalConstant
            Return Me.PropertyPackage.GetUniversalConstant(Me, props)
        End Function

        ''' <summary>
        ''' It returns the phases existing in the Material Object at that moment.
        ''' </summary>
        ''' <value></value>
        ''' <returns>List of phases</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property PhaseIds() As Object Implements CapeOpen.ICapeThermoMaterialObject.PhaseIds
            Get
                Dim pl As New ArrayList
                For Each pi As PhaseInfo In Me.PropertyPackage.PhaseMappings.Values
                    If pi.PhaseLabel <> "Disabled" And Me.Fases(pi.DWPhaseIndex).SPMProperties.molarfraction > 0 Then
                        pl.Add(pi.PhaseLabel)
                    End If
                Next
                Dim arr(pl.Count - 1) As String
                Array.Copy(pl.ToArray, arr, pl.Count)
                Return arr
            End Get
        End Property

        ''' <summary>
        ''' Checks to see if a list of given properties can be calculated.
        ''' </summary>
        ''' <param name="props">Properties to check.</param>
        ''' <returns>Returns Boolean List associated to list of properties to be checked.</returns>
        ''' <remarks>Not implemented in DTL. As it was unclear from the original specification what PropCheck should exactly be checking, and as the
        ''' argument list does not include a phase specification, implementations vary. It is generally expected that
        ''' PropCheck at least verifies that the Property is available for calculation in the Material Object. However, this
        ''' can also be verified with PropList. It is advised not to use PropCheck.
        ''' The Material Object may or may not delegate this call to a Property Package.</remarks>
        Public Function PropCheck(ByVal props As Object) As Object Implements CapeOpen.ICapeThermoMaterialObject.PropCheck
            Throw New NotImplementedException
        End Function

        ''' <summary>
        ''' RemoveResults
        ''' </summary>
        ''' <param name="props">Properties to be removed. UNDEFINED to remove all properties.</param>
        ''' <remarks>Not implemented.</remarks>
        Public Sub RemoveResults(ByVal props As Object) Implements CapeOpen.ICapeThermoMaterialObject.RemoveResults
            Throw New NotImplementedException
        End Sub

        ''' <summary>
        ''' SetIndependentVar
        ''' </summary>
        ''' <param name="indVars">Sets the independent variable for a given Material Object. This method is deprecated.</param>
        ''' <param name="values">Independent variables to be set</param>
        ''' <remarks>Values of independent variables.</remarks>
        Public Sub SetIndependentVar(ByVal indVars As Object, ByVal values As Object) Implements CapeOpen.ICapeThermoMaterialObject.SetIndependentVar
            Throw New NotImplementedException
        End Sub

        ''' <summary>
        ''' This method is responsible for setting the values for properties of the Material Object.
        ''' </summary>
        ''' <param name="property">The property for which the values need to be set.</param>
        ''' <param name="phase">Phase for which the property is to be set.</param>
        ''' <param name="compIds">Compounds for which values are to be set. UNDEFINED to specify all
        ''' compounds in the Material Object. For scalar mixture properties such
        ''' as liquid enthalpy, this qualifier should not be specified. Use
        ''' UNDEFINED as place holder.</param>
        ''' <param name="calcType">The calculation type. (valid Calculation Types: Pure and Mixture)</param>
        ''' <param name="basis">Qualifies the basis (mole / mass). See also 3.3.2.</param>
        ''' <param name="values">Values to set for the property.</param>
        ''' <remarks>DWSIM doesn't implement "Pure" calculation type.</remarks>
        Public Sub SetProp(ByVal [property] As String, ByVal phase As String, ByVal compIds As Object, ByVal calcType As String, ByVal basis As String, ByVal values As Object) Implements CapeOpen.ICapeThermoMaterialObject.SetProp
            Me.PropertyPackage.CurrentMaterialStream = Me
            If Not calcType Is Nothing Then
                If calcType = "Pure" Then Throw New NotImplementedException
            End If
            Dim res As New ArrayList
            Dim length As Integer, i As Integer, comps As New ArrayList
            If Not compIds Is Nothing Then
                length = compIds.length
                For i = 0 To length - 1
                    comps.Add(compIds(i))
                Next
            Else
                For Each c As Substancia In Me.Fases(0).Componentes.Values
                    comps.Add(c.Nome)
                Next
            End If
            Dim f As Integer = -1
            Dim phs As DTL.SimulationObjects.PropertyPackages.Fase
            Select Case phase.ToLower
                Case "overall"
                    f = 0
                    phs = PropertyPackages.Fase.Mixture
                Case Else
                    For Each pi As PhaseInfo In Me.PropertyPackage.PhaseMappings.Values
                        If phase = pi.PhaseLabel Then
                            f = pi.DWPhaseIndex
                            phs = pi.DWPhaseID
                            Exit For
                        End If
                    Next
            End Select
            Select Case [property].ToLower
                Case "compressibilityfactor"
                    Me.Fases(f).SPMProperties.compressibilityFactor = values(0)
                Case "heatcapacity", "heatcapacitycp"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Me.Fases(f).SPMProperties.heatCapacityCp = values(0) / Me.PropertyPackage.AUX_MMM(phs)
                        Case "Mass", "mass"
                            Me.Fases(f).SPMProperties.heatCapacityCp = values(0) / 1000
                    End Select
                Case "heatcapacitycv"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Me.Fases(f).SPMProperties.heatCapacityCv = values(0) / Me.PropertyPackage.AUX_MMM(phs)
                        Case "Mass", "mass"
                            Me.Fases(f).SPMProperties.heatCapacityCv = values(0) / 1000
                    End Select
                Case "excessenthalpy"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Me.Fases(f).SPMProperties.excessEnthalpy = values(0) / Me.PropertyPackage.AUX_MMM(phs)
                        Case "Mass", "mass"
                            Me.Fases(f).SPMProperties.excessEnthalpy = values(0) / 1000
                    End Select
                Case "excessentropy"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Me.Fases(f).SPMProperties.excessEntropy = values(0) / Me.PropertyPackage.AUX_MMM(phs)
                        Case "Mass", "mass"
                            Me.Fases(f).SPMProperties.excessEntropy = values(0) / 1000
                    End Select
                Case "viscosity"
                    Me.Fases(f).SPMProperties.viscosity = values(0)
                    Me.Fases(f).SPMProperties.kinematic_viscosity = values(0) / Me.Fases(f).SPMProperties.density.GetValueOrDefault
                Case "thermalconductivity"
                    Me.Fases(f).SPMProperties.thermalConductivity = values(0)
                Case "fugacity"
                    i = 0
                    For Each c As String In comps
                        Me.Fases(f).Componentes(c).FugacityCoeff = values(comps.IndexOf(c)) / (Me.Fases(0).SPMProperties.pressure.GetValueOrDefault * Me.Fases(f).Componentes(c).FracaoMolar.GetValueOrDefault)
                        i += 1
                    Next
                Case "fugacitycoefficient"
                    i = 0
                    For Each c As String In comps
                        Me.Fases(f).Componentes(c).FugacityCoeff = values(comps.IndexOf(c))
                        i += 1
                    Next
                Case "activitycoefficient"
                    i = 0
                    For Each c As String In comps
                        Me.Fases(f).Componentes(c).ActivityCoeff = values(comps.IndexOf(c))
                        i += 1
                    Next
                Case "logfugacitycoefficient"
                    i = 0
                    For Each c As String In comps
                        Me.Fases(f).Componentes(c).FugacityCoeff = Math.Exp(values(comps.IndexOf(c)))
                        i += 1
                    Next
                Case "density"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Me.Fases(f).SPMProperties.density = values(0) / 1000 * Me.PropertyPackage.AUX_MMM(phs)
                        Case "Mass", "mass"
                            Me.Fases(f).SPMProperties.density = values(0)
                    End Select
                Case "volume"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Me.Fases(f).SPMProperties.density = 1 / values(0) / 1000 * Me.PropertyPackage.AUX_MMM(phs)
                        Case "Mass", "mass"
                            Me.Fases(f).SPMProperties.density = 1 / values(0)
                    End Select
                Case "enthalpy", "enthalpynf"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Me.Fases(f).SPMProperties.molar_enthalpy = values(0)
                            Dim val = Me.PropertyPackage.AUX_MMM(phs)
                            If val <> 0.0# Then Me.Fases(f).SPMProperties.enthalpy = values(0) / val
                        Case "Mass", "mass"
                            Me.Fases(f).SPMProperties.enthalpy = values(0) / 1000
                            Me.Fases(f).SPMProperties.molar_enthalpy = values(0) * Me.PropertyPackage.AUX_MMM(phs)
                    End Select
                Case "entropy", "entropynf"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Me.Fases(f).SPMProperties.molar_entropy = values(0)
                            Dim val = Me.PropertyPackage.AUX_MMM(phs)
                            If val <> 0.0# Then Me.Fases(f).SPMProperties.entropy = values(0) / val
                        Case "Mass", "mass"
                            Me.Fases(f).SPMProperties.entropy = values(0) / 1000
                            Me.Fases(f).SPMProperties.molar_entropy = values(0) * Me.PropertyPackage.AUX_MMM(phs)
                    End Select
                Case "enthalpyf"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Dim val = Me.PropertyPackage.AUX_MMM(phs)
                            Me.Fases(f).SPMProperties.molar_enthalpyF = values(0)
                            If val <> 0.0# Then Me.Fases(f).SPMProperties.enthalpyF = values(0) / val
                        Case "Mass", "mass"
                            Me.Fases(f).SPMProperties.enthalpyF = values(0) / 1000
                            Me.Fases(f).SPMProperties.molar_enthalpyF = values(0) * Me.PropertyPackage.AUX_MMM(phs)
                    End Select
                Case "entropyf"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Dim val = Me.PropertyPackage.AUX_MMM(phs)
                            Me.Fases(f).SPMProperties.molar_entropyF = values(0)
                            If val <> 0.0# Then Me.Fases(f).SPMProperties.entropyF = values(0) / val
                        Case "Mass", "mass"
                            Me.Fases(f).SPMProperties.entropyF = values(0) / 1000
                            Me.Fases(f).SPMProperties.molar_entropyF = values(0) * Me.PropertyPackage.AUX_MMM(phs)
                    End Select
                Case "moles"
                    Me.Fases(f).SPMProperties.molarflow = values(0)
                Case "mass"
                    Me.Fases(f).SPMProperties.massflow = values(0)
                Case "molecularweight"
                    Me.Fases(f).SPMProperties.molecularWeight = values(0)
                Case "temperature"
                    Me.Fases(0).SPMProperties.temperature = values(0)
                Case "pressure"
                    Me.Fases(0).SPMProperties.pressure = values(0)
                Case "flow"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            i = 0
                            For Each c As String In comps
                                Me.Fases(f).Componentes(c).MolarFlow = values(comps.IndexOf(c))
                                i += 1
                            Next
                        Case "Mass", "mass"
                            i = 0
                            For Each c As String In comps
                                Me.Fases(f).Componentes(c).MassFlow = values(comps.IndexOf(c))
                                i += 1
                            Next
                    End Select
                Case "fraction"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            i = 0
                            For Each c As String In comps
                                Me.Fases(f).Componentes(c).FracaoMolar = values(comps.IndexOf(c))
                                i += 1
                            Next
                        Case "Mass", "mass"
                            i = 0
                            For Each c As String In comps
                                Me.Fases(f).Componentes(c).FracaoMassica = values(comps.IndexOf(c))
                                i += 1
                            Next
                    End Select
                Case "phasefraction"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Me.Fases(f).SPMProperties.molarfraction = values(0)
                        Case "Mass", "mass"
                            Me.Fases(f).SPMProperties.massfraction = values(0)
                    End Select
                Case "totalflow"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Me.Fases(f).SPMProperties.molarflow = values(0)
                            'Me.Fases(f).SPMProperties.massflow = values(0) / 1000 * Me.PropertyPackage.AUX_MMM(phs)
                        Case "Mass", "mass"
                            Me.Fases(f).SPMProperties.massflow = values(0)
                            'Me.Fases(f).SPMProperties.molarflow = values(0)
                    End Select
                Case "kvalue"
                    i = 0
                    For Each c As String In comps
                        Me.Fases(0).Componentes(c).Kvalue = values(comps.IndexOf(c))
                        i += 1
                    Next
                Case "logkvalue"
                    i = 0
                    For Each c As String In comps
                        Me.Fases(0).Componentes(c).lnKvalue = values(comps.IndexOf(c))
                        i += 1
                    Next
                Case "surfacetension"
                    Me.Fases(0).TPMProperties.surfaceTension = values(0)
                Case Else
                    Dim ex As New Exception
                    Dim hcode As Integer = 0
                    ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoMaterialObject", ex.Source, ex.StackTrace, "ComponentIds", hcode)
            End Select
        End Sub

        ''' <summary>
        ''' Checks the validity of the calculation. This method is deprecated.
        ''' </summary>
        ''' <param name="props">The properties for which reliability is checked.</param>
        ''' <returns>Returns the reliability scale of the calculation.</returns>
        ''' <remarks>The ValidityCheck method must not be used, since the ICapeThermoReliability interface is not yet defined.</remarks>
        Public Function ValidityCheck(ByVal props As Object) As Object Implements CapeOpen.ICapeThermoMaterialObject.ValidityCheck
            Return Me.PropertyPackage.ValidityCheck(Me, props)
        End Function

        Public Sub CalcProp1(ByVal materialObject As Object, ByVal props As Object, ByVal phases As Object, ByVal calcType As String) Implements CapeOpen.ICapeThermoCalculationRoutine.CalcProp
            CalcProp1(materialObject, props, phases, calcType)
        End Sub

        Public Function PropCheck1(ByVal materialObject As Object, ByVal props As Object) As Object Implements CapeOpen.ICapeThermoCalculationRoutine.PropCheck
            Return PropCheck1(materialObject, props)
        End Function

        Public Function ValidityCheck1(ByVal materialObject As Object, ByVal props As Object) As Object Implements CapeOpen.ICapeThermoCalculationRoutine.ValidityCheck
            Me.PropertyPackage.CurrentMaterialStream = Me
            Return ValidityCheck(props)
        End Function

        Public Sub CalcEquilibrium2(ByVal materialObject As Object, ByVal flashType As String, ByVal props As Object) Implements CapeOpen.ICapeThermoEquilibriumServer.CalcEquilibrium
            Try
                Me.PropertyPackage.CalcEquilibrium(materialObject, flashType, props)
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoEquilibriumServer", ex.Source, ex.StackTrace, "ComponentIds", hcode)
            End Try
        End Sub

        Public Sub PropList(ByRef flashType As Object, ByRef props As Object, ByRef phases As Object, ByRef calcType As Object) Implements CapeOpen.ICapeThermoEquilibriumServer.PropList
            Try
                Me.PropertyPackage.PropList(flashType, props, phases, calcType)
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoEquilibriumServer", ex.Source, ex.StackTrace, "PropList", hcode)
            End Try
        End Sub

        Public Sub CalcEquilibrium3(ByVal materialObject As Object, ByVal flashType As String, ByVal props As Object) Implements CapeOpen.ICapeThermoPropertyPackage.CalcEquilibrium
            Me.PropertyPackage.CurrentMaterialStream = Me
            Try
                Me.PropertyPackage.CalcEquilibrium(materialObject, flashType, props)
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoPropertyPackage", ex.Source, ex.StackTrace, "CalcEquilibrium", hcode)
            End Try
        End Sub

        Public Sub CalcProp2(ByVal materialObject As Object, ByVal props As Object, ByVal phases As Object, ByVal calcType As String) Implements CapeOpen.ICapeThermoPropertyPackage.CalcProp
            Me.PropertyPackage.CurrentMaterialStream = Me
            Try
                Me.PropertyPackage.CalcProp(materialObject, props, phases, calcType)
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoPropertyPackage", ex.Source, ex.StackTrace, "CalcProp", hcode)
            End Try
        End Sub

        Public Function GetComponentConstant1(ByVal materialObject As Object, ByVal props As Object) As Object Implements CapeOpen.ICapeThermoPropertyPackage.GetComponentConstant
            Me.PropertyPackage.CurrentMaterialStream = Me
            Try
                Return Me.PropertyPackage.GetComponentConstant(materialObject, props)
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoPropertyPackage", ex.Source, ex.StackTrace, "GetComponentConstant", hcode)
                Return Nothing
            End Try
        End Function

        Public Sub GetComponentList(ByRef compIds As Object, ByRef formulae As Object, ByRef names As Object, ByRef boilTemps As Object, ByRef molWt As Object, ByRef casNo As Object) Implements CapeOpen.ICapeThermoPropertyPackage.GetComponentList
            Me.PropertyPackage.CurrentMaterialStream = Me
            Try
                Me.PropertyPackage.GetComponentList(compIds, formulae, names, boilTemps, molWt, casNo)
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoPropertyPackage", ex.Source, ex.StackTrace, "GetComponentList", hcode)
            End Try
        End Sub

        Public Function GetPhaseList1() As Object Implements CapeOpen.ICapeThermoPropertyPackage.GetPhaseList
            Try
                Return Me.PropertyPackage.GetPhaseList()
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoPropertyPackage", ex.Source, ex.StackTrace, "GetPhaseList", hcode)
                Return Nothing
            End Try
        End Function

        Public Function GetPropList2() As Object Implements CapeOpen.ICapeThermoPropertyPackage.GetPropList
            Try
                Return Me.PropertyPackage.GetPropList()
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoPropertyPackage", ex.Source, ex.StackTrace, "GetPropList", hcode)
                Return Nothing
            End Try
        End Function

        Public Function GetUniversalConstant1(ByVal materialObject As Object, ByVal props As Object) As Object Implements CapeOpen.ICapeThermoPropertyPackage.GetUniversalConstant
            Return Me.PropertyPackage.GetUniversalConstant(materialObject, props)
        End Function

        Public Function PropCheck3(ByVal materialObject As Object, ByVal props As Object) As Object Implements CapeOpen.ICapeThermoPropertyPackage.PropCheck
            Try
                Return Me.PropertyPackage.PropCheck(materialObject, props)
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoPropertyPackage", ex.Source, ex.StackTrace, "PropCheck", hcode)
                Return Nothing
            End Try
        End Function

        Public Function ValidityCheck3(ByVal materialObject As Object, ByVal props As Object) As Object Implements CapeOpen.ICapeThermoPropertyPackage.ValidityCheck
            Return Me.PropertyPackage.ValidityCheck(materialObject, props)
        End Function

#End Region

#Region "    CAPE-OPEN 1.1 Thermo & Properties"

        ''' <summary>
        ''' Returns the values of constant Physical Properties for the specified Compounds.
        ''' </summary>
        ''' <param name="props">The list of Physical Property identifiers. Valid
        ''' identifiers for constant Physical Properties are listed in section 7.5.2.</param>
        ''' <param name="compIds">List of Compound identifiers for which constants are to
        ''' be retrieved. Set compIds to UNDEFINED to denote all Compounds in the component that implements the ICapeThermoCompounds interface.</param>
        ''' <returns>Values of constants for the specified Compounds.</returns>
        ''' <remarks>The GetConstPropList method can be used in order to check which constant Physical
        ''' Properties are available.
        ''' If the number of requested Physical Properties is P and the number of Compounds is C, the
        ''' propvals array will contain C*P variants. The first C variants will be the values for the first
        ''' requested Physical Property (one variant for each Compound) followed by C values of constants
        ''' for the second Physical Property, and so on. The actual type of values returned
        ''' (Double, String, etc.) depends on the Physical Property as specified in section 7.5.2.
        ''' Physical Properties are returned in a fixed set of units as specified in section 7.5.2.
        ''' If the compIds argument is set to UNDEFINED this is a request to return property values for
        ''' all compounds in the component that implements the ICapeThermoCompounds interface
        ''' with the compound order the same as that returned by the GetCompoundList method. For
        ''' example, if the interface is implemented by a Property Package component the property
        ''' request with compIds set to UNDEFINED means all compounds in the Property Package
        ''' rather than all compounds in the Material Object passed to the Property package.
        ''' If any Physical Property is not available for one or more Compounds, then undefined values
        ''' must be returned for those combinations and an ECapeThrmPropertyNotAvailable exception
        ''' must be raised. If the exception is raised, the client should check all the values returned to
        ''' determine which is undefined.</remarks>
        Public Function GetCompoundConstant(ByVal props As Object, ByVal compIds As Object) As Object Implements ICapeThermoCompounds.GetCompoundConstant
            Me.PropertyPackage.CurrentMaterialStream = Me
            Try
                Return Me.PropertyPackage.GetCompoundConstant(props, compIds)
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoCompounds", ex.Source, ex.StackTrace, "GetCompoundConstant", hcode)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Returns the list of all Compounds. This includes the Compound identifiers recognised and extra
        '''information that can be used to further identify the Compounds.
        ''' </summary>
        ''' <param name="compIds">List of Compound identifiers</param>
        ''' <param name="formulae">List of Compound formulae</param>
        ''' <param name="names">List of Compound names.</param>
        ''' <param name="boilTemps">List of boiling point temperatures.</param>
        ''' <param name="molwts">List of molecular weights.</param>
        ''' <param name="casnos">List of Chemical Abstract Service (CAS) Registry numbers.</param>
        ''' <remarks>If any item cannot be returned then the value should be set to UNDEFINED. The same information
        ''' can also be extracted using the GetCompoundConstant method. The equivalences
        ''' between GetCompoundList arguments and Compound constant Physical Properties, as
        ''' specified in section 7.5.2, is given in the table below.
        ''' When the ICapeThermoCompounds interface is implemented by a Material Object, the list
        ''' of Compounds returned is fixed when the Material Object is configured.
        ''' For a Property Package component, the Property Package will normally contain a limited set
        ''' of Compounds selected for a particular application, rather than all possible Compounds that
        ''' could be available to a proprietary Properties System.
        ''' The compIds returned by the GetCompoundList method must be unique within the
        ''' component that implements the ICapeThermoCompounds interface. There is no restriction
        ''' on the length of the strings returned in compIds. However, it should be recognised that a
        ''' PME may restrict the length of Compound identifiers internally. In such a case the PME’s
        ''' CAPE-OPEN socket must maintain a method of mapping the, potentially long, identifiers
        ''' used by a CAPE-OPEN Property package component to the identifiers used within the PME.
        ''' In order to identify the Compounds of a Property Package, the PME, or other client, will use
        ''' the casnos argument rather than the compIds. This is because different PMEs and different
        ''' Property Packages may give different names to the same Compounds and the casnos is
        ''' (almost always) unique. If the casnos is not available (e.g. for petroleum fractions), or not
        ''' unique, the other pieces of information returned by GetCompoundList can be used to
        ''' distinguish the Compounds. It should be noted, however, that for communication with a
        ''' Property Package a client must use the Compound identifiers returned in the compIds
        ''' argument. It is the responsibility of the client to maintain appropriate data structures that
        ''' allow it to reconcile the different Compound identifiers used by different Property Packages
        ''' and any native property system.</remarks>
        Public Sub GetCompoundList(ByRef compIds As Object, ByRef formulae As Object, ByRef names As Object, ByRef boilTemps As Object, ByRef molwts As Object, ByRef casnos As Object) Implements ICapeThermoCompounds.GetCompoundList
            Me.PropertyPackage.CurrentMaterialStream = Me
            Try
                Me.PropertyPackage.GetCompoundList(compIds, formulae, names, boilTemps, molwts, casnos)
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoCompounds", ex.Source, ex.StackTrace, "GetCompoundList", hcode)
            End Try
        End Sub

        ''' <summary>
        ''' Returns the list of supported constant Physical Properties.
        ''' </summary>
        ''' <returns>List of identifiers for all supported constant Physical Properties. The standard constant property identifiers are listed in section 7.5.2.</returns>
        ''' <remarks>GetConstPropList returns identifiers for all the constant Physical Properties that can be
        ''' retrieved by the GetCompoundConstant method. If no properties are supported,
        ''' UNDEFINED should be returned. The CAPE-OPEN standards do not define a minimum list
        ''' of Physical Properties to be made available by a software component that implements the
        ''' ICapeThermoCompounds interface.
        ''' A component that implements the ICapeThermoCompounds interface may return constant
        ''' Physical Property identifiers which do not belong to the list defined in section 7.5.2.
        ''' However, these proprietary identifiers may not be understood by most of the clients of this
        ''' component.</remarks>
        Public Function GetConstPropList() As Object Implements ICapeThermoCompounds.GetConstPropList
            Me.PropertyPackage.CurrentMaterialStream = Me
            Try
                Return Me.PropertyPackage.GetConstPropList()
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoCompounds", ex.Source, ex.StackTrace, "GetConstPropList", hcode)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Returns the number of Compounds supported.
        ''' </summary>
        ''' <returns>Number of Compounds supported.</returns>
        ''' <remarks>The number of Compounds returned by this method must be equal to the number of
        ''' Compound identifiers that are returned by the GetCompoundList method of this interface. It
        ''' must be zero or a positive number.</remarks>
        Public Function GetNumCompounds() As Integer Implements ICapeThermoCompounds.GetNumCompounds
            Return Me.Fases(0).Componentes.Count
        End Function

        ''' <summary>
        ''' Returns the values of pressure-dependent Physical Properties for the specified pure Compounds.
        ''' </summary>
        ''' <param name="props">The list of Physical Property identifiers. Valid identifiers for pressure-dependent 
        ''' Physical Properties are listed in section 7.5.4</param>
        ''' <param name="pressure">Pressure (in Pa) at which Physical Properties are evaluated</param>
        ''' <param name="compIds">List of Compound identifiers for which Physical Properties are to be retrieved. 
        ''' Set compIds to UNDEFINED to denote all Compounds in the component that implements the ICapeThermoCompounds interface.</param>
        ''' <param name="propVals">Property values for the Compounds specified.</param>
        ''' <remarks></remarks>
        Public Sub GetPDependentProperty(ByVal props As Object, ByVal pressure As Double, ByVal compIds As Object, ByRef propVals As Object) Implements ICapeThermoCompounds.GetPDependentProperty
            Me.PropertyPackage.CurrentMaterialStream = Me
            Try
                Me.PropertyPackage.GetPDependentProperty(props, pressure, compIds, propVals)
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoCompounds", ex.Source, ex.StackTrace, "GetPDependentProperty", hcode)
            End Try
        End Sub

        ''' <summary>
        ''' Returns the list of supported pressure-dependent properties.
        ''' </summary>
        ''' <returns>The list of Physical Property identifiers for all supported pressure-dependent properties. The standard identifiers are listed in section 7.5.4</returns>
        ''' <remarks>GetPDependentPropList returns identifiers for all the pressure-dependent properties that can
        ''' be retrieved by the GetPDependentProperty method. If no properties are supported
        ''' UNDEFINED should be returned. The CAPE-OPEN standards do not define a minimum list
        ''' of Physical Properties to be made available by a software component that implements the
        ''' ICapeThermoCompounds interface.
        ''' A component that implements the ICapeThermoCompounds interface may return identifiers
        ''' which do not belong to the list defined in section 7.5.4. However, these proprietary
        ''' identifiers may not be understood by most of the clients of this component.</remarks>
        Public Function GetPDependentPropList() As Object Implements ICapeThermoCompounds.GetPDependentPropList
            Me.PropertyPackage.CurrentMaterialStream = Me
            Try
                Return Me.PropertyPackage.GetPDependentPropList()
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoCompounds", ex.Source, ex.StackTrace, "GetPDependentList", hcode)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Returns the values of temperature-dependent Physical Properties for the specified pure Compounds.
        ''' </summary>
        ''' <param name="props">The list of Physical Property identifiers. Valid identifiers for 
        ''' temperature-dependent Physical Properties are listed in section 7.5.3</param>
        ''' <param name="temperature">Temperature (in K) at which properties are evaluated</param>
        ''' <param name="compIds">List of Compound identifiers for which Physical Properties are to be retrieved. 
        ''' Set compIds to UNDEFINED to denote all Compounds in the component that implements the ICapeThermoCompounds interface.</param>
        ''' <param name="propVals">Physical Property values for the Compounds specified.</param>
        ''' <remarks>The GetTDependentPropList method can be used in order to check which Physical
        ''' Properties are available.
        ''' If the number of requested Physical Properties is P and the number of Compounds is C, the
        ''' propvals array will contain C*P values. The first C will be the values for the first requested
        ''' Physical Property followed by C values for the second Physical Property, and so on.
        ''' Properties are returned in a fixed set of units as specified in section 7.5.3.
        ''' If the compIds argument is set to UNDEFINED this is a request to return property values for
        ''' all compounds in the component that implements the ICapeThermoCompounds interface
        ''' with the compound order the same as that returned by the GetCompoundList method. For
        ''' example, if the interface is implemented by a Property Package component the property
        ''' request with compIds set to UNDEFINED means all compounds in the Property Package
        ''' rather than all compounds in the Material Object passed to the Property package.
        ''' If any Physical Property is not available for one or more Compounds, then undefined values
        ''' must be returned for those combinations and an ECapeThrmPropertyNotAvailable exception
        ''' must be raised. If the exception is raised, the client should check all the values returned to
        ''' determine which is undefined.</remarks>
        Public Sub GetTDependentProperty(ByVal props As Object, ByVal temperature As Double, ByVal compIds As Object, ByRef propVals As Object) Implements ICapeThermoCompounds.GetTDependentProperty
            Me.PropertyPackage.CurrentMaterialStream = Me
            Try
                Me.PropertyPackage.GetTDependentProperty(props, temperature, compIds, propVals)
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoCompounds", ex.Source, ex.StackTrace, "GetTDependentProperty", hcode)
            End Try
        End Sub

        ''' <summary>
        ''' Returns the list of supported temperature-dependent Physical Properties.
        ''' </summary>
        ''' <returns>The list of Physical Property identifiers for all supported temperature-dependent 
        ''' properties. The standard identifiers are listed in section 7.5.3</returns>
        ''' <remarks>GetTDependentPropList returns identifiers for all the temperature-dependent Physical
        ''' Properties that can be retrieved by the GetTDependentProperty method. If no properties are
        ''' supported UNDEFINED should be returned. The CAPE-OPEN standards do not define a
        ''' minimum list of properties to be made available by a software component that implements
        ''' the ICapeThermoCompounds interface.
        ''' A component that implements the ICapeThermoCompounds interface may return identifiers
        ''' which do not belong to the list defined in section 7.5.3. However, these proprietary identifiers
        ''' may not be understood by most of the clients of this component.</remarks>
        Public Function GetTDependentPropList() As Object Implements ICapeThermoCompounds.GetTDependentPropList
            Me.PropertyPackage.CurrentMaterialStream = Me
            Try
                Return Me.PropertyPackage.GetTDependentPropList()
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoCompounds", ex.Source, ex.StackTrace, "GetTDependentPropList", hcode)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Remove all stored Physical Property values.
        ''' </summary>
        ''' <remarks>ClearAllProps removes all stored Physical Properties that have been set using the
        ''' SetSinglePhaseProp, SetTwoPhaseProp or SetOverallProp methods. This means that any
        ''' subsequent call to retrieve Physical Properties will result in an exception until new values
        ''' have been stored using one of the Set methods. ClearAllProps does not remove the
        ''' configuration information for a Material, i.e. the list of Compounds and Phases.
        ''' Using the ClearAllProps method results in a Material Object that is in the same state as
        ''' when it was first created. It is an alternative to using the CreateMaterial method but it is
        ''' expected to have a smaller overhead in operating system resources.</remarks>
        Public Sub ClearAllProps() Implements ICapeThermoMaterial.ClearAllProps
            Me.PropertyPackage.CurrentMaterialStream = Me
            Me.PropertyPackage.DW_ZerarPhaseProps(PropertyPackages.Fase.Vapor)
            Me.PropertyPackage.DW_ZerarPhaseProps(PropertyPackages.Fase.Liquid)
            Me.PropertyPackage.DW_ZerarPhaseProps(PropertyPackages.Fase.Liquid1)
            Me.PropertyPackage.DW_ZerarPhaseProps(PropertyPackages.Fase.Liquid2)
            Me.PropertyPackage.DW_ZerarPhaseProps(PropertyPackages.Fase.Liquid3)
            Me.PropertyPackage.DW_ZerarPhaseProps(PropertyPackages.Fase.Aqueous)
            Me.PropertyPackage.DW_ZerarPhaseProps(PropertyPackages.Fase.Solid)
            Me.PropertyPackage.DW_ZerarPhaseProps(PropertyPackages.Fase.Mixture)
            Me.PropertyPackage.DW_ZerarComposicoes(PropertyPackages.Fase.Vapor)
            Me.PropertyPackage.DW_ZerarComposicoes(PropertyPackages.Fase.Liquid)
            Me.PropertyPackage.DW_ZerarComposicoes(PropertyPackages.Fase.Liquid1)
            Me.PropertyPackage.DW_ZerarComposicoes(PropertyPackages.Fase.Liquid2)
            Me.PropertyPackage.DW_ZerarComposicoes(PropertyPackages.Fase.Liquid3)
            Me.PropertyPackage.DW_ZerarComposicoes(PropertyPackages.Fase.Aqueous)
            Me.PropertyPackage.DW_ZerarComposicoes(PropertyPackages.Fase.Solid)
            Me.PropertyPackage.DW_ZerarComposicoes(PropertyPackages.Fase.Mixture)
        End Sub

        ''' <summary>
        ''' Copies all the stored non-constant Physical Properties (which have been set using the SetSinglePhaseProp, 
        ''' SetTwoPhaseProp or SetOverallProp) from the source Material Object to the current instance of the Material Object.</summary>
        ''' <param name="source">Source Material Object from which stored properties will be copied.</param>
        ''' <remarks>Before using this method, the Material Object must have been configured with the same
        ''' exact list of Compounds and Phases as the source one. Otherwise, calling the method will
        ''' raise an exception. There are two ways to perform the configuration: through the PME
        ''' proprietary mechanisms and with CreateMaterial. Calling CreateMaterial on a Material
        ''' Object S and subsequently calling CopyFromMaterial(S) on the newly created Material
        ''' Object N is equivalent to the deprecated method ICapeMaterialObject.Duplicate.
        ''' The method is intended to be used by a client, for example a Unit Operation that needs a
        ''' Material Object to have the same state as one of the Material Objects it has been connected
        ''' to. One example is the representation of an internal stream in a distillation column.
        ''' If the Material Object supports the Petroleum Fractions Interface [7] the petroleum fraction
        ''' properties are also copied from the source Material Object to the current instance of the
        ''' Material Object.</remarks>
        Public Sub CopyFromMaterial(ByRef source As Object) Implements ICapeThermoMaterial.CopyFromMaterial

            If Not Marshal.IsComObject(source) Then
                Me.Assign(source)
                Me.AssignProps(source)
            End If

        End Sub

        ''' <summary>
        ''' Creates a Material Object with the same configuration as the current Material Object.
        ''' </summary>
        ''' <returns>The Material Object created does not contain any non-constant Physical Property value but
        ''' has the same configuration (Compounds and Phases) as the current Material Object. These
        ''' Physical Property values must be set using SetSinglePhaseProp, SetTwoPhaseProp or
        ''' SetOverallProp. Any attempt to retrieve Physical Property values before they have been set
        ''' will result in an exception.</returns>
        ''' <remarks></remarks>
        Public Function CreateMaterial() As Object Implements ICapeThermoMaterial.CreateMaterial
            Dim mat As Streams.MaterialStream = Me.Clone
            mat.ClearAllProps()
            Return mat
        End Function

        ''' <summary>
        ''' Retrieves non-constant Physical Property values for the overall mixture.
        ''' </summary>
        ''' <param name="property">The identifier of the Physical Property for which values are requested. 
        ''' This must be one of the single-phase Physical Properties or derivatives that can be stored for 
        ''' the overall mixture. The standard identifiers are listed in sections 7.5.5 and 7.6.</param>
        ''' <param name="basis">Basis of the results. Valid settings are: “Mass” for Physical Properties per 
        ''' unit mass or “Mole” for molar properties. Use UNDEFINED as a place holder for a Physical Property 
        ''' for which basis does not apply. See section 7.5.5 for details.</param>
        ''' <param name="results">Results vector containing Physical Property value(s) in SI units.</param>
        ''' <remarks>The Physical Property values returned by GetOverallProp refer to the overall mixture. These
        ''' values are set by calling the SetOverallProp method. Overall mixture Physical Properties are
        ''' not calculated by components that implement the ICapeThermoMaterial interface. The
        ''' property values are only used as input specifications for the CalcEquilibrium method of a
        ''' component that implements the ICapeThermoEquilibriumRoutine interface.
        ''' It is expected that this method will normally be able to provide Physical Property values on
        ''' any basis, i.e. it should be able to convert values from the basis on which they are stored to
        ''' the basis requested. This operation will not always be possible. For example, if the
        ''' molecular weight is not known for one or more Compounds, it is not possible to convert
        ''' between a mass basis and a molar basis.
        ''' Although the result of some calls to GetOverallProp will be a single value, the return type is
        ''' CapeArrayDouble and the method must always return an array even if it contains only a
        ''' single element.</remarks>
        Public Sub GetOverallProp(ByVal [property] As String, ByVal basis As String, ByRef results As Object) Implements ICapeThermoMaterial.GetOverallProp
            Try
                GetSinglePhaseProp([property], "overall", basis, results)
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoMaterial", ex.Source, ex.StackTrace, "GetOverallProp", hcode)
            End Try
        End Sub

        ''' <summary>
        ''' Retrieves temperature, pressure and composition for the overall mixture.
        ''' </summary>
        ''' <param name="temperature">Temperature (in K)</param>
        ''' <param name="pressure">Pressure (in Pa)</param>
        ''' <param name="composition">Composition (mole fractions)</param>
        ''' <remarks>This method is provided to make it easier for developers to make efficient use of the CAPEOPEN interfaces. 
        ''' It returns the most frequently requested information from a Material Object in a single call. There is no choice 
        ''' of basis in this method. The composition is always returned as mole fractions.</remarks>
        Public Sub GetOverallTPFraction(ByRef temperature As Double, ByRef pressure As Double, ByRef composition As Object) Implements ICapeThermoMaterial.GetOverallTPFraction
            Try
                Me.GetTPFraction("overall", temperature, pressure, composition)
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoMaterial", ex.Source, ex.StackTrace, "GetOverallTPFraction", hcode)
            End Try
        End Sub

        ''' <summary>
        ''' Returns Phase labels for the Phases that are currently present in the Material Object.
        ''' </summary>
        ''' <param name="phaseLabels">The list of Phase labels (identifiers – names) for the Phases present 
        ''' in the Material Object. The Phase labels in the Material Object must be a subset of the labels 
        ''' returned by the GetPhaseList method of the ICapeThermoPhases interface.</param>
        ''' <param name="phaseStatus"></param>
        ''' <remarks>A Phase is ‘present’ in a Material Object (or other component that implements the
        ''' ICapeThermoMaterial interface) if it has been explicitly made present by calling the
        ''' SetPresentPhases method or if any properties have been set by calling the
        ''' SetSinglePhaseProp or SetTwoPhaseProp methods. Even if a Phase is present, it does not
        ''' necessarily imply that any Physical Properties are actually set unless the phaseStatus is
        ''' Cape_AtEquilibrium or Cape_Estimates (see below). Note that calling the SetPresentPhases
        ''' method of the ICapeThermoMaterial interface will cause any phases not specified in its
        ''' phaseLabels list to not present, even if previously present as a result of a SetSingle-
        ''' PhaseProp or SetTwoPhaseProp call.
        ''' If no Phases are present, UNDEFINED should be returned for both the phaseLabels and
        ''' phaseStatus arguments.
        ''' The phaseStatus argument contains as many entries as there are Phase labels. The valid
        ''' settings are listed in the following table:
        ''' 
        ''' Identifier                    Meaning
        ''' Cape_UnknownPhaseStatus       This is the normal setting when a Phase is specified as being available for an Equilibrium Calculation.
        ''' Cape_AtEquilibrium            The Phase has been set as present as a result of an Equilibrium Calculation.
        ''' Cape_Estimates                Estimates of the equilibrium state have been set in the Material Object.
        ''' 
        ''' All the Phases with a status of Cape_AtEquilibrium have values of temperature, pressure,
        ''' composition and Phase fraction set that correspond to an equilibrium state, i.e. equal
        ''' temperature, pressure and fugacities of each Compound. Phases with a Cape_Estimates
        ''' status have values of temperature, pressure, composition and Phase fraction set in the
        ''' Material Object. These values are available for use by an Equilibrium Calculator component
        ''' to initialise an Equilibrium Calculation. The stored values are available but there is no
        ''' guarantee that they will be used.
        ''' GetPresentPhases is intended to be used in several contexts.
        ''' A Property Package, Property Calculator or other PMC may use this method to check
        ''' whether a phase is present in the Material Object prior to requesting and/or calculating
        ''' some properties.
        ''' An Equilibrium Calculator component will use this method to obtain the list of phases to
        ''' consider in an equilibrium calculation or when checking an equilibrium specification
        ''' (see below for more details).
        ''' The method will be used by the PME or PMC to obtain the list of phases present as the
        ''' result of an equilibrium calculation (see below for more details).
        ''' A Unit Operation (or other PMC) will use this method to get the list of phases present at
        ''' an inlet port or during its calculations.
        ''' In the context of Equilibrium Calculations the GetPresentPhases method is intended to work
        ''' in conjunction with the SetPresentPhases method. Together these methods provide a means
        ''' of communication between a PME (or another client) and an Equilibrium Calculator (or
        ''' other component that implements the ICapeThermoEquilibriumRoutine interface). The
        ''' following sequence of operations is envisaged.
        ''' 
        ''' 1. Prior to requesting an Equilibrium Calculation, a PME will use the SetPresentPhases
        ''' method to define a list of Phases that may be considered in the Equilibrium
        ''' Calculation. Typically, this is necessary because an Equilibrium Calculator may be
        ''' capable of handling a large number of Phases but for a particular application, it may
        ''' be known that only certain Phases will be involved. For example, if the complete
        ''' Phase list contains Phases with the following labels (with the obvious interpretation):
        ''' vapour, hydrocarbonLiquid and aqueousLiquid and it is required to model a liquid
        ''' decanter, the present Phases might be set to hydrocarbonLiquid and aqueousLiquid.
        ''' 
        ''' 2. The GetPresentPhases method is then used by the CalcEquilibrium method of the
        ''' ICapeThermoEquilibriumRoutine interface to obtain the list of Phase labels corresponding
        ''' to the Phases that may be present at equilibrium.
        ''' 
        ''' 3. The Equilibrium Calculation determines which Phases actually co-exist at
        ''' equilibrium. This list of Phases may be a sub-set of the Phases considered because
        ''' some Phases may not be present at the prevailing conditions. For example, if the
        ''' amount of water is sufficiently small the aqueousLiquid Phase in the above example
        ''' may not exist because all the water dissolves in the hydrocarbonLiquid Phase.
        ''' 
        ''' 4. The CalcEquilibrium method uses the SetPresentPhases method to indicate the
        ''' Phases present following the equilibrium calculation (and sets the phase properties).
        ''' 
        ''' 5. The PME uses the GetPresentPhases method to find out the Phases present following
        ''' the calculation and it can then use the GetSinglePhaseProp or GetTPFraction
        ''' methods to get the Phase properties.</remarks>
        Public Sub GetPresentPhases(ByRef phaseLabels As Object, ByRef phaseStatus As Object) Implements ICapeThermoMaterial.GetPresentPhases

            Dim pl As New ArrayList, stat As New ArrayList

            If AtEquilibrium Then
                For Each pi As PhaseInfo In Me.PropertyPackage.PhaseMappings.Values
                    If pi.PhaseLabel <> "Disabled" Then
                        If Me.Fases(pi.DWPhaseIndex).SPMProperties.molarfraction.HasValue Then
                            pl.Add(pi.PhaseLabel)
                            stat.Add(CapeOpen.eCapePhaseStatus.CAPE_ATEQUILIBRIUM)
                        End If
                    End If
                Next
            Else
                For Each pi As PhaseInfo In Me.PropertyPackage.PhaseMappings.Values
                    If pi.PhaseLabel <> "Disabled" Then
                        pl.Add(pi.PhaseLabel)
                        stat.Add(CapeOpen.eCapePhaseStatus.CAPE_UNKNOWNPHASESTATUS)
                    End If
                Next
            End If

            Dim arr(pl.Count - 1) As String
            Array.Copy(pl.ToArray, arr, pl.Count)
            phaseLabels = arr

            Dim arr2(stat.Count - 1) As CapeOpen.eCapePhaseStatus
            Array.Copy(stat.ToArray, arr2, stat.Count)
            phaseStatus = arr2

        End Sub

        ''' <summary>
        ''' Retrieves single-phase non-constant Physical Property values for a mixture.
        ''' </summary>
        ''' <param name="property">The identifier of the Physical Property for which values are requested. 
        ''' This must be one of the single-phase Physical Properties or derivatives. The standard identifiers 
        ''' are listed in sections 7.5.5 and 7.6.</param>
        ''' <param name="phaseLabel">Phase label of the Phase for which the Physical Property is required. 
        ''' The Phase label must be one of the identifiers returned by the GetPresentPhases method of this interface.</param>
        ''' <param name="basis">Basis of the results. Valid settings are: “Mass” for Physical Properties per unit mass or “Mole” 
        ''' for molar properties. Use UNDEFINED as a place holder for a Physical Property for which basis does not apply. See section 
        ''' 7.5.5 for details.</param>
        ''' <param name="results">Results vector (CapeArrayDouble) containing Physical Property value(s) in SI units or CapeInterface 
        ''' (see notes).</param>
        ''' <remarks>The results argument returned by GetSinglePhaseProp is either a CapeArrayDouble that
        ''' contains one or more numerical values, e.g. temperature, or a CapeInterface that may be
        ''' used to retrieve single-phase Physical Properties described by a more complex data
        ''' structure, e.g. distributed properties.
        ''' It is required that a component that implements the ICapeThermoMaterial interface will
        ''' always support the following properties: temperature, pressure, fraction, phaseFraction,
        ''' flow, totalFlow.
        ''' Although the result of some calls to GetSinglePhaseProp may be a single numerical value,
        ''' the return type for numerical values is CapeArrayDouble and in such a case the method must
        ''' return an array even if it contains only a single element.
        ''' A Phase is ‘present’ in a Material if its identifier is returned by the GetPresentPhases
        ''' method. An exception is raised by the GetSinglePhaseProp method if the Phase specified is
        ''' not present. Even if a Phase is present, this does not necessarily mean that any Physical
        ''' Properties are available.
        ''' The Physical Property values returned by GetSinglePhaseProp refer to a single Phase. These
        ''' values may be set by the SetSinglePhaseProp method, which may be called directly, or by
        ''' ICapeThermoPropertyRoutine interface or the CalcEquilibrium method of the
        ''' ICapeThermoEquilibriumRoutine interface. Note: Physical Properties that depend on more
        ''' than one Phase, for example surface tension or K-values, are returned by the
        ''' GetTwoPhaseProp method.
        ''' It is expected that this method will normally be able to provide Physical Property values on
        ''' any basis, i.e. it should be able to convert values from the basis on which they are stored to
        ''' the basis requested. This operation will not always be possible. For example, if the
        ''' molecular weight is not known for one or more Compounds, it is not possible to convert
        ''' from mass fractions or mass flows to mole fractions or molar flows.</remarks>
        Public Sub GetSinglePhaseProp(ByVal [property] As String, ByVal phaseLabel As String, ByVal basis As String, ByRef results As Object) Implements ICapeThermoMaterial.GetSinglePhaseProp

            Me.PropertyPackage.CurrentMaterialStream = Me

            Dim res As New ArrayList
            Dim comps As New ArrayList
            For Each c As Substancia In Me.Fases(0).Componentes.Values
                comps.Add(c.Nome)
            Next
            Dim f As Integer = -1
            Dim phs As DTL.SimulationObjects.PropertyPackages.Fase
            Select Case phaseLabel.ToLower
                Case "overall"
                    f = 0
                    phs = PropertyPackages.Fase.Mixture
                Case Else
                    For Each pi As PhaseInfo In Me.PropertyPackage.PhaseMappings.Values
                        If phaseLabel = pi.PhaseLabel Then
                            f = pi.DWPhaseIndex
                            phs = pi.DWPhaseID
                            Exit For
                        End If
                    Next
            End Select

            If f = -1 Then
                Dim ex As New ArgumentException("Invalid Phase ID")
                Dim hcode As Integer = 0
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoMaterial", ex.Source, ex.StackTrace, "GetOverallTPFraction", hcode)
            End If

            Select Case [property].ToLower
                Case "compressibilityfactor"
                    res.Add(Me.Fases(f).SPMProperties.compressibilityFactor.GetValueOrDefault)
                Case "heatofvaporization"
                Case "heatcapacity", "heatcapacitycp"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            res.Add(Me.Fases(f).SPMProperties.heatCapacityCp.GetValueOrDefault * Me.PropertyPackage.AUX_MMM(phs))
                        Case "Mass", "mass"
                            res.Add(Me.Fases(f).SPMProperties.heatCapacityCp.GetValueOrDefault * 1000)
                    End Select
                Case "heatcapacitycv"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            res.Add(Me.Fases(f).SPMProperties.heatCapacityCv.GetValueOrDefault * Me.PropertyPackage.AUX_MMM(phs))
                        Case "Mass", "mass"
                            res.Add(Me.Fases(f).SPMProperties.heatCapacityCv.GetValueOrDefault * 1000)
                    End Select
                Case "idealgasheatcapacity"
                    If f = 1 Then
                        res.Add(Me.PropertyPackage.AUX_CPm(PropertyPackages.Fase.Liquid, Me.Fases(0).SPMProperties.temperature * 1000))
                    Else
                        res.Add(Me.PropertyPackage.AUX_CPm(PropertyPackages.Fase.Vapor, Me.Fases(0).SPMProperties.temperature * 1000))
                    End If
                Case "idealgasenthalpy"
                    If f = 1 Then
                        res.Add(Me.PropertyPackage.RET_Hid(298.15, Me.Fases(0).SPMProperties.temperature.GetValueOrDefault * 1000, PropertyPackages.Fase.Liquid))
                    Else
                        res.Add(Me.PropertyPackage.RET_Hid(298.15, Me.Fases(0).SPMProperties.temperature.GetValueOrDefault * 1000, PropertyPackages.Fase.Vapor))
                    End If
                Case "excessenthalpy"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            res.Add(Me.Fases(f).SPMProperties.excessEnthalpy.GetValueOrDefault * Me.PropertyPackage.AUX_MMM(phs))
                        Case "Mass", "mass"
                            res.Add(Me.Fases(f).SPMProperties.excessEnthalpy.GetValueOrDefault * 1000)
                    End Select
                Case "excessentropy"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            res.Add(Me.Fases(f).SPMProperties.excessEntropy.GetValueOrDefault * Me.PropertyPackage.AUX_MMM(phs))
                        Case "Mass", "mass"
                            res.Add(Me.Fases(f).SPMProperties.excessEntropy.GetValueOrDefault * 1000)
                    End Select
                Case "viscosity"
                    res.Add(Me.Fases(f).SPMProperties.viscosity.GetValueOrDefault)
                Case "thermalconductivity"
                    res.Add(Me.Fases(f).SPMProperties.thermalConductivity.GetValueOrDefault)
                Case "fugacity"
                    For Each c As String In comps
                        res.Add(Me.Fases(f).Componentes(c).FracaoMolar.GetValueOrDefault * Me.Fases(f).Componentes(c).FugacityCoeff.GetValueOrDefault * Me.Fases(0).SPMProperties.pressure.GetValueOrDefault)
                    Next
                Case "activity"
                    For Each c As String In comps
                        res.Add(Me.Fases(f).Componentes(c).ActivityCoeff.GetValueOrDefault * Me.Fases(f).Componentes(c).FracaoMolar.GetValueOrDefault)
                    Next
                Case "fugacitycoefficient"
                    For Each c As String In comps
                        res.Add(Me.Fases(f).Componentes(c).FugacityCoeff.GetValueOrDefault)
                    Next
                Case "activitycoefficient"
                    For Each c As String In comps
                        res.Add(Me.Fases(f).Componentes(c).ActivityCoeff.GetValueOrDefault)
                    Next
                Case "logfugacitycoefficient"
                    For Each c As String In comps
                        res.Add(Math.Log(Me.Fases(f).Componentes(c).FugacityCoeff.GetValueOrDefault))
                    Next
                Case "volume"
                    res.Add(Me.Fases(f).SPMProperties.molecularWeight / Me.Fases(f).SPMProperties.density / 1000)
                Case "density"
                    res.Add(Me.Fases(f).SPMProperties.density.GetValueOrDefault)
                Case "enthalpy", "enthalpynf"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Dim val = Me.Fases(f).SPMProperties.molecularWeight.GetValueOrDefault
                            If val = 0.0# Then
                                res.Add(Me.Fases(f).SPMProperties.molar_enthalpy.GetValueOrDefault)
                            Else
                                res.Add(Me.Fases(f).SPMProperties.enthalpy.GetValueOrDefault * val)
                            End If
                        Case "Mass", "mass"
                            res.Add(Me.Fases(f).SPMProperties.enthalpy.GetValueOrDefault * 1000)
                    End Select
                Case "entropy", "entropynf"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Dim val = Me.Fases(f).SPMProperties.molecularWeight.GetValueOrDefault
                            If val = 0.0# Then
                                res.Add(Me.Fases(f).SPMProperties.molar_entropy.GetValueOrDefault)
                            Else
                                res.Add(Me.Fases(f).SPMProperties.entropy.GetValueOrDefault * val)
                            End If
                        Case "Mass", "mass"
                            res.Add(Me.Fases(f).SPMProperties.entropy.GetValueOrDefault * 1000)
                    End Select
                Case "enthalpyf"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Dim val = Me.Fases(f).SPMProperties.molecularWeight.GetValueOrDefault
                            If val = 0.0# Then
                                res.Add(Me.Fases(f).SPMProperties.molar_enthalpyF.GetValueOrDefault)
                            Else
                                res.Add(Me.Fases(f).SPMProperties.enthalpyF.GetValueOrDefault * val)
                            End If
                        Case "Mass", "mass"
                            res.Add(Me.Fases(f).SPMProperties.enthalpyF.GetValueOrDefault * 1000)
                    End Select
                Case "entropyf"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Dim val = Me.Fases(f).SPMProperties.molecularWeight.GetValueOrDefault
                            If val = 0.0# Then
                                res.Add(Me.Fases(f).SPMProperties.molar_entropyF.GetValueOrDefault)
                            Else
                                res.Add(Me.Fases(f).SPMProperties.entropyF.GetValueOrDefault * val)
                            End If
                        Case "Mass", "mass"
                            res.Add(Me.Fases(f).SPMProperties.entropy.GetValueOrDefault * 1000)
                    End Select
                Case "moles"
                    res.Add(Me.Fases(f).SPMProperties.molarflow.GetValueOrDefault)
                Case "mass"
                    res.Add(Me.Fases(f).SPMProperties.massflow.GetValueOrDefault)
                Case "molecularweight"
                    res.Add(Me.Fases(f).SPMProperties.molecularWeight.GetValueOrDefault)
                Case "temperature"
                    res.Add(Me.Fases(0).SPMProperties.temperature.GetValueOrDefault)
                Case "pressure"
                    res.Add(Me.Fases(0).SPMProperties.pressure.GetValueOrDefault)
                Case "flow"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            For Each c As String In comps
                                res.Add(Me.Fases(f).Componentes(c).MolarFlow.GetValueOrDefault)
                            Next
                        Case "Mass", "mass"
                            For Each c As String In comps
                                res.Add(Me.Fases(f).Componentes(c).MassFlow.GetValueOrDefault)
                            Next
                    End Select
                Case "fraction", "massfraction", "molarfraction"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            For Each c As String In comps
                                res.Add(Me.Fases(f).Componentes(c).FracaoMolar.GetValueOrDefault)
                            Next
                        Case "Mass", "mass"
                            For Each c As String In comps
                                res.Add(Me.Fases(f).Componentes(c).FracaoMassica.GetValueOrDefault)
                            Next
                        Case ""
                            If [property].ToLower.Contains("mole") Then
                                For Each c As String In comps
                                    res.Add(Me.Fases(f).Componentes(c).FracaoMolar.GetValueOrDefault)
                                Next
                            ElseIf [property].ToLower.Contains("mass") Then
                                For Each c As String In comps
                                    res.Add(Me.Fases(f).Componentes(c).FracaoMassica.GetValueOrDefault)
                                Next
                            End If
                    End Select
                Case "concentration"
                    For Each c As String In comps
                        res.Add(Me.Fases(f).Componentes(c).MassFlow.GetValueOrDefault / Me.Fases(f).SPMProperties.volumetric_flow.GetValueOrDefault)
                    Next
                Case "molarity"
                    For Each c As String In comps
                        res.Add(Me.Fases(f).Componentes(c).MolarFlow.GetValueOrDefault / Me.Fases(f).SPMProperties.volumetric_flow.GetValueOrDefault)
                    Next
                Case "phasefraction"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            res.Add(Me.Fases(f).SPMProperties.molarfraction.GetValueOrDefault)
                        Case "Mass", "mass"
                            res.Add(Me.Fases(f).SPMProperties.massfraction.GetValueOrDefault)
                    End Select
                Case "totalflow"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            res.Add(Me.Fases(f).SPMProperties.molarflow.GetValueOrDefault)
                        Case "Mass", "mass"
                            res.Add(Me.Fases(f).SPMProperties.massflow.GetValueOrDefault)
                    End Select
                Case Else
                    Dim ex = New Exception
                    Dim hcode As Integer = 0
                    ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoMaterial", ex.Source, ex.StackTrace, "GetSinglePhaseProp", hcode)
            End Select
            Dim arr(res.Count - 1) As Double
            Array.Copy(res.ToArray, arr, res.Count)
            results = arr
        End Sub

        ''' <summary>
        ''' Retrieves temperature, pressure and composition for a Phase.
        ''' </summary>
        ''' <param name="phaseLabel">Phase label of the Phase for which the property is required. The Phase label 
        ''' must be one of the identifiers returned by the GetPresentPhases method of this interface.</param>
        ''' <param name="temperature">Temperature (in K)</param>
        ''' <param name="pressure">Pressure (in Pa)</param>
        ''' <param name="composition">Composition (mole fractions)</param>
        ''' <remarks>This method is provided to make it easier for developers to make efficient use of the CAPEOPEN
        ''' interfaces. It returns the most frequently requested information from a Material
        ''' Object in a single call.
        ''' There is no choice of basis in this method. The composition is always returned as mole
        ''' fractions.
        ''' To get the equivalent information for the overall mixture the GetOverallTPFraction method
        ''' of the ICapeThermoMaterial interface should be used.</remarks>
        Public Sub GetTPFraction(ByVal phaseLabel As String, ByRef temperature As Double, ByRef pressure As Double, ByRef composition As Object) Implements ICapeThermoMaterial.GetTPFraction

            If Me.Fases(0).SPMProperties.temperature Is Nothing Or Me.Fases(0).SPMProperties.pressure Is Nothing Then
                Throw New ArgumentException("Temperature and/or Pressure not set.")
            End If

            temperature = Me.Fases(0).SPMProperties.temperature
            pressure = Me.Fases(0).SPMProperties.pressure

            Dim arr As New ArrayList
            Dim comps As New ArrayList
            For Each c As Substancia In Me.Fases(0).Componentes.Values
                comps.Add(c.Nome)
            Next
            Select Case phaseLabel.ToLower
                Case "overall"
                    For Each c As String In comps
                        arr.Add(Me.Fases(0).Componentes(c).FracaoMolar)
                    Next
                Case Else
                    For Each pi As PhaseInfo In Me.PropertyPackage.PhaseMappings.Values
                        If phaseLabel = pi.PhaseLabel Then
                            For Each c As String In comps
                                arr.Add(Me.Fases(pi.DWPhaseIndex).Componentes(c).FracaoMolar)
                            Next
                            Exit For
                        End If
                    Next
            End Select
            Dim arr2(arr.Count - 1) As Double
            Array.Copy(arr.ToArray, arr2, arr.Count)
            composition = arr2
        End Sub

        ''' <summary>
        ''' Retrieves two-phase non-constant Physical Property values for a mixture.
        ''' </summary>
        ''' <param name="property">The identifier of the property for which values are requested. This must 
        ''' be one of the two-phase Physical Properties or Physical Property derivatives listed in sections 
        ''' 7.5.6 and 7.6.</param>
        ''' <param name="phaseLabels">List of Phase labels of the Phases for which the property is required. 
        ''' The Phase labels must be two of the identifiers returned by the GetPhaseList method of the Material 
        ''' Object.</param>
        ''' <param name="basis">Basis of the results. Valid settings are: “Mass” for Physical Properties per unit 
        ''' mass or “Mole” for molar properties. Use UNDEFINED as a place holder for a Physical Property for which 
        ''' basis does not apply. See section 7.5.5 for details.</param>
        ''' <param name="results">Results vector (CapeArrayDouble) containing property value(s) in SI units or 
        ''' CapeInterface (see notes).</param>
        ''' <remarks>The results argument returned by GetTwoPhaseProp is either a CapeArrayDouble that
        ''' contains one or more numerical values, e.g. kvalues, or a CapeInterface that may be used to
        ''' retrieve 2-phase Physical Properties described by a more complex data structure, e.g.
        ''' distributed Physical Properties.
        ''' Although the result of some calls to GetTwoPhaseProp may be a single numerical value, the
        ''' return type for numerical values is CapeArrayDouble and in such a case the method must
        ''' return an array even if it contains only a single element.
        ''' A Phase is ‘present’ in a Material if its identifier is returned by the GetPresentPhases
        ''' method. An exception is raised by the GetTwoPhaseProp method if any of the Phases
        ''' specified is not present. Even if all Phases are present, this does not necessarily mean that
        ''' any Physical Properties are available.
        ''' The Physical Property values returned by GetTwoPhaseProp depend on two Phases, for
        ''' example surface tension or K-values. These values may be set by the SetTwoPhaseProp
        ''' method that may be called directly, or by other methods such as the CalcTwoPhaseProp
        ''' method of the ICapeThermoPropertyRoutine interface, or the CalcEquilibrium method of the
        ''' ICapeThermoEquilibriumRoutine interface. Note: Physical Properties that depend on a
        ''' single Phase are returned by the GetSinglePhaseProp method.
        ''' It is expected that this method will normally be able to provide Physical Property values on
        ''' any basis, i.e. it should be able to convert values from the basis on which they are stored to
        ''' the basis requested. This operation will not always be possible. For example, if the
        ''' molecular weight is not known for one or more Compounds, it is not possible to convert
        ''' between a mass basis and a molar basis.
        ''' If a composition derivative is requested this means that the derivatives are returned for both
        ''' Phases in the order in which the Phase labels are specified. The number of values returned
        ''' for a composition derivative will depend on the dimensionality of the property. For example,
        ''' if there are N Compounds then the results vector for the surface tension derivative will
        ''' contain N composition derivative values for the first Phase, followed by N composition
        ''' derivative values for the second Phase. For K-value derivative there will be N2 derivative
        ''' values for the first phase followed by N2 values for the second phase in the order defined in
        ''' 7.6.2.</remarks>
        Public Sub GetTwoPhaseProp(ByVal [property] As String, ByVal phaseLabels As Object, ByVal basis As String, ByRef results As Object) Implements ICapeThermoMaterial.GetTwoPhaseProp

            Me.PropertyPackage.CurrentMaterialStream = Me
            Dim res As New ArrayList
            Dim f1 As Integer = -1
            Dim f2 As Integer = -1
            Dim comps As New ArrayList
            For Each c As Substancia In Me.Fases(0).Componentes.Values
                comps.Add(c.Nome)
            Next

            Select Case phaseLabels(0).ToLower
                Case "overall"
                    f1 = 0
                Case Else
                    For Each pi As PhaseInfo In Me.PropertyPackage.PhaseMappings.Values
                        If phaseLabels(0) = pi.PhaseLabel Then
                            f1 = pi.DWPhaseIndex
                            Exit For
                        End If
                    Next
            End Select

            Select Case phaseLabels(1).ToLower
                Case "overall"
                    f2 = 0
                Case Else
                    For Each pi As PhaseInfo In Me.PropertyPackage.PhaseMappings.Values
                        If phaseLabels(1) = pi.PhaseLabel Then
                            f2 = pi.DWPhaseIndex
                            Exit For
                        End If
                    Next
            End Select

            Select Case [property].ToLower
                Case "kvalue"
                    For Each c As String In comps
                        res.Add(Me.Fases(f1).Componentes(c).FracaoMolar.GetValueOrDefault / Me.Fases(f2).Componentes(c).FracaoMolar.GetValueOrDefault)
                    Next
                Case "logkvalue"
                    For Each c As String In comps
                        res.Add(Math.Log(Me.Fases(f1).Componentes(c).FracaoMolar.GetValueOrDefault / Me.Fases(f2).Componentes(c).FracaoMolar.GetValueOrDefault))
                    Next
                Case "surfacetension"
                    res.Add(Me.Fases(f2).TPMProperties.surfaceTension.GetValueOrDefault)
                Case Else
                    Throw New Exception
            End Select
            Dim arr2(res.Count - 1) As Double
            Array.Copy(res.ToArray, arr2, res.Count)
            results = arr2
        End Sub

        ''' <summary>
        ''' Sets non-constant property values for the overall mixture.
        ''' </summary>
        ''' <param name="property">The identifier of the property for which values are set.</param>
        ''' <param name="basis">Basis of the results. Valid settings are: “Mass” for Physical Properties 
        ''' per unit mass or “Mole” for molar properties. Use UNDEFINED as a place holder for a Physical 
        ''' Property for which basis does not apply.</param>
        ''' <param name="values">Values to set for the property.s</param>
        ''' <remarks>The property values set by SetOverallProp refer to the overall mixture. These values are
        ''' retrieved by calling the GetOverallProp method. Overall mixture properties are not
        ''' calculated by components that implement the ICapeThermoMaterial interface. The property
        ''' values are only used as input specifications for the CalcEquilibrium method of a component
        ''' that implements the ICapeThermoEquilibriumRoutine interface.
        ''' Although some properties set by calls to SetOverallProp will have a single value, the type of
        ''' argument values is CapeArrayDouble and the method must always be called with values as
        ''' an array even if it contains only a single element.</remarks>
        Public Sub SetOverallProp(ByVal [property] As String, ByVal basis As String, ByVal values As Object) Implements ICapeThermoMaterial.SetOverallProp
            Me.SetSinglePhaseProp([property], "Overall", basis, values)
        End Sub

        ''' <summary>
        ''' Allows the PME or the Property Package to specify the list of Phases that are currently present.
        ''' </summary>
        ''' <param name="phaseLabels">The list of Phase labels for the Phases present.
        ''' The Phase labels in the Material Object must be a
        ''' subset of the labels returned by the GetPhaseList
        ''' method of the ICapeThermoPhases interface.</param>
        ''' <param name="phaseStatus"></param>
        ''' <remarks>SetPresentPhases is intended to be used in the following ways:
        ''' * To restrict an Equilibrium Calculation (using the CalcEquilibrium method of a
        ''' component that implements the ICapeThermoEquilibriumRoutine interface) to a subset
        ''' of the Phases supported by the Property Package component;
        ''' * When the component that implements the ICapeThermoEquilibriumRoutine interface
        ''' needs to specify which Phases are present in a Material Object after an Equilibrium
        ''' Calculation has been performed.
        ''' * In the context of dynamic simulations to specify the state of a Material Object that is an
        ''' output of a unit operation. This is the equivalent of calculating equilibrium in steadystate
        ''' simulations.
        ''' If a Phase in the list is already present, its Physical Properties are unchanged by the action of
        ''' this method. Any Phases not in the list when SetPresentPhases is called are removed from
        ''' the Material Object. This means that any Physical Property values that may have been stored
        ''' on the removed Phases are no longer available (i.e. a call to GetSinglePhaseProp or
        ''' GetTwoPhaseProp including this Phase will return an exception). A call to the
        ''' GetPresentPhases method of the Material Object will return the same list as specified by
        ''' SetPresentPhases.
        ''' The phaseStatus argument must contain as many entries as there are Phase labels. The valid
        ''' settings are listed in the following table:
        ''' 
        ''' Identifier                 Meaning
        ''' Cape_UnknownPhaseStatus    This is the normal setting when a Phase is specified as being available for an Equilibrium Calculation.
        ''' Cape_AtEquilibrium         The Phase has been set as present as a result of an Equilibrium Calculation.
        ''' Cape_Estimates             Estimates of the equilibrium state have been set in the Material Object.
        ''' 
        ''' All the Phases with a status of Cape_AtEquilibrium must have properties that correspond to
        ''' an equilibrium state, i.e. equal temperature, pressure and fugacities of each Compound (this
        ''' does not imply that the fugacities are set as a result of the Equilibrium Calculation). The
        ''' Cape_AtEquilibrium status should be set by the CalcEquilibrium method of a component
        ''' that implements the ICapeThermoEquilibriumRoutine interface following a successful
        ''' Equilibrium Calculation. If the temperature, pressure or composition of an equilibrium Phase
        ''' is changed, the Material Object implementation is responsible for resetting the status of the
        ''' Phase to Cape_UnknownPhaseStatus. Other property values stored for that Phase should not
        ''' be affected.
        ''' Phases with an Estimates status must have values of temperature, pressure, composition and
        ''' phase fraction set in the Material Object. These values are available for use by an
        ''' Equilibrium Calculator component to initialise an Equilibrium Calculation. The stored
        ''' values are available but there is no guarantee that they will be used.</remarks>
        Public Sub SetPresentPhases(ByVal phaseLabels As Object, ByVal phaseStatus As Object) Implements ICapeThermoMaterial.SetPresentPhases
            'do nothing, done automatically by CalcEquilibrium
        End Sub

        ''' <summary>
        ''' Sets single-phase non-constant property values for a mixture.
        ''' </summary>
        ''' <param name="property">The identifier of the property for which values are set. This must be 
        ''' one of the single-phase properties or derivatives. The standard identifiers are listed in 
        ''' sections 7.5.5 and 7.6.</param>
        ''' <param name="phaseLabel">Phase label of the Phase for which the property is set. The phase 
        ''' label must be one of the strings returned by the GetPhaseList method of the ICapeThermoPhases 
        ''' interface.</param>
        ''' <param name="basis">Basis of the results. Valid settings are: “Mass” for Physical Properties 
        ''' per unit mass or “Mole” for molar properties. Use UNDEFINED as a place holder for a Physical 
        ''' Property for which basis does not apply. See section 7.5.5 for details.</param>
        ''' <param name="values">Values to set for the property (CapeArrayDouble) or CapeInterface (see notes).</param>
        ''' <remarks>The values argument of SetSinglePhaseProp is either a CapeArrayDouble that contains one
        ''' or more numerical values to be set for a property, e.g. temperature, or a CapeInterface that
        ''' may be used to set single-phase properties described by a more complex data structure, e.g.
        ''' distributed properties.
        ''' It is required that a component that implements the ICapeThermoMaterial interface will
        ''' always support the following properties: temperature, pressure, fraction, phaseFraction,
        ''' flow, totalFlow.
        ''' Although some properties set by calls to SetSinglePhaseProp will have a single numerical
        ''' value, the type of the values argument for numerical values is CapeArrayDouble and in such
        ''' a case the method must be called with values containing an array even if it contains only a
        ''' single element.
        ''' The property values set by SetSinglePhaseProp refer to a single Phase. Properties that depend
        ''' on more than one Phase, for example surface tension or K-values, are set by the
        ''' SetTwoPhaseProp method of the ICapeThermoMaterial Interface.
        ''' To set a property using SetSinglePhaseProp, a phaseLabel identifier should be passed that is
        ''' supported by the Property Package or Material Object, i.e. one that appears in the list
        ''' returned by the GetPhaseList method of the ICapeThermoPhases interface. Setting such a
        ''' property should cause the phase to be present on the Material Object, as if it were specified
        ''' in a call to SetPresentPhases with status Cape_UnknownPhaseStatus. The SetPresentPhases
        ''' method of this interface does not need to be called before calling SetSinglePhaseProp.</remarks>
        Public Sub SetSinglePhaseProp(ByVal [property] As String, ByVal phaseLabel As String, ByVal basis As String, ByVal values As Object) Implements ICapeThermoMaterial.SetSinglePhaseProp

            Dim comps As New ArrayList
            For Each c As Substancia In Me.Fases(0).Componentes.Values
                comps.Add(c.Nome)
            Next
            Dim f As Integer = -1
            Dim phs As DTL.SimulationObjects.PropertyPackages.Fase
            Select Case phaseLabel.ToLower
                Case "overall"
                    f = 0
                    phs = PropertyPackages.Fase.Mixture
                Case Else
                    For Each pi As PhaseInfo In Me.PropertyPackage.PhaseMappings.Values
                        If phaseLabel = pi.PhaseLabel Then
                            f = pi.DWPhaseIndex
                            phs = pi.DWPhaseID
                            Exit For
                        End If
                    Next
            End Select

            If f = -1 Then
                Dim ex As New ArgumentException("Invalid Phase ID")
                Dim hcode As Integer = 0
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoMaterial", ex.Source, ex.StackTrace, "GetOverallTPFraction", hcode)
            End If

            Select Case [property].ToLower
                Case "compressibilityfactor"
                    Me.Fases(f).SPMProperties.compressibilityFactor = values(0)
                Case "heatcapacity", "heatcapacitycp"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Me.Fases(f).SPMProperties.heatCapacityCp = values(0) / Me.PropertyPackage.AUX_MMM(phs)
                        Case "Mass", "mass"
                            Me.Fases(f).SPMProperties.heatCapacityCp = values(0) / 1000
                    End Select
                Case "heatcapacitycv"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Me.Fases(f).SPMProperties.heatCapacityCv = values(0) / Me.PropertyPackage.AUX_MMM(phs)
                        Case "Mass", "mass"
                            Me.Fases(f).SPMProperties.heatCapacityCv = values(0) / 1000
                    End Select
                Case "excessenthalpy"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Me.Fases(f).SPMProperties.excessEnthalpy = values(0) / Me.PropertyPackage.AUX_MMM(phs)
                        Case "Mass", "mass"
                            Me.Fases(f).SPMProperties.excessEnthalpy = values(0) / 1000
                    End Select
                Case "excessentropy"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Me.Fases(f).SPMProperties.excessEntropy = values(0) / Me.PropertyPackage.AUX_MMM(phs)
                        Case "Mass", "mass"
                            Me.Fases(f).SPMProperties.excessEntropy = values(0) / 1000
                    End Select
                Case "viscosity"
                    Me.Fases(f).SPMProperties.viscosity = values(0)
                    Me.Fases(f).SPMProperties.kinematic_viscosity = values(0) / Me.Fases(f).SPMProperties.density.GetValueOrDefault
                Case "thermalconductivity"
                    Me.Fases(f).SPMProperties.thermalConductivity = values(0)
                Case "fugacity"
                    Dim i As Integer = 0
                    For Each c As String In comps
                        Me.Fases(f).Componentes(c).FugacityCoeff = values(comps.IndexOf(c)) / (Me.Fases(0).SPMProperties.pressure.GetValueOrDefault * Me.Fases(f).Componentes(c).FracaoMolar.GetValueOrDefault)
                        i += 1
                    Next
                Case "fugacitycoefficient"
                    Dim i As Integer = 0
                    For Each c As String In comps
                        Me.Fases(f).Componentes(c).FugacityCoeff = values(comps.IndexOf(c))
                        i += 1
                    Next
                Case "activitycoefficient"
                    Dim i As Integer = 0
                    For Each c As String In comps
                        Me.Fases(f).Componentes(c).ActivityCoeff = values(comps.IndexOf(c))
                        i += 1
                    Next
                Case "logfugacitycoefficient"
                    Dim i As Integer = 0
                    For Each c As String In comps
                        Me.Fases(f).Componentes(c).FugacityCoeff = Math.Exp(values(comps.IndexOf(c)))
                        i += 1
                    Next
                Case "density"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Me.Fases(f).SPMProperties.density = values(0) / 1000 * Me.PropertyPackage.AUX_MMM(phs)
                        Case "Mass", "mass"
                            Me.Fases(f).SPMProperties.density = values(0)
                    End Select
                Case "volume"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Me.Fases(f).SPMProperties.density = 1 / values(0) / 1000 * Me.PropertyPackage.AUX_MMM(phs)
                        Case "Mass", "mass"
                            Me.Fases(f).SPMProperties.density = 1 / values(0)
                    End Select
                Case "enthalpy", "enthalpynf"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Me.Fases(f).SPMProperties.molar_enthalpy = values(0)
                            Dim val = Me.PropertyPackage.AUX_MMM(phs)
                            If val <> 0.0# Then Me.Fases(f).SPMProperties.enthalpy = values(0) / val
                        Case "Mass", "mass"
                            Me.Fases(f).SPMProperties.enthalpy = values(0) / 1000
                            Me.Fases(f).SPMProperties.molar_enthalpy = values(0) * Me.PropertyPackage.AUX_MMM(phs)
                    End Select
                Case "entropy", "entropynf"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Me.Fases(f).SPMProperties.molar_entropy = values(0)
                            Dim val = Me.PropertyPackage.AUX_MMM(phs)
                            If val <> 0.0# Then Me.Fases(f).SPMProperties.entropy = values(0) / val
                        Case "Mass", "mass"
                            Me.Fases(f).SPMProperties.entropy = values(0) / 1000
                            Me.Fases(f).SPMProperties.molar_entropy = values(0) * Me.PropertyPackage.AUX_MMM(phs)
                    End Select
                Case "enthalpyf"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Dim val = Me.PropertyPackage.AUX_MMM(phs)
                            Me.Fases(f).SPMProperties.molar_enthalpyF = values(0)
                            If val <> 0.0# Then Me.Fases(f).SPMProperties.enthalpyF = values(0) / val
                        Case "Mass", "mass"
                            Me.Fases(f).SPMProperties.enthalpyF = values(0) / 1000
                            Me.Fases(f).SPMProperties.molar_enthalpyF = values(0) * Me.PropertyPackage.AUX_MMM(phs)
                    End Select
                Case "entropyf"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Dim val = Me.PropertyPackage.AUX_MMM(phs)
                            Me.Fases(f).SPMProperties.molar_entropyF = values(0)
                            If val <> 0.0# Then Me.Fases(f).SPMProperties.entropyF = values(0) / val
                        Case "Mass", "mass"
                            Me.Fases(f).SPMProperties.entropyF = values(0) / 1000
                            Me.Fases(f).SPMProperties.molar_entropyF = values(0) * Me.PropertyPackage.AUX_MMM(phs)
                    End Select
                Case "moles"
                    Me.Fases(f).SPMProperties.molarflow = values(0)
                Case "Mass"
                    Me.Fases(f).SPMProperties.massflow = values(0)
                Case "molecularweight"
                    Me.Fases(f).SPMProperties.molecularWeight = values(0)
                Case "temperature"
                    Me.Fases(0).SPMProperties.temperature = values(0)
                Case "pressure"
                    Me.Fases(0).SPMProperties.pressure = values(0)
                Case "flow"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Dim i As Integer = 0
                            For Each c As String In comps
                                Me.Fases(f).Componentes(c).MolarFlow = values(comps.IndexOf(c))
                                i += 1
                            Next
                        Case "Mass", "mass"
                            Dim i As Integer = 0
                            For Each c As String In comps
                                Me.Fases(f).Componentes(c).MassFlow = values(comps.IndexOf(c))
                                i += 1
                            Next
                    End Select
                Case "fraction"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Dim i As Integer = 0
                            For Each c As String In comps
                                Me.Fases(f).Componentes(c).FracaoMolar = values(comps.IndexOf(c))
                                i += 1
                            Next
                        Case "Mass", "mass"
                            Dim i As Integer = 0
                            For Each c As String In comps
                                Me.Fases(f).Componentes(c).FracaoMassica = values(comps.IndexOf(c))
                                i += 1
                            Next
                    End Select
                Case "phasefraction"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Me.Fases(f).SPMProperties.molarfraction = values(0)
                        Case "Mass", "mass"
                            Me.Fases(f).SPMProperties.massfraction = values(0)
                    End Select
                Case "totalflow"
                    Select Case basis
                        Case "Molar", "molar", "mole", "Mole"
                            Me.Fases(f).SPMProperties.molarflow = values(0)
                            'Me.Fases(f).SPMProperties.massflow = values(0) / 1000 * Me.PropertyPackage.AUX_MMM(phs)
                        Case "Mass", "mass"
                            Me.Fases(f).SPMProperties.massflow = values(0)
                            'Me.Fases(f).SPMProperties.molarflow = values(0)
                    End Select
                Case Else
                    Dim ex = New Exception
                    Dim hcode As Integer = 0
                    ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoMaterial", ex.Source, ex.StackTrace, "GetSinglePhaseProp", hcode)
            End Select
        End Sub

        ''' <summary>
        ''' Sets two-phase non-constant property values for a mixture.
        ''' </summary>
        ''' <param name="property">The property for which values are set in the Material Object. This 
        ''' must be one of the two-phase properties or derivatives included in sections 7.5.6 and 7.6.</param>
        ''' <param name="phaseLabels">Phase labels of the Phases for which the property is set. The Phase
        ''' labels must be two of the identifiers returned by the GetPhaseList method of the ICapeThermoPhases
        ''' interface.</param>
        ''' <param name="basis">Basis of the results. Valid settings are: “Mass” for Physical Properties per unit
        ''' mass or “Mole” for molar properties. Use UNDEFINED as a place holder for a Physical Property for which 
        ''' basis does not apply. See section 7.5.5 for details.</param>
        ''' <param name="values">Value(s) to set for the property (CapeArrayDouble) or CapeInterface (see notes).</param>
        ''' <remarks>The values argument of SetTwoPhaseProp is either a CapeArrayDouble that contains one or
        ''' more numerical values to be set for a property, e.g. kvalues, or a CapeInterface that may be
        ''' used to set two-phase properties described by a more complex data structure, e.g. distributed
        ''' properties.
        ''' Although some properties set by calls to SetTwoPhaseProp will have a single numerical
        ''' value, the type of the values argument for numerical values is CapeArrayDouble and in such
        ''' a case the method must be called with the values argument containing an array even if it
        ''' contains only a single element.
        ''' The Physical Property values set by SetTwoPhaseProp depend on two Phases, for example
        ''' surface tension or K-values. Properties that depend on a single Phase are set by the
        ''' SetSinglePhaseProp method.
        ''' If a Physical Property with composition derivative is specified, the derivative values will be
        ''' set for both Phases in the order in which the Phase labels are specified. The number of
        ''' values returned for a composition derivative will depend on the property. For example, if
        ''' there are N Compounds then the values vector for the surface tension derivative will contain
        ''' N composition derivative values for the first Phase, followed by N composition derivative
        ''' values for the second Phase. For K-values there will be N2 derivative values for the first
        ''' phase followed by N2 values for the second phase in the order defined in 7.6.2.
        ''' To set a property using SetTwoPhaseProp, phaseLabels identifiers should be passed that are
        ''' supported by the Property Package or Material Object, i.e. one that appears in the list returned by the
        ''' GetPhaseList method of the ICapeThermoPhases interface. Setting such a property should cause the
        ''' phases to be present on the Material Object, as if it were present in a call to SetPresentPhases with
        ''' status Cape_UnknownPhaseStatus. The SetPresentPhases method of this interface does not need to
        ''' be called before calling SetTwoPhaseProp.</remarks>
        Public Sub SetTwoPhaseProp(ByVal [property] As String, ByVal phaseLabels As Object, ByVal basis As String, ByVal values As Object) Implements ICapeThermoMaterial.SetTwoPhaseProp
            Dim comps As New ArrayList
            For Each c As Substancia In Me.Fases(0).Componentes.Values
                comps.Add(c.Nome)
            Next
            Select Case [property].ToLower
                Case "kvalue"
                    Dim i As Integer = 0
                    For Each c As String In comps
                        Me.Fases(0).Componentes(c).Kvalue = values(i)
                        i += 1
                    Next
                Case "logkvalue"
                    Dim i As Integer = 0
                    For Each c As String In comps
                        Me.Fases(0).Componentes(c).lnKvalue = values(i)
                        i += 1
                    Next
                Case "surfacetension"
                    Me.Fases(0).TPMProperties.surfaceTension = values(0)
                Case Else
                    Dim ex = New Exception
                    Dim hcode As Integer = 0
                    ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoMaterial", ex.Source, ex.StackTrace, "GetSinglePhaseProp", hcode)
            End Select
        End Sub

        ''' <summary>
        ''' Returns the number of Phases.
        ''' </summary>
        ''' <returns>The number of Phases supported.</returns>
        ''' <remarks>The number of Phases returned by this method must be equal to the number of Phase labels
        ''' that are returned by the GetPhaseList method of this interface. It must be zero, or a positive
        ''' number.</remarks>
        Public Function GetNumPhases() As Integer Implements ICapeThermoPhases.GetNumPhases
            Me.PropertyPackage.CurrentMaterialStream = Me
            Return Me.PropertyPackage.GetNumPhases()
        End Function

        ''' <summary>
        ''' Returns information on an attribute associated with a Phase for the purpose of understanding 
        ''' what lies behind a Phase label.
        ''' </summary>
        ''' <param name="phaseLabel">A (single) Phase label. This must be one of the values returned by GetPhaseList method.</param>
        ''' <param name="phaseAttribute">One of the Phase attribute identifiers from the table below.</param>
        ''' <returns>The value corresponding to the Phase attribute identifier – see table below.</returns>
        ''' <remarks>GetPhaseInfo is intended to allow a PME, or other client, to identify a Phase with an arbitrary
        ''' label. A PME, or other client, will need to do this to map stream data into a Material
        ''' Object, or when importing a Property Package. If the client cannot identify the Phase, it can
        ''' ask the user to provide a mapping based on the values of these properties.
        ''' The list of supported Phase attributes is defined in the following table:
        ''' 
        ''' Phase attribute identifier            Supported values
        ''' 
        ''' StateOfAggregation                    One of the following strings:
        '''                                       Vapor
        '''                                       Liquid
        '''                                       Solid
        '''                                       Unknown
        ''' 
        ''' KeyCompoundId                         The identifier of the Compound (compId as returned by GetCompoundList) 
        '''                                       that is expected to be present in highest concentration in the Phase. 
        '''                                       May be undefined in which case UNDEFINED should be returned.
        ''' 
        ''' ExcludedCompoundId                    The identifier of the Compound (compId as returned by
        '''                                       GetCompoundList) that is expected to be present in low or zero
        '''                                       concentration in the Phase. May not be defined in which case
        '''                                       UNDEFINED should be returned.
        ''' 
        ''' DensityDescription                    A description that indicates the density range expected for the Phase.
        '''                                       One of the following strings or UNDEFINED:
        '''                                       Heavy
        '''                                       Light
        ''' 
        ''' UserDescription                       A description that helps the user or PME to identify the Phase.
        '''                                       It can be any string or UNDEFINED.
        ''' 
        ''' TypeOfSolid                           A description that provides more information about a solid Phase. For
        '''                                       Phases with a “Solid” state of aggregation it may be one of the
        '''                                       following standard strings or UNDEFINED:
        '''                                       PureSolid
        '''                                       SolidSolution
        '''                                       HydrateI
        '''                                       HydrateII
        '''                                       HydrateH
        '''                                       Other values may be returned for solid Phases but these may not be
        '''                                       understood by most clients.
        '''                                       For Phases with any other state of aggregation it must be
        '''                                       UNDEFINED.</remarks>
        Public Function GetPhaseInfo(ByVal phaseLabel As String, ByVal phaseAttribute As String) As Object Implements ICapeThermoPhases.GetPhaseInfo
            Me.PropertyPackage.CurrentMaterialStream = Me
            Try
                Return Me.PropertyPackage.GetPhaseInfo(phaseLabel, phaseAttribute)
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoPhases", ex.Source, ex.StackTrace, "GetPhaseInfo", hcode)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Returns Phase labels and other important descriptive information for all the Phases supported.
        ''' </summary>
        ''' <param name="phaseLabels">he list of Phase labels for the Phases supported. A Phase label can 
        ''' be any string but each Phase must have a unique label. If, for some reason, no Phases are 
        ''' supported an UNDEFINED value should be returned for the phaseLabels. The number of Phase labels 
        ''' must also be equal to the number of Phases returned by the GetNumPhases method.</param>
        ''' <param name="stateOfAggregation">The physical State of Aggregation associated with each of the 
        ''' Phases. This must be one of the following strings: ”Vapor”, “Liquid”, “Solid” or “Unknown”. Each 
        ''' Phase must have a single State of Aggregation. The value must not be left undefined, but may be 
        ''' set to “Unknown”.</param>
        ''' <param name="keyCompoundId">The key Compound for the Phase. This must be the Compound identifier 
        ''' (as returned by GetCompoundList), or it may be undefined in which case a UNDEFINED value is returned. 
        ''' The key Compound is an indication of the Compound that is expected to be present in high concentration 
        ''' in the Phase, e.g. water for an aqueous liquid phase. Each Phase can have a single key Compound.</param>
        ''' <remarks>The Phase label allows the phase to be uniquely identified in methods of the ICapeThermo-
        ''' Phases interface and other CAPE-OPEN interfaces. The State of Aggregation and key
        ''' Compound provide a way for the PME, or other client, to interpret the meaning of a Phase
        ''' label in terms of the physical characteristics of the Phase.
        ''' All arrays returned by this method must be of the same length, i.e. equal to the number of
        ''' Phase labels.
        ''' To get further information about a Phase, use the GetPhaseInfo method.</remarks>
        Public Sub GetPhaseList(ByRef phaseLabels As Object, ByRef stateOfAggregation As Object, ByRef keyCompoundId As Object) Implements ICapeThermoPhases.GetPhaseList
            Me.PropertyPackage.CurrentMaterialStream = Me
            Try
                Me.PropertyPackage.GetPhaseList1(phaseLabels, stateOfAggregation, keyCompoundId)
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoPhases", ex.Source, ex.StackTrace, "GetPhaseList", hcode)
            End Try
        End Sub

        ''' <summary>
        ''' Retrieves the value of a Universal Constant.
        ''' </summary>
        ''' <param name="constantId">Identifier of Universal Constant. The list of constants supported should be 
        ''' obtained by using the GetUniversalConstantList method.</param>
        ''' <returns>Value of Universal Constant. This could be a numeric or a string value. For numeric values 
        ''' the units of measurement are specified in section 7.5.1.</returns>
        ''' <remarks>Universal Constants (often called fundamental constants) are quantities like the gas constant,
        ''' or the Avogadro constant.</remarks>
        Public Function GetUniversalConstant(ByVal constantId As String) As Object Implements ICapeThermoUniversalConstant.GetUniversalConstant
            Me.PropertyPackage.CurrentMaterialStream = Me
            Return Me.PropertyPackage.GetUniversalConstant1(constantId)
        End Function

        ''' <summary>
        ''' Returns the identifiers of the supported Universal Constants.
        ''' </summary>
        ''' <returns>List of identifiers of Universal Constants. The list of standard identifiers is given in section 7.5.1.</returns>
        ''' <remarks>A component may return Universal Constant identifiers that do not belong to the list defined
        ''' in section 7.5.1. However, these proprietary identifiers may not be understood by most of the
        ''' clients of this component.</remarks>
        Public Function GetUniversalConstantList() As Object Implements ICapeThermoUniversalConstant.GetUniversalConstantList
            Me.PropertyPackage.CurrentMaterialStream = Me
            Return Me.PropertyPackage.GetUniversalConstantList()
        End Function

        ''' <summary>
        ''' This method is used to calculate the natural logarithm of the fugacity coefficients (and
        ''' optionally their derivatives) in a single Phase mixture. The values of temperature, pressure
        ''' and composition are specified in the argument list and the results are also returned through
        ''' the argument list.
        ''' </summary>
        ''' <param name="phaseLabel">Phase label of the Phase for which the properties are to be calculated. 
        ''' The Phase label must be one of the strings returned by the GetPhaseList method on the ICapeThermoPhases interface.</param>
        ''' <param name="temperature">The temperature (K) for the calculation.</param>
        ''' <param name="pressure">The pressure (Pa) for the calculation.</param>
        ''' <param name="moleNumbers">Number of moles of each Compound in the mixture.</param>
        ''' <param name="fFlags">Code indicating whether natural logarithm of the fugacity coefficients and/or derivatives 
        ''' should be calculated (see notes).</param>
        ''' <param name="lnPhi">Natural logarithm of the fugacity coefficients (if requested).</param>
        ''' <param name="lnPhiDT">Derivatives of natural logarithm of the fugacity coefficients w.r.t. temperature (if requested).</param>
        ''' <param name="lnPhiDP">Derivatives of natural logarithm of the fugacity coefficients w.r.t. pressure (if requested).</param>
        ''' <param name="lnPhiDn">Derivatives of natural logarithm of the fugacity coefficients w.r.t. mole numbers (if requested).</param>
        ''' <remarks>This method is provided to allow the natural logarithm of the fugacity coefficient, which is
        ''' the most commonly used thermodynamic property, to be calculated and returned in a highly
        ''' efficient manner.
        ''' The temperature, pressure and composition (mole numbers) for the calculation are specified
        ''' by the arguments and are not obtained from the Material Object by a separate request. Likewise,
        ''' any quantities calculated are returned through the arguments and are not stored in the
        ''' Material Object. The state of the Material Object is not affected by calling this method. It
        ''' should be noted however, that prior to calling CalcAndGetLnPhi a valid Material Object
        ''' must have been defined by calling the SetMaterial method on the
        ''' ICapeThermoMaterialContext interface of the component that implements the
        ''' ICapeThermoPropertyRoutine interface. The compounds in the Material Object must have
        ''' been identified and the number of values supplied in the moleNumbers argument must be
        ''' equal to the number of Compounds in the Material Object.
        ''' The fugacity coefficient information is returned as the natural logarithm of the fugacity
        ''' coefficient. This is because thermodynamic models naturally provide the natural logarithm
        ''' of this quantity and also a wider range of values may be safely returned.
        ''' The quantities actually calculated and returned by this method are controlled by an integer
        ''' code fFlags. The code is formed by summing contributions for the property and each
        ''' derivative required using the enumerated constants eCapeCalculationCode (defined in the
        ''' Thermo version 1.1 IDL) shown in the following table. For example, to calculate log
        ''' fugacity coefficients and their T-derivatives the fFlags argument would be set to
        ''' CAPE_LOG_FUGACITY_COEFFICIENTS + CAPE_T_DERIVATIVE.
        ''' 
        '''                                       code                            numerical value
        ''' no calculation                        CAPE_NO_CALCULATION             0
        ''' log fugacity coefficients             CAPE_LOG_FUGACITY_COEFFICIENTS  1
        ''' T-derivative                          CAPE_T_DERIVATIVE               2
        ''' P-derivative                          CAPE_P_DERIVATIVE               4
        ''' mole number derivatives               CAPE_MOLE_NUMBERS_DERIVATIVES   8
        ''' 
        ''' If CalcAndGetLnPhi is called with fFlags set to CAPE_NO_CALCULATION no property
        ''' values are returned.
        ''' A typical sequence of operations for this method when implemented by a Property Package
        ''' component would be:
        ''' - Check that the phaseLabel specified is valid.
        ''' - Check that the moleNumbers array contains the number of values expected
        ''' (should be consistent with the last call to the SetMaterial method).
        ''' - Calculate the requested properties/derivatives at the T/P/composition specified in
        ''' the argument list.
        ''' - Store values for the properties/derivatives in the corresponding arguments.
        ''' Note that this calculation can be carried out irrespective of whether the Phase actually exists
        ''' in the Material Object.</remarks>
        Public Sub CalcAndGetLnPhi(ByVal phaseLabel As String, ByVal temperature As Double, ByVal pressure As Double, ByVal moleNumbers As Object, ByVal fFlags As Integer, ByRef lnPhi As Object, ByRef lnPhiDT As Object, ByRef lnPhiDP As Object, ByRef lnPhiDn As Object) Implements ICapeThermoPropertyRoutine.CalcAndGetLnPhi
            Me.PropertyPackage.CurrentMaterialStream = Me
            Try
                Me.PropertyPackage.CalcAndGetLnPhi(phaseLabel, temperature, pressure, moleNumbers, fFlags, lnPhi, lnPhiDT, lnPhiDP, lnPhiDn)
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoPropertyRoutine", ex.Source, ex.StackTrace, "CalcAndGetLnPhi", hcode)
            End Try
        End Sub

        ''' <summary>
        ''' CalcSinglePhaseProp is used to calculate properties and property derivatives of a mixture in
        ''' a single Phase at the current values of temperature, pressure and composition set in the
        ''' Material Object. CalcSinglePhaseProp does not perform phase Equilibrium Calculations.
        ''' </summary>
        ''' <param name="props">The list of identifiers for the single-phase properties or derivatives to 
        ''' be calculated. See sections 7.5.5 and 7.6 for the standard identifiers.</param>
        ''' <param name="phaseLabel">Phase label of the Phase for which the properties are to be calculated. 
        ''' The Phase label must be one of the strings returned by the GetPhaseList method on the 
        ''' ICapeThermoPhases interface and the phase must be present in the Material Object.</param>
        ''' <remarks>CalcSinglePhaseProp calculates properties, such as enthalpy or viscosity that are defined for
        ''' a single Phase. Physical Properties that depend on more than one Phase, for example surface
        ''' tension or K-values, are handled by CalcTwoPhaseProp method.
        ''' Components that implement this method must get the input specification for the calculation
        ''' (temperature, pressure and composition) from the associated Material Object and set the
        ''' results in the Material Object.
        ''' Thermodynamic and Physical Properties Components, such as a Property Package or Property
        ''' Calculator, must implement the ICapeThermoMaterialContext interface so that an
        ''' ICapeThermoMaterial interface can be passed via the SetMaterial method.
        ''' The component that implements the ICapeThermoPropertyRoutine interface (e.g. a Property
        ''' Package or Property Calculator) must also implement the ICapeThermoPhases interface so
        ''' that it is possible to get a list of supported phases. The phaseLabel passed to this method
        ''' must be one of the phase labels returned by the GetPhaseList method of the
        ''' ICapeThermoPhases interface and it must also be present in the Material Object, ie. one of
        ''' the phase labels returned by the GetPresentPhases method of the ICapeThermoMaterial
        ''' interface. This latter condition will be satisfied if the phase is made present explicitly by
        ''' calling the SetPresentPhases method or if any phase properties have been set by calling the
        ''' SetSinglePhaseProp or SetTwoPhaseProp methods.
        ''' A typical sequence of operations for CalcSinglePhaseProp when implemented by a Property
        ''' Package component would be:
        ''' - Check that the phaseLabel specified is valid.
        ''' - Use the GetTPFraction method (of the Material Object specified in the last call to the
        ''' SetMaterial method) to get the temperature, pressure and composition of the
        ''' specified Phase.
        ''' - Calculate the properties.
        ''' - Store values for the properties of the Phase in the Material Object using the
        ''' SetSinglePhaseProp method of the ICapeThermoMaterial interface.
        ''' CalcSinglePhaseProp will request the input Property values it requires from the Material
        ''' Object through GetSinglePhaseProp calls. If a requested property is not available, the
        ''' exception raised will be ECapeThrmPropertyNotAvailable. If this error occurs then the
        ''' Property Package can return it to the client, or request a different property. Material Object
        ''' implementations must be able to supply property values using the client’s choice of basis by
        ''' implementing conversion from one basis to another.
        ''' Clients should not assume that Phase fractions and Compound fractions in a Material Object
        ''' are normalised. Fraction values may also lie outside the range 0 to 1. If fractions are not
        ''' normalised, or are outside the expected range, it is the responsibility of the Property Package
        ''' to decide how to deal with the situation.
        ''' It is recommended that properties are requested one at a time in order to simplify error
        ''' handling. However, it is recognised that there are cases where the potential efficiency gains
        ''' of requesting several properties simultaneously are more important. One such example
        ''' might be when a property and its derivatives are required.
        ''' If a client uses multiple properties in a call and one of them fails then the whole call should
        ''' be considered to have failed. This implies that no value should be written back to the Material
        ''' Object by the Property Package until it is known that the whole request can be satisfied.
        ''' It is likely that a PME might request values of properties for a Phase at conditions of temperature,
        ''' pressure and composition where the Phase does not exist (according to the
        ''' mathematical/physical models used to represent properties). The exception
        ''' ECapeThrmPropertyNotAvailable may be raised or an extrapolated value may be returned.
        ''' It is responsibility of the implementer to decide how to handle this circumstance.</remarks>
        Public Sub CalcSinglePhaseProp(ByVal props As Object, ByVal phaseLabel As String) Implements ICapeThermoPropertyRoutine.CalcSinglePhaseProp
            Me.PropertyPackage.CurrentMaterialStream = Me
            Try
                Me.PropertyPackage.CalcSinglePhaseProp(props, phaseLabel)
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoPropertyRoutine", ex.Source, ex.StackTrace, "CalcSinglePhaseProp", hcode)
            End Try
        End Sub

        ''' <summary>
        ''' CalcTwoPhaseProp is used to calculate mixture properties and property derivatives that depend on
        ''' two Phases at the current values of temperature, pressure and composition set in the Material Object.
        ''' It does not perform Equilibrium Calculations.
        ''' </summary>
        ''' <param name="props">The list of identifiers for properties to be calculated. This must be one or more 
        ''' of the supported two-phase properties and derivatives (as given by the GetTwoPhasePropList method). 
        ''' The standard identifiers for two-phase properties are given in section 7.5.6 and 7.6.</param>
        ''' <param name="phaseLabels">Phase labels of the phases for which the properties are to be calculated. 
        ''' The phase labels must be two of the strings returned by the GetPhaseList method on the ICapeThermoPhases 
        ''' interface and the phases must also be present in the Material Object.</param>
        ''' <remarks>CalcTwoPhaseProp calculates the values of properties such as surface tension or K-values.
        ''' Properties that pertain to a single Phase are handled by the CalcSinglePhaseProp method of
        ''' the ICapeThermoPropertyRoutine interface.Components that implement this method must
        ''' get the input specification for the calculation (temperature, pressure and composition) from
        ''' the associated Material Object and set the results in the Material Object.
        ''' Components such as a Property Package or Property Calculator must implement the
        ''' ICapeThermoMaterialContext interface so that an ICapeThermoMaterial interface can be
        ''' passed via the SetMaterial method.
        ''' The component that implements the ICapeThermoPropertyRoutine interface (e.g. a Property
        ''' Package or Property Calculator) must also implement the ICapeThermoPhases interface so
        ''' that it is possible to get a list of supported phases. The phaseLabels passed to this method
        ''' must be in the list of phase labels returned by the GetPhaseList method of the
        ''' ICapeThermoPhases interface and they must also be present in the Material Object, ie. in the
        ''' list of phase labels returned by the GetPresentPhases method of the ICapeThermoMaterial
        ''' interface. This latter condition will be satisfied if the phases are made present explicitly by
        ''' calling the SetPresentPhases method or if any phase properties have been set by calling the
        ''' SetSinglePhaseProp or SetTwoPhaseProp methods.</remarks>
        Public Sub CalcTwoPhaseProp(ByVal props As Object, ByVal phaseLabels As Object) Implements ICapeThermoPropertyRoutine.CalcTwoPhaseProp
            Me.PropertyPackage.CurrentMaterialStream = Me
            Try
                Me.PropertyPackage.CalcTwoPhaseProp(props, phaseLabels)
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoPropertyRoutine", ex.Source, ex.StackTrace, "CalcTwoPhaseProp", hcode)
            End Try
        End Sub

        ''' <summary>
        ''' Checks whether it is possible to calculate a property with the CalcSinglePhaseProp method for a given Phase.
        ''' </summary>
        ''' <param name="property">The identifier of the property to check. To be valid this must be one of the supported 
        ''' single-phase properties or derivatives (as given by the GetSinglePhasePropList method).</param>
        ''' <param name="phaseLabel">The Phase label for the calculation check. This must be one of the labels 
        ''' returned by the GetPhaseList method on the ICapeThermoPhases interface.</param>
        ''' <returns>Set to True if the combination of property and phaseLabel is supported or False if 
        ''' not supported.</returns>
        ''' <remarks>The result of the check should only depend on the capabilities and configuration
        ''' (Compounds and Phases supported) of the component that implements the
        ''' ICapeThermoPropertyRoutine interface (e.g. a Property Package). It should not depend on
        ''' whether a Material Object has been set nor on the state (temperature, pressure, composition
        ''' etc.), or configuration of a Material Object that might be set.
        ''' It is expected that the PME, or other client, will use this method to check whether the properties
        ''' it requires are supported by the Property Package when the package is imported. If any
        ''' essential properties are not available, the import process should be aborted.
        ''' If either the property or the phaseLabel arguments are not recognised by the component that
        ''' implements the ICapeThermoPropertyRoutine interface this method should return False.</remarks>
        Public Function CheckSinglePhasePropSpec(ByVal [property] As String, ByVal phaseLabel As String) As Boolean Implements ICapeThermoPropertyRoutine.CheckSinglePhasePropSpec
            Me.PropertyPackage.CurrentMaterialStream = Me
            Return Me.PropertyPackage.CheckSinglePhasePropSpec([property], phaseLabel)
        End Function

        ''' <summary>
        ''' Checks whether it is possible to calculate a property with the CalcTwoPhaseProp method for a given set of Phases.
        ''' </summary>
        ''' <param name="property">The identifier of the property to check. To be valid this must be one of the supported 
        ''' two-phase properties (including derivatives), as given by the GetTwoPhasePropList method.</param>
        ''' <param name="phaseLabels">Phase labels of the Phases for which the properties are to be calculated. The Phase 
        ''' labels must be two of the identifiers returned by the GetPhaseList method on the ICapeThermoPhases interface.</param>
        ''' <returns>Set to True if the combination of property and phaseLabels is supported, or False if not supported.</returns>
        ''' <remarks>The result of the check should only depend on the capabilities and configuration
        ''' (Compounds and Phases supported) of the component that implements the
        ''' ICapeThermoPropertyRoutine interface (e.g. a Property Package). It should not depend on
        ''' whether a Material Object has been set nor on the state (temperature, pressure, composition
        ''' etc.), or configuration of a Material Object that might be set.
        ''' It is expected that the PME, or other client, will use this method to check whether the
        ''' properties it requires are supported by the Property Package when the Property Package is
        ''' imported. If any essential properties are not available, the import process should be aborted.
        ''' If either the property argument or the values in the phaseLabels arguments are not
        ''' recognised by the component that implements the ICapeThermoPropertyRoutine interface
        ''' this method should return False.</remarks>
        Public Function CheckTwoPhasePropSpec(ByVal [property] As String, ByVal phaseLabels As Object) As Boolean Implements ICapeThermoPropertyRoutine.CheckTwoPhasePropSpec
            Me.PropertyPackage.CurrentMaterialStream = Me
            Return Me.PropertyPackage.CheckTwoPhasePropSpec([property], phaseLabels)
        End Function

        ''' <summary>
        ''' Returns the list of supported non-constant single-phase Physical Properties.
        ''' </summary>
        ''' <returns>List of all supported non-constant single-phase property identifiers. 
        ''' The standard single-phase property identifiers are listed in section 7.5.5.</returns>
        ''' <remarks>A non-constant property depends on the state of the Material Object.
        ''' Single-phase properties, e.g. enthalpy, only depend on the state of one phase.
        ''' GetSinglePhasePropList must return all the single-phase properties that can be calculated by
        ''' CalcSinglePhaseProp. If derivatives can be calculated these must also be returned. The list
        ''' of standard property identifiers in section 7.5.5 also contains properties such as temperature,
        ''' pressure, fraction, phaseFraction, flow and totalFlow that are not usually calculated by the
        ''' CalcSinglePhaseProp method and hence these property identifiers would not be returned by
        ''' GetSinglePhasePropList. These properties would normally be used in calls to the
        ''' Set/GetSinglePhaseProp methods of the ICapeThermoMaterial interface.
        ''' If no single-phase properties are supported this method should return UNDEFINED.
        ''' To get the list of supported two-phase properties, use GetTwoPhasePropList.
        ''' A component that implements this method may return non-constant single-phase property
        ''' identifiers which do not belong to the list defined in section 7.5.5. However, these
        ''' proprietary identifiers may not be understood by most of the clients of this component.</remarks>
        Public Function GetSinglePhasePropList() As Object Implements ICapeThermoPropertyRoutine.GetSinglePhasePropList
            Me.PropertyPackage.CurrentMaterialStream = Me
            Return Me.PropertyPackage.GetSinglePhasePropList()
        End Function

        ''' <summary>
        ''' Returns the list of supported non-constant two-phase properties.
        ''' </summary>
        ''' <returns>List of all supported non-constant two-phase property identifiers. The standard two-phase 
        ''' property identifiers are listed in section 7.5.6.</returns>
        ''' <remarks>A non-constant property depends on the state of the Material Object. Two-phase properties
        ''' are those that depend on more than one co-existing phase, e.g. K-values.
        ''' GetTwoPhasePropList must return all the properties that can be calculated by
        ''' CalcTwoPhaseProp. If derivatives can be calculated, these must also be returned.
        ''' If no two-phase properties are supported this method should return UNDEFINED.
        ''' To check whether a property can be evaluated for a particular set of phase labels use the
        ''' CheckTwoPhasePropSpec method.
        ''' A component that implements this method may return non-constant two-phase property
        ''' identifiers which do not belong to the list defined in section 7.5.6. However, these
        ''' proprietary identifiers may not be understood by most of the clients of this component.
        ''' To get the list of supported single-phase properties, use GetSinglePhasePropList.</remarks>
        Public Function GetTwoPhasePropList() As Object Implements ICapeThermoPropertyRoutine.GetTwoPhasePropList
            Me.PropertyPackage.CurrentMaterialStream = Me
            Return Me.PropertyPackage.GetTwoPhasePropList()
        End Function

        ''' <summary>
        ''' CalcEquilibrium is used to calculate the amounts and compositions of Phases at equilibrium.
        ''' CalcEquilibrium will calculate temperature and/or pressure if these are not among the two
        ''' specifications that are mandatory for each Equilibrium Calculation considered.
        ''' </summary>
        ''' <param name="specification1">First specification for the Equilibrium Calculation. The 
        ''' specification information is used to retrieve the value of the specification from the 
        ''' Material Object. See below for details.</param>
        ''' <param name="specification2">Second specification for the Equilibrium Calculation in 
        ''' the same format as specification1.</param>
        ''' <param name="solutionType">The identifier for the required solution type. The
        ''' standard identifiers are given in the following list:
        ''' Unspecified
        ''' Normal
        ''' Retrograde
        ''' The meaning of these terms is defined below in the notes. Other identifiers may be supported 
        ''' but their interpretation is not part of the CO standard.</param>
        ''' <remarks>The specification1 and specification2 arguments provide the information necessary to
        ''' retrieve the values of two specifications, for example the pressure and temperature, for the
        ''' Equilibrium Calculation. The CheckEquilibriumSpec method can be used to check for
        ''' supported specifications. Each specification variable contains a sequence of strings in the
        ''' order defined in the following table (hence, the specification arguments may have 3 or 4
        ''' items):
        ''' 
        ''' item                        meaning
        ''' 
        ''' property identifier         The property identifier can be any of the identifiers listed in section 7.5.5 but
        '''                             only certain property specifications will normally be supported by any
        '''                             Equilibrium Routine.
        ''' 
        ''' basis                       The basis for the property value. Valid settings for basis are given in section
        '''                             7.4. Use UNDEFINED as a placeholder for a property for which basis does
        '''                             not apply. For most Equilibrium Specifications, the result of the calculation
        '''                             is not dependent on the basis, but, for example, for phase fraction
        '''                             specifications the basis (Mole or Mass) does make a difference.
        ''' 
        ''' phase label                 The phase label denotes the Phase to which the specification applies. It must
        '''                             either be one of the labels returned by GetPresentPhases, or the special value
        '''                             “Overall”.
        ''' 
        ''' compound identifier         The compound identifier allows for specifications that depend on a particular
        '''                             Compound. This item of the specification array is optional and may be
        '''                             omitted. In case of a specification without compound identifier, the array
        '''                             element may be present and empty, or may be absent.
        '''                             The values corresponding to the specifications in the argument list and the overall
        '''                             composition of the mixture must be set in the associated Material Object before a call to
        '''                             CalcEquilibrium.
        ''' 
        ''' Components such as a Property Package or an Equilibrium Calculator must implement the
        ''' ICapeThermoMaterialContext interface, so that an ICapeThermoMaterial interface can be
        ''' passed via the SetMaterial method. It is the responsibility of the implementation of
        ''' CalcEquilibrium to validate the Material Object before attempting a calculation.
        ''' The Phases that will be considered in the Equilibrium Calculation are those that exist in the
        ''' Material Object, i.e. the list of phases specified in a SetPresentPhases call. This provides a
        ''' way for a client to specify whether, for example, a vapour-liquid, liquid-liquid, or vapourliquid-
        ''' liquid calculation is required. CalcEquilibrium must use the GetPresentPhases method
        ''' to retrieve the list of Phases and the associated Phase status flags. The Phase status flags may
        ''' be used by the client to provide information about the Phases, for example whether estimates
        ''' of the equilibrium state are provided. See the description of the GetPresentPhases and
        ''' SetPresentPhases methods of the ICapeThermoMaterial interface for details. When the
        ''' Equilibrium Calculation has been completed successfully, the SetPresentPhases method
        ''' must be used to specify which Phases are present at equilibrium and the Phase status flags
        ''' for the phases should be set to Cape_AtEquilibrium. This must include any Phases that are
        ''' present in zero amount such as the liquid Phase in a dew point calculation.
        ''' Some types of Phase equilibrium specifications may result in more than one solution. A
        ''' common example of this is the case of a dew point calculation. However, CalcEquilibrium
        ''' can provide only one solution through the Material Object. The solutionType argument
        ''' allows the “Normal” or “Retrograde” solution to be explicitly requested. When none of the
        ''' specifications includes a phase fraction, the solutionType argument should be set to
        ''' “Unspecified”.
        ''' 
        ''' CalcEquilibrium must set the amounts (phase fractions), compositions, temperature and
        ''' pressure for all Phases present at equilibrium, as well as the temperature and pressure for the
        ''' overall mixture if not set as part of the calculation specifications. It must not set any other
        ''' values – in particular it must not set any values for phases that are not present.
        ''' 
        ''' As an example, the following sequence of operations might be performed by
        ''' CalcEquilibrium in the case of an Equilibrium Calculation at fixed pressure and temperature:
        ''' 
        ''' - With the ICapeThermoMaterial interface of the supplied Material Object:
        ''' 
        ''' -- Use the GetPresentPhases method to find the list of Phases that the Equilibrium
        ''' Calculation should consider.
        ''' 
        ''' -- With the ICapeThermoCompounds interface of the Material Object use the
        ''' GetCompoundList method to find which Compounds are present.
        ''' 
        ''' -- Use the GetOverallProp method to get the temperature, pressure and composition
        ''' for the overall mixture.
        ''' 
        ''' - Perform the Equilibrium Calculation.
        ''' 
        ''' -- Use SetPresentPhases to specify the Phases present at equilibrium and set the
        ''' Phase status flags to Cape_AtEquilibrium.
        ''' 
        ''' -- Use SetSinglePhaseProp to set pressure, temperature, Phase amount (or Phase
        ''' fraction) and composition for all Phases present.</remarks>
        Public Sub CalcEquilibrium1(ByVal specification1 As Object, ByVal specification2 As Object, ByVal solutionType As String) Implements ICapeThermoEquilibriumRoutine.CalcEquilibrium
            Me.PropertyPackage.CurrentMaterialStream = Me
            Try
                Me.PropertyPackage.CalcEquilibrium1(specification1, specification2, solutionType)
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoEquilibriumRoutine", ex.Source, ex.StackTrace, "CalcEquilibrium", hcode)
            End Try
        End Sub

        ''' <summary>
        ''' Checks whether the Property Package can support a particular type of Equilibrium Calculation.
        ''' </summary>
        ''' <param name="specification1">First specification for the Equilibrium Calculation.</param>
        ''' <param name="specification2">Second specification for the Equilibrium Calculation.</param>
        ''' <param name="solutionType">The required solution type.</param>
        ''' <returns>Set to True if the combination of specifications and solutionType is supported 
        ''' for a particular combination of present phases or False if not supported.</returns>
        ''' <remarks>The meaning of the specification1, specification2 and solutionType arguments is the same as
        ''' for the CalcEquilibrium method. If solutionType, specification1 and specification2
        ''' arguments appear valid but the actual specifications are not supported or not recognised a
        ''' False value should be returned.
        ''' The result of the check should depend primarily on the capabilities and configuration
        ''' (compounds and phases supported) of the component that implements the ICapeThermo-
        ''' EquilibriumRoutine interface (egg. a Property package). A component that supports
        ''' calculation specifications for any combination of supported phases is capable of checking
        ''' the specification without any reference to a Material Object. However, it is possible that
        ''' there may be restrictions on the combinations of phases supported in an equilibrium
        ''' calculation. For example a component may support vapor-liquid and liquid-liquid
        ''' calculations but not vapor-liquid-liquid calculations. In general it is therefore a necessary
        ''' prerequisite that a Material Object has been set (using the SetMaterial method of the
        ''' ICapeThermoMaterialContext interface) and that the SetPresentPhases method of the
        ''' ICapeThermoMaterial interface has been called to specify the combination of phases for the
        ''' equilibrium calculation. The result of the check should not depend on the state (temperature,
        ''' pressure, composition etc.) of the Material Object.</remarks>
        Public Function CheckEquilibriumSpec(ByVal specification1 As Object, ByVal specification2 As Object, ByVal solutionType As String) As Boolean Implements ICapeThermoEquilibriumRoutine.CheckEquilibriumSpec
            Me.PropertyPackage.CurrentMaterialStream = Me
            Try
                Return Me.PropertyPackage.CheckEquilibriumSpec(specification1, specification2, solutionType)
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoEquilibriumRoutine", ex.Source, ex.StackTrace, "CalcEquilibrium", hcode)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Allows the client of a component that implements this interface to pass an ICapeThermoMaterial 
        ''' interface to the component, so that it can access the properties of a Material.
        ''' </summary>
        ''' <param name="material">The Material interface.</param>
        ''' <remarks>The SetMaterial method allows a Thermodynamic and Physical Properties component, such
        ''' as a Property Package, to be given the ICapeThermoMaterial interface of a Material Object.
        ''' This interface gives the component access to the description of the Material for which
        ''' Property Calculations or Equilibrium Calculations are required. The component can access
        ''' property values directly using this interface. A client can also use the ICapeThermoMaterial
        ''' interface to query a Material Object for its ICapeThermoCompounds and ICapeThermo-
        ''' Phases interfaces, which provide access to Compound and Phase information, respectively.
        ''' It is envisaged that the SetMaterial method will be used to check that the Material Interface
        ''' supplied is valid and useable. For example, a Property Package may check that there are
        ''' some Compounds in a Material Object and that those Compounds can be identified by the
        ''' Property Package. In addition a Property Package may perform any initialisation that
        ''' depends on the configuration of a Material Object. A Property Calculator component might
        ''' typically use this method to query the Material Object for any required information
        ''' concerning the Compounds.
        ''' Calling the UnsetMaterial method of the ICapeThermoMaterialContext interface has the
        ''' effect of removing the interface set by the SetMaterial method.
        ''' After a call to SetMaterial() has been received, the object implementing the ICapeThermo-
        ''' MaterialContext interface can assume that the number, name and order of compounds for
        ''' that Material Object will remain fixed until the next call to SetMaterial() or UnsetMaterial().</remarks>
        Public Sub SetMaterial(ByVal material As Object) Implements ICapeThermoMaterialContext.SetMaterial
            Try
                Me.PropertyPackage.SetMaterial(material)
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoMaterialContext", ex.Source, ex.StackTrace, "SetMaterial", hcode)
            End Try
        End Sub

        ''' <summary>
        ''' Removes any previously set Material interface.
        ''' </summary>
        ''' <remarks>The UnsetMaterial method removes any Material interface previously set by a call to the
        ''' SetMaterial method of the ICapeThermoMaterialContext interface. This means that any
        ''' methods of other interfaces that depend on having a valid Material Interface, for example
        ''' methods of the ICapeThermoPropertyRoutine or ICapeThermoEquilibriumRoutine
        ''' interfaces, should behave in the same way as if the SetMaterial method had never been
        ''' called.
        ''' If UnsetMaterial is called before a call to SetMaterial it has no effect and no exception
        ''' should be raised.</remarks>
        Public Sub UnsetMaterial() Implements ICapeThermoMaterialContext.UnsetMaterial
            Try
                Me.PropertyPackage.UnsetMaterial()
            Catch ex As Exception
                Dim hcode As Integer = 0
                Dim comEx As COMException = New COMException(ex.Message.ToString, ex)
                If Not IsNothing(comEx) Then hcode = comEx.ErrorCode
                ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoMaterialContext", ex.Source, ex.StackTrace, "UnsetMaterial", hcode)
            End Try
        End Sub

#End Region

#Region "    CAPE-OPEN Error Interfaces"

        Sub ThrowCAPEException(ByRef ex As Exception, ByVal name As String, ByVal description As String, ByVal interf As String, ByVal moreinfo As String, ByVal operation As String, ByVal scope As String, ByVal code As Integer)

            _name = name
            _code = code
            _description = description
            _interfacename = interf
            _moreinfo = moreinfo
            _operation = operation
            _scope = scope

            Throw ex

        End Sub

        Private _name, _description, _interfacename, _moreinfo, _operation, _scope As String, _code As Integer

        Public ReadOnly Property Name() As String Implements CapeOpen.ECapeRoot.name
            Get
                Return _name
            End Get
        End Property

        Public ReadOnly Property code() As Integer Implements CapeOpen.ECapeUser.code
            Get
                Return _code
            End Get
        End Property

        Public ReadOnly Property description() As String Implements CapeOpen.ECapeUser.description
            Get
                Return _description
            End Get
        End Property

        Public ReadOnly Property interfaceName() As String Implements CapeOpen.ECapeUser.interfaceName
            Get
                Return _interfacename
            End Get
        End Property

        Public ReadOnly Property moreInfo() As String Implements CapeOpen.ECapeUser.moreInfo
            Get
                Return _moreinfo
            End Get
        End Property

        Public ReadOnly Property operation() As String Implements CapeOpen.ECapeUser.operation
            Get
                Return _operation
            End Get
        End Property

        Public ReadOnly Property scope() As String Implements ECapeUser.scope
            Get
                Return _scope
            End Get
        End Property

#End Region

        Function Flowsheet() As Object
            Throw New NotImplementedException
        End Function

    End Class

End Namespace
