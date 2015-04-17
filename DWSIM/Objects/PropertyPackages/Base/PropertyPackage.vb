'    Property Package Base Class
'    Copyright 2008-2014 Daniel Wagner O. de Medeiros
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

Imports DTL.DTL.SimulationObjects.Streams
Imports DTL.DTL.SimulationObjects
Imports DTL.DTL.ClassesBasicasTermodinamica
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Runtime.Serialization
Imports System.IO
Imports System.Math
Imports CapeOpen = CAPEOPEN110
Imports System.Runtime.InteropServices.ComTypes
Imports iop = System.Runtime.InteropServices
Imports System.Xml.Serialization
Imports System.Runtime.Serialization.Formatters
Imports CAPEOPEN110
Imports System.Runtime.InteropServices
Imports System.Threading.Tasks
Imports DTL.DTL.MathEx
Imports Microsoft.VisualBasic.FileIO

Namespace DTL.SimulationObjects.PropertyPackages

#Region "    Global Enumerations"

    Public Enum Fase

        Liquid
        Liquid1
        Liquid2
        Liquid3
        Aqueous
        Vapor
        Solid
        Mixture

    End Enum

    Public Enum State
        Liquid = 0
        Vapor = 1
        Solid = 2
    End Enum

    Public Enum FlashSpec

        P
        T
        S
        H
        V
        U
        VAP

    End Enum

    Public Enum ThermoProperty
        ActivityCoefficient
        Fugacity
        FugacityCoefficient
    End Enum

    Public Enum PackageType
        EOS = 0
        ActivityCoefficient = 1
        ChaoSeader = 2
        VaporPressure = 3
        Miscelaneous = 4
        CorrespondingStates = 5
        CAPEOPEN = 6
    End Enum

    Public Enum FlashMethod
        GlobalSetting = 2
        DWSIMDefault = 0
        InsideOut = 1
        InsideOut3P = 3
        GibbsMin2P = 4
        GibbsMin3P = 5
        NestedLoops3P = 6
        NestedLoopsSLE = 7
        NestedLoopsImmiscible = 8
        SimpleLLE = 9
        NestedLoopsSLE_SS = 10
        NestedLoops3PV2 = 11
        NestedLoops3PV3 = 12
    End Enum

    Public Enum Parameter
        PHFlash_Internal_Loop_Tolerance = 0
        PSFlash_Internal_Loop_Tolerance = 1
        PHFlash_External_Loop_Tolerance = 2
        PSFlash_External_Loop_Tolerance = 3
        PHFlash_Maximum_Number_Of_External_Iterations = 4
        PSFlash_Maximum_Number_Of_External_Iterations = 5
        PHFlash_Maximum_Number_Of_Internal_Iterations = 6
        PSFlash_Maximum_Number_Of_Internal_Iterations = 7
        PTFlash_Maximum_Number_Of_External_Iterations = 8
        PTFlash_Maximum_Number_Of_Internal_Iterations = 9
        PTFlash_External_Loop_Tolerance = 10
        PTFlash_Internal_Loop_Tolerance = 11
        FlashAlgorithmFastMode = 12
    End Enum

#End Region

    ''' <summary>
    ''' The Property Package Class contains methods to do thermodynamic calculations for all supported phases in DTL.
    ''' </summary>
    ''' <remarks>The base class is inherited by each implemented property package, which contains its own methods.</remarks>
    <System.Serializable()> Public MustInherit Class PropertyPackage

        'CAPE-OPEN 1.0 Interfaces
        Implements ICloneable, ICapeIdentification, ICapeThermoPropertyPackage, ICapeUtilities, ICapeThermoEquilibriumServer, ICapeThermoCalculationRoutine

        'CAPE-OPEN 1.1 Interfaces
        Implements ICapeThermoPhases, ICapeThermoPropertyRoutine, ICapeThermoCompounds, ICapeThermoUniversalConstant
        Implements ICapeThermoMaterialContext, ICapeThermoEquilibriumRoutine

        'CAPE-OPEN Error Interfaces
        Implements ECapeUser, ECapeUnknown, ECapeRoot

        'IDisposable
        Implements IDisposable

        Public Const ClassId As String = ""

        Private m_props As New DTL.SimulationObjects.PropertyPackages.Auxiliary.PROPS
        Public m_Henry As New System.Collections.Generic.Dictionary(Of String, HenryParam)

        Private m_ms As DTL.SimulationObjects.Streams.MaterialStream = Nothing
        Private m_ss As New System.Collections.Generic.List(Of String)
        Private m_configurable As Boolean = False

        Public m_par As System.Collections.Generic.Dictionary(Of String, Double)

        Private _tag As String = ""
        Private _uniqueID As String = ""

        Private m_ip As DataTable

        Private _flashalgorithm As FlashMethod

        Public _packagetype As PackageType

        <System.NonSerialized()> Public _brio3 As Auxiliary.FlashAlgorithms.BostonFournierInsideOut3P
        <System.NonSerialized()> Public _bbio As Auxiliary.FlashAlgorithms.BostonBrittInsideOut
        <System.NonSerialized()> Public _dwdf As Auxiliary.FlashAlgorithms.DWSIMDefault
        <System.NonSerialized()> Public _gm3 As Auxiliary.FlashAlgorithms.GibbsMinimization3P
        <System.NonSerialized()> Public _nl3 As Auxiliary.FlashAlgorithms.NestedLoops3PV3
        <System.NonSerialized()> Public _nlsle As Auxiliary.FlashAlgorithms.NestedLoopsSLE
        <System.NonSerialized()> Public _nli As Auxiliary.FlashAlgorithms.NestedLoopsImmiscible
        <System.NonSerialized()> Public _simplelle As Auxiliary.FlashAlgorithms.SimpleLLE

        Public _ioquick As Boolean = True
        Public _tpseverity As Integer = 0
        Public _tpcompids As String() = New String() {}

        Public _phasemappings As New Dictionary(Of String, PhaseInfo)

        Private LoopVarF, LoopVarX As Double, LoopVarState As State

        Public Property ForceNewFlashAlgorithmInstance As Boolean = False

        <System.NonSerialized()> Private _como As Object 'CAPE-OPEN Material Object

#Region "   Constructor"

        Sub New()
            MyBase.New()
            Initialize()
        End Sub

        Sub New(ByVal capeopenmode As Boolean)

            My.Application.CAPEOPENMode = capeopenmode

            If capeopenmode Then

                'initialize collections

                _selectedcomps = New Dictionary(Of String, ConstantProperties)
                _availablecomps = New Dictionary(Of String, ConstantProperties)

            End If

            Initialize()

        End Sub

        Sub ConfigParameters()
            m_par = New System.Collections.Generic.Dictionary(Of String, Double)
            With Me.Parameters
                .Clear()
                .Add("PP_PHFILT", 0.001)
                .Add("PP_PSFILT", 0.001)
                .Add("PP_PHFELT", 0.001)
                .Add("PP_PSFELT", 0.001)
                .Add("PP_PHFMEI", 50)
                .Add("PP_PSFMEI", 50)
                .Add("PP_PHFMII", 100)
                .Add("PP_PSFMII", 100)
                .Add("PP_PTFMEI", 100)
                .Add("PP_PTFMII", 100)
                .Add("PP_PTFILT", 0.001)
                .Add("PP_PTFELT", 0.001)
                .Add("PP_RIG_BUB_DEW_FLASH_INIT", 0)
                .Add("PP_IDEAL_MIXRULE_LIQDENS", 0)
                .Add("PP_FLASHALGORITHM", 2)
                .Add("PP_FLASHALGORITHMFASTMODE", 1)
                .Add("PP_USEEXPLIQDENS", 0)
                .Add("PP_USEEXPLIQTHERMALCOND", 1)
            End With
        End Sub

        ''' <summary>
        ''' Globally sets a value for the maximum number of iteractions and tolerances for the flash algorithms.
        ''' </summary>
        ''' <param name="p">Parameter to be set.</param>
        ''' <param name="value">Value of the parameter.</param>
        ''' <remarks></remarks>
        Public Sub SetParameterValue(p As Parameter, value As Object)
            Select Case p
                Case Parameter.PHFlash_External_Loop_Tolerance
                    Me.Parameters("PP_PHFELT") = value
                Case Parameter.PHFlash_Internal_Loop_Tolerance
                    Me.Parameters("PP_PHFILT") = value
                Case Parameter.PHFlash_Maximum_Number_Of_External_Iterations
                    Me.Parameters("PP_PHFMEI") = value
                Case Parameter.PHFlash_Maximum_Number_Of_Internal_Iterations
                    Me.Parameters("PP_PHFMII") = value
                Case Parameter.PSFlash_External_Loop_Tolerance
                    Me.Parameters("PP_PSFELT") = value
                Case Parameter.PSFlash_Internal_Loop_Tolerance
                    Me.Parameters("PP_PSFILT") = value
                Case Parameter.PSFlash_Maximum_Number_Of_External_Iterations
                    Me.Parameters("PP_PSFMEI") = value
                Case Parameter.PSFlash_Maximum_Number_Of_Internal_Iterations
                    Me.Parameters("PP_PSFMII") = value
                Case Parameter.PTFlash_External_Loop_Tolerance
                    Me.Parameters("PP_PTFELT") = value
                Case Parameter.PTFlash_Internal_Loop_Tolerance
                    Me.Parameters("PP_PTFILT") = value
                Case Parameter.PTFlash_Maximum_Number_Of_External_Iterations
                    Me.Parameters("PP_PTFMEI") = value
                Case Parameter.PTFlash_Maximum_Number_Of_Internal_Iterations
                    Me.Parameters("PP_PTFMII") = value
                Case Parameter.FlashAlgorithmFastMode
                    Me.Parameters("PP_FLASHALGORITHMFASTMODE") = value
            End Select
        End Sub

        Public Function AddCompound(ByVal compname As String) As Boolean

            If Not _selectedcomps.ContainsKey(compname) Then
                Dim tmpcomp As New ConstantProperties
                tmpcomp = _availablecomps(compname)
                _selectedcomps.Add(tmpcomp.Name, tmpcomp)
                _availablecomps.Remove(tmpcomp.Name)
                Return True
            Else
                Return False
            End If

        End Function

        Public Sub CreatePhaseMappings()
            Me._phasemappings = New Dictionary(Of String, PhaseInfo)
            With Me._phasemappings
                .Add("Vapor", New PhaseInfo("", 2, Fase.Vapor))
                .Add("Liquid1", New PhaseInfo("", 3, Fase.Liquid1))
                .Add("Liquid2", New PhaseInfo("", 4, Fase.Liquid2))
                .Add("Liquid3", New PhaseInfo("", 5, Fase.Liquid3))
                .Add("Aqueous", New PhaseInfo("", 6, Fase.Aqueous))
                .Add("Solid", New PhaseInfo("", 7, Fase.Solid))
            End With
        End Sub

        Public Sub CreatePhaseMappingsDW()
            Me._phasemappings = New Dictionary(Of String, PhaseInfo)
            With Me._phasemappings
                .Add("Vapor", New PhaseInfo("Vapor", 2, Fase.Vapor))
                .Add("Liquid1", New PhaseInfo("Liquid", 3, Fase.Liquid1))
                .Add("Liquid2", New PhaseInfo("Liquid2", 4, Fase.Liquid2))
                .Add("Solid", New PhaseInfo("Solid", 7, Fase.Solid))
            End With
        End Sub

#End Region

#Region "   Properties"

        ''' <summary>
        ''' Get or sets the list of compounds to be used in the liquid phase stability test during three-phase flash calculations.
        ''' </summary>
        ''' <value>A string array containing the names of the compounds.</value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property StabilityTestKeyCompounds As String()
            Get
                Return _tpcompids
            End Get
            Set(value As String())
                _tpcompids = value
            End Set
        End Property

        ''' <summary>
        ''' Defines the severity of the liquid phase stability test during three-phase flash calculations.
        ''' </summary>
        ''' <value>0 is the lowest, 2 is the highest.</value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property StabilityTestSeverity() As Integer
            Get
                Return _tpseverity
            End Get
            Set(value As Integer)
                _tpseverity = value
            End Set
        End Property




        Public ReadOnly Property PhaseMappings() As Dictionary(Of String, PhaseInfo)
            Get
                If Me._phasemappings Is Nothing Then
                    CreatePhaseMappingsDW()
                ElseIf Me._phasemappings.Count = 0 Then
                    CreatePhaseMappingsDW()
                End If
                Return _phasemappings
            End Get
        End Property

        ''' <summary>
        ''' Returns the flash algorithm selected for this property package.
        ''' </summary>
        ''' <value></value>
        ''' <returns>A FlashMethod value with information about the selected flash algorithm.</returns>
        ''' <remarks></remarks>
        Public Property FlashAlgorithm() As FlashMethod
            Get
                Return _flashalgorithm
            End Get
            Set(ByVal value As FlashMethod)
                _flashalgorithm = value
            End Set
        End Property

        ''' <summary>
        ''' Returns the FlashAlgorithm object instance for this property package.
        ''' </summary>
        ''' <value></value>
        ''' <returns>A FlashAlgorithm object to be used in flash calculations.</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property FlashBase() As Auxiliary.FlashAlgorithms.FlashAlgorithm
            Get
                If Not My.Application.CAPEOPENMode And Not My.MyApplication.IsRunningParallelTasks Then
                    If Not Me.Parameters.ContainsKey("PP_FLASHALGORITHM") Then
                        Me.Parameters.Add("PP_FLASHALGORITHM", 2)
                    End If
                    Me.FlashAlgorithm = Me.Parameters("PP_FLASHALGORITHM")
                End If
                Select Case FlashAlgorithm
                    Case FlashMethod.DWSIMDefault
                        If My.MyApplication.IsRunningParallelTasks Or ForceNewFlashAlgorithmInstance Then
                            Return New Auxiliary.FlashAlgorithms.DWSIMDefault
                        Else
                            If _dwdf Is Nothing Then _dwdf = New Auxiliary.FlashAlgorithms.DWSIMDefault
                            Return _dwdf
                        End If
                    Case FlashMethod.InsideOut
                        If My.MyApplication.IsRunningParallelTasks Or ForceNewFlashAlgorithmInstance Then
                            Return New Auxiliary.FlashAlgorithms.BostonBrittInsideOut
                        Else
                            If _bbio Is Nothing Then _bbio = New Auxiliary.FlashAlgorithms.BostonBrittInsideOut
                            Return _bbio
                        End If
                    Case FlashMethod.InsideOut3P
                        If My.MyApplication.IsRunningParallelTasks Or ForceNewFlashAlgorithmInstance Then
                            Return New Auxiliary.FlashAlgorithms.BostonFournierInsideOut3P With
                                                        {.StabSearchCompIDs = _tpcompids, .StabSearchSeverity = _tpseverity}
                        Else
                            If _brio3 Is Nothing Then _brio3 = New Auxiliary.FlashAlgorithms.BostonFournierInsideOut3P With
                                {.StabSearchCompIDs = _tpcompids, .StabSearchSeverity = _tpseverity}
                            Return _brio3
                        End If
                    Case FlashMethod.GibbsMin2P
                        If My.MyApplication.IsRunningParallelTasks Or ForceNewFlashAlgorithmInstance Then
                            Return New Auxiliary.FlashAlgorithms.GibbsMinimization3P With
                                                        {.ForceTwoPhaseOnly = True}
                        Else
                            If _gm3 Is Nothing Then _gm3 = New Auxiliary.FlashAlgorithms.GibbsMinimization3P With {.ForceTwoPhaseOnly = True}
                            Return _gm3
                        End If
                    Case FlashMethod.GibbsMin3P
                        If My.MyApplication.IsRunningParallelTasks Or ForceNewFlashAlgorithmInstance Then
                            Return New Auxiliary.FlashAlgorithms.GibbsMinimization3P With
                                                        {.ForceTwoPhaseOnly = False, .StabSearchCompIDs = _tpcompids, .StabSearchSeverity = _tpseverity}
                        Else
                            If _gm3 Is Nothing Then _gm3 = New Auxiliary.FlashAlgorithms.GibbsMinimization3P With
                                {.ForceTwoPhaseOnly = False, .StabSearchCompIDs = _tpcompids, .StabSearchSeverity = _tpseverity}
                            Return _gm3
                        End If
                    Case FlashMethod.NestedLoops3P, FlashMethod.NestedLoops3PV2, FlashMethod.NestedLoops3PV3
                        If My.MyApplication.IsRunningParallelTasks Or ForceNewFlashAlgorithmInstance Then
                            Return New Auxiliary.FlashAlgorithms.NestedLoops3P With
                                                        {.StabSearchCompIDs = _tpcompids, .StabSearchSeverity = _tpseverity}
                        Else
                            If _nl3 Is Nothing Then _nl3 = New Auxiliary.FlashAlgorithms.NestedLoops3PV3 With
                                {.StabSearchCompIDs = _tpcompids, .StabSearchSeverity = _tpseverity}
                            Return _nl3
                        End If
                    Case FlashMethod.NestedLoopsSLE
                        Dim constprops As New List(Of ConstantProperties)
                        For Each su As Substancia In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                            constprops.Add(su.ConstantProperties)
                        Next
                        If My.MyApplication.IsRunningParallelTasks Or ForceNewFlashAlgorithmInstance Then
                            Return New Auxiliary.FlashAlgorithms.NestedLoopsSLE With {.CompoundProperties = constprops}
                        Else
                            If _nlsle Is Nothing Then _nlsle = New Auxiliary.FlashAlgorithms.NestedLoopsSLE With {.CompoundProperties = constprops}
                            Return _nlsle
                        End If
                    Case FlashMethod.NestedLoopsSLE_SS
                        Dim constprops As New List(Of ConstantProperties)
                        For Each su As Substancia In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                            constprops.Add(su.ConstantProperties)
                        Next
                        If My.MyApplication.IsRunningParallelTasks Or ForceNewFlashAlgorithmInstance Then
                            Return New Auxiliary.FlashAlgorithms.NestedLoopsSLE With {.CompoundProperties = constprops, .SolidSolution = True}
                        Else
                            If _nlsle Is Nothing Then _nlsle = New Auxiliary.FlashAlgorithms.NestedLoopsSLE With {.CompoundProperties = constprops, .SolidSolution = True}
                            Return _nlsle
                        End If
                    Case FlashMethod.NestedLoopsImmiscible
                        Dim constprops As New List(Of ConstantProperties)
                        For Each su As Substancia In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                            constprops.Add(su.ConstantProperties)
                        Next
                        If My.MyApplication.IsRunningParallelTasks Or ForceNewFlashAlgorithmInstance Then
                            Return New Auxiliary.FlashAlgorithms.NestedLoopsImmiscible With
                                                        {.CompoundProperties = constprops, .StabSearchCompIDs = _tpcompids, .StabSearchSeverity = _tpseverity}
                        Else
                            If _nli Is Nothing Then _nli = New Auxiliary.FlashAlgorithms.NestedLoopsImmiscible With
                            {.CompoundProperties = constprops, .StabSearchCompIDs = _tpcompids, .StabSearchSeverity = _tpseverity}
                            Return _nli
                        End If
                    Case FlashMethod.SimpleLLE
                        If My.MyApplication.IsRunningParallelTasks Or ForceNewFlashAlgorithmInstance Then
                            Return New Auxiliary.FlashAlgorithms.SimpleLLE
                        Else
                            If _simplelle Is Nothing Then _simplelle = New Auxiliary.FlashAlgorithms.SimpleLLE
                            Return _simplelle
                        End If
                    Case Else
                        If My.MyApplication.IsRunningParallelTasks Or ForceNewFlashAlgorithmInstance Then
                            Return New Auxiliary.FlashAlgorithms.DWSIMDefault
                        Else
                            If _dwdf Is Nothing Then _dwdf = New Auxiliary.FlashAlgorithms.DWSIMDefault
                            Return _dwdf
                        End If
                End Select
            End Get
        End Property

        Public Property UniqueID() As String
            Get
                Return _uniqueID
            End Get
            Set(ByVal value As String)
                _uniqueID = value
            End Set
        End Property

        Public Property Tag() As String
            Get
                If _tag = "" Then Return Me.ComponentName Else Return _tag
            End Get
            Set(ByVal value As String)
                _tag = value
            End Set
        End Property

        Public ReadOnly Property PackageType() As PackageType
            Get
                Return _packagetype
            End Get
        End Property

        Public Property ParametrosDeInteracao() As DataTable
            Get
                Return m_ip
            End Get
            Set(ByVal value As DataTable)
                m_ip = value
            End Set
        End Property

        Public ReadOnly Property Parameters() As Dictionary(Of String, Double)
            Get
                If Me.m_par Is Nothing Then ConfigParameters()
                Return Me.m_par
            End Get
        End Property

        Public Property IsConfigurable() As Boolean
            Get
                Return m_configurable
            End Get
            Set(ByVal value As Boolean)
                m_configurable = value
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the current material stream for this property package.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend Property CurrentMaterialStream() As MaterialStream
            Get
                Return m_ms
            End Get
            Set(ByVal MatStr As MaterialStream)
                m_ms = MatStr
            End Set
        End Property

        Public ReadOnly Property SupportedComponents() As System.Collections.Generic.List(Of String)
            Get
                Return m_ss
            End Get
        End Property
#End Region

#Region "   Cloning Support"

        Public Function Clone() As Object Implements System.ICloneable.Clone

            Return ObjectCopy(Me)

        End Function

        Function ObjectCopy(ByVal obj As Object) As Object

            Dim objMemStream As New MemoryStream(50000)
            Dim objBinaryFormatter As New BinaryFormatter(Nothing, New StreamingContext(StreamingContextStates.Clone))

            objBinaryFormatter.Serialize(objMemStream, obj)

            objMemStream.Seek(0, SeekOrigin.Begin)

            ObjectCopy = objBinaryFormatter.Deserialize(objMemStream)

            objMemStream.Close()
        End Function
#End Region

#Region "   Must Override or Overridable Functions"

        Public Overridable Sub ShowConfigForm()

        End Sub

        Public Overridable Sub ReconfigureConfigForm()

        End Sub

        ''' <summary>
        ''' Provides a wrapper function for CAPE-OPEN CalcProp/CalcSingleProp functions.
        ''' </summary>
        ''' <param name="property">The property to be calculated.</param>
        ''' <param name="phase">The phase where the property must be calculated for.</param>
        ''' <remarks>This function is necessary since DWSIM's internal property calculation function calculates all properties at once,
        ''' while the CAPE-OPEN standards require that only the property that was asked for to be calculated, leaving the others unchanged.</remarks>
        Public MustOverride Sub DW_CalcProp(ByVal [property] As String, ByVal phase As Fase)
        ''' <summary>
        ''' Provides a default implementation for solid phase property calculations in CAPE-OPEN mode. Should be used by all derived propety packages.
        ''' </summary>
        ''' <remarks></remarks>

        Public Overridable Sub DW_CalcSolidPhaseProps()

            Dim phaseID As Integer = 7
            Dim result As Double = 0.0#

            Dim T, P As Double
            T = Me.CurrentMaterialStream.Fases(0).SPMProperties.temperature
            P = Me.CurrentMaterialStream.Fases(0).SPMProperties.pressure

            result = Me.AUX_SOLIDDENS
            Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.density = result
            Dim constprops As New List(Of ConstantProperties)
            For Each su As Substancia In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                constprops.Add(su.ConstantProperties)
            Next
            result = Me.DW_CalcSolidEnthalpy(T, RET_VMOL(PropertyPackages.Fase.Solid), constprops)
            Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.enthalpy = result
            Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.entropy = result / T
            Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.compressibilityFactor = 0.0#
            result = Me.DW_CalcSolidHeatCapacityCp(T, RET_VMOL(PropertyPackages.Fase.Solid), constprops)
            Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.heatCapacityCp = result
            Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.heatCapacityCv = result
            result = Me.AUX_MMM(Fase.Solid)
            Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.molecularWeight = result
            result = Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.enthalpy.GetValueOrDefault * Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.molecularWeight.GetValueOrDefault
            Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.molar_enthalpy = result
            result = Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.entropy.GetValueOrDefault * Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.molecularWeight.GetValueOrDefault
            Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.molar_entropy = result
            Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.thermalConductivity = 0.0#
            Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.viscosity = 1.0E+20
            Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.kinematic_viscosity = 1.0E+20

        End Sub

        ''' <summary>
        ''' Calculates the enthalpy of a mixture.
        ''' </summary>
        ''' <param name="Vx">Vector of doubles containing the molar composition of the mixture.</param>
        ''' <param name="T">Temperature (K)</param>
        ''' <param name="P">Pressure (Pa)</param>
        ''' <param name="st">State enum indicating the state of the mixture (liquid or vapor).</param>
        ''' <returns>The enthalpy of the mixture in kJ/kg.</returns>
        ''' <remarks>The basis for the calculated enthalpy/entropy in DWSIM is zero at 25 C and 1 atm.</remarks>
        Public MustOverride Function DW_CalcEnthalpy(ByVal Vx As Array, ByVal T As Double, ByVal P As Double, ByVal st As State) As Double

        ''' <summary>
        ''' Calculates the enthalpy departure of a mixture.
        ''' </summary>
        ''' <param name="Vx">Vector of doubles containing the molar composition of the mixture.</param>
        ''' <param name="T">Temperature (K)</param>
        ''' <param name="P">Pressure (Pa)</param>
        ''' <param name="st">State enum indicating the state of the mixture (liquid or vapor).</param>
        ''' <returns>The enthalpy departure of the mixture in kJ/kg.</returns>
        ''' <remarks>The basis for the calculated enthalpy/entropy in DWSIM is zero at 25 C and 1 atm.</remarks>
        Public MustOverride Function DW_CalcEnthalpyDeparture(ByVal Vx As Array, ByVal T As Double, ByVal P As Double, ByVal st As State) As Double

        ''' <summary>
        ''' Calculates the entropy of a mixture.
        ''' </summary>
        ''' <param name="Vx">Vector of doubles containing the molar composition of the mixture.</param>
        ''' <param name="T">Temperature (K)</param>
        ''' <param name="P">Pressure (Pa)</param>
        ''' <param name="st">State enum indicating the state of the mixture (liquid or vapor).</param>
        ''' <returns>The entropy of the mixture in kJ/kg.K.</returns>
        ''' <remarks>The basis for the calculated enthalpy/entropy in DWSIM is zero at 25 C and 1 atm.</remarks>
        Public MustOverride Function DW_CalcEntropy(ByVal Vx As Array, ByVal T As Double, ByVal P As Double, ByVal st As State) As Double

        ''' <summary>
        ''' Calculates the entropy departure of a mixture.
        ''' </summary>
        ''' <param name="Vx">Vector of doubles containing the molar composition of the mixture.</param>
        ''' <param name="T">Temperature (K)</param>
        ''' <param name="P">Pressure (Pa)</param>
        ''' <param name="st">State enum indicating the state of the mixture (liquid or vapor).</param>
        ''' <returns>The entropy departure of the mixture in kJ/kg.K.</returns>
        ''' <remarks>The basis for the calculated enthalpy/entropy in DWSIM is zero at 25 C and 1 atm.</remarks>
        Public MustOverride Function DW_CalcEntropyDeparture(ByVal Vx As Array, ByVal T As Double, ByVal P As Double, ByVal st As State) As Double

        ''' <summary>
        ''' Calculates K-values of components in a mixture.
        ''' </summary>
        ''' <param name="Vx">Vector of doubles containing the molar composition of the liquid phase.</param>
        ''' <param name="Vy">Vector of doubles containing the molar composition of the vapor phase.</param>
        ''' <param name="T">Temperature of the system.</param>
        ''' <param name="P">Pressure of the system.</param>
        ''' <returns>An array containing K-values for all components in the mixture.</returns>
        ''' <remarks>The composition vector must follow the same sequence as the components which were added in the material stream.</remarks>
        Public Overridable Overloads Function DW_CalcKvalue(ByVal Vx As System.Array, ByVal Vy As System.Array, ByVal T As Double, ByVal P As Double, Optional ByVal type As String = "LV") As Object

            Dim fugvap As Object = Nothing
            Dim fugliq As Object = Nothing

            Dim alreadymt As Boolean = False

            If My.MyApplication._EnableParallelProcessing Then
                My.MyApplication.IsRunningParallelTasks = True
                If My.MyApplication._EnableGPUProcessing Then
                    If Not My.MyApplication.gpu.IsMultithreadingEnabled Then
                        My.MyApplication.gpu.EnableMultithreading()
                    Else
                        alreadymt = True
                    End If
                End If
                Try
                    Dim task1 As Task = New Task(Sub()
                                                     fugliq = Me.DW_CalcFugCoeff(Vx, T, P, State.Liquid)
                                                 End Sub)
                    Dim task2 As Task = New Task(Sub()
                                                     If type = "LV" Then
                                                         fugvap = Me.DW_CalcFugCoeff(Vy, T, P, State.Vapor)
                                                     Else ' LL
                                                         fugvap = Me.DW_CalcFugCoeff(Vy, T, P, State.Liquid)
                                                     End If
                                                 End Sub)

                    task1.Start()
                    task2.Start()
                    Task.WaitAll(task1, task2)
                Catch ae As AggregateException
                    For Each ex As Exception In ae.InnerExceptions
                        Throw ex
                    Next
                Finally
                    If My.MyApplication._EnableGPUProcessing Then
                        If Not alreadymt Then
                            My.MyApplication.gpu.DisableMultithreading()
                            My.MyApplication.gpu.FreeAll()
                        End If
                    End If
                End Try
                My.MyApplication.IsRunningParallelTasks = False
            Else
                fugliq = Me.DW_CalcFugCoeff(Vx, T, P, State.Liquid)
                If type = "LV" Then
                    fugvap = Me.DW_CalcFugCoeff(Vy, T, P, State.Vapor)
                Else ' LL
                    fugvap = Me.DW_CalcFugCoeff(Vy, T, P, State.Liquid)
                End If
            End If

            Dim n As Integer = UBound(fugvap)
            Dim i As Integer
            Dim K(n) As Double

            For i = 0 To n
                K(i) = fugliq(i) / fugvap(i)
            Next

            i = 0
            For Each subst As ClassesBasicasTermodinamica.Substancia In Me.CurrentMaterialStream.Fases(1).Componentes.Values
                If K(i) = 0 Or Double.IsInfinity(K(i)) Or Double.IsNaN(K(i)) Then
                    Dim Pc, Tc, w As Double
                    Pc = subst.ConstantProperties.Critical_Pressure
                    Tc = subst.ConstantProperties.Critical_Temperature
                    w = subst.ConstantProperties.Acentric_Factor
                    If type = "LV" Then
                        K(i) = Pc / P * Math.Exp(5.373 * (1 + w) * (1 - Tc / T))
                    Else
                        K(i) = 1.0#
                    End If
                End If
                i += 1
            Next

            If Me.AUX_CheckTrivial(K) Then
                i = 0
                For Each subst As ClassesBasicasTermodinamica.Substancia In Me.CurrentMaterialStream.Fases(1).Componentes.Values
                    Dim Pc, Tc, w As Double
                    Pc = subst.ConstantProperties.Critical_Pressure
                    Tc = subst.ConstantProperties.Critical_Temperature
                    w = subst.ConstantProperties.Acentric_Factor
                    If type = "LV" Then
                        K(i) = Pc / P * Math.Exp(5.373 * (1 + w) * (1 - Tc / T))
                    Else
                        K(i) = 1.0#
                    End If
                    i += 1
                Next
            End If

            Return K

        End Function

        ''' <summary>
        ''' Calculates K-values of components in a mixture.
        ''' </summary>
        ''' <param name="Vx">Vector of doubles containing the molar composition of the mixture.</param>
        ''' <param name="T">Temperature of the system, in K.</param>
        ''' <param name="P">Pressure of the system, in Pa.</param>
        ''' <returns>An array containing K-values for all components in the mixture.</returns>
        ''' <remarks>The composition vector must follow the same sequence as the components which were added in the material stream.</remarks>
        Public Overridable Overloads Function DW_CalcKvalue(ByVal Vx As System.Array, ByVal T As Double, ByVal P As Double) As [Object]

            Dim i As Integer
            Dim result = Me.FlashBase.Flash_PT(Vx, P, T, Me)
            Dim n As Integer = UBound(Vx)
            Dim K(n) As Double

            i = 0
            For Each subst As ClassesBasicasTermodinamica.Substancia In Me.CurrentMaterialStream.Fases(1).Componentes.Values
                K(i) = (result(3)(i) / result(2)(i))
                i += 1
            Next

            i = 0
            For Each subst As ClassesBasicasTermodinamica.Substancia In Me.CurrentMaterialStream.Fases(1).Componentes.Values
                If K(i) = 0 Then K(i) = Me.AUX_PVAPi(subst.Nome, T) / P
                If Double.IsInfinity(K(i)) Or Double.IsNaN(K(i)) Then
                    Dim Pc, Tc, w As Double
                    Pc = subst.ConstantProperties.Critical_Pressure
                    Tc = subst.ConstantProperties.Critical_Temperature
                    w = subst.ConstantProperties.Acentric_Factor
                    K(i) = Pc / P * Math.Exp(5.373 * (1 + w) * (1 - Tc / T))
                End If
                i += 1
            Next

            If Me.AUX_CheckTrivial(K) Then
                For i = 0 To UBound(Vx)
                    K(i) = Me.AUX_PVAPi(i, T) / P
                    i += 1
                Next
            End If

            Return K

        End Function

        ''' <summary>
        ''' Does a Bubble Pressure calculation for the specified liquid composition at the specified temperature.
        ''' </summary>
        ''' <param name="Vx">Vector of doubles containing liquid phase molar composition for each component in the mixture.</param>
        ''' <param name="T">Temperature in K</param>
        ''' <param name="Pref">Initial estimate for Pressure, in Pa</param>
        ''' <param name="K">Vector with initial estimates for K-values</param>
        ''' <param name="ReuseK">Boolean indicating wether to use the initial estimates for K-values or not.</param>
        ''' <returns>Returns the object vector {L, V, Vx, Vy, P, ecount, Ki} where L is the liquid phase molar fraction, 
        ''' V is the vapor phase molar fraction, Vx is the liquid phase molar composition vector, Vy is the vapor phase molar 
        ''' composition vector, P is the calculated Pressure in Pa, ecount is the number of iterations and Ki is a vector containing 
        ''' the calculated K-values.</returns>
        ''' <remarks>The composition vector must follow the same sequence as the components which were added in the material stream.</remarks>
        Public Overridable Function DW_CalcBubP(ByVal Vx As System.Array, ByVal T As Double, Optional ByVal Pref As Double = 0, Optional ByVal K As System.Array = Nothing, Optional ByVal ReuseK As Boolean = False) As Object
            Return Me.FlashBase.Flash_TV(Vx, T, 0, Pref, Me, ReuseK, K)
        End Function

        ''' <summary>
        ''' Does a Bubble Temperature calculation for the specified liquid composition at the specified pressure.
        ''' </summary>
        ''' <param name="Vx">Vector of doubles containing liquid phase molar composition for each component in the mixture.</param>
        ''' <param name="P"></param>
        ''' <param name="Tref"></param>
        ''' <param name="K">Vector with initial estimates for K-values</param>
        ''' <param name="ReuseK">Boolean indicating wether to use the initial estimates for K-values or not.</param>
        ''' <returns>Returns the object vector {L, V, Vx, Vy, T, ecount, Ki} where L is the liquid phase molar fraction, 
        ''' V is the vapor phase molar fraction, Vx is the liquid phase molar composition vector, Vy is the vapor phase molar 
        ''' composition vector, T is the calculated Temperature in K, ecount is the number of iterations and Ki is a vector containing 
        ''' the calculated K-values.</returns>
        ''' <remarks>The composition vector must follow the same sequence as the components which were added in the material stream.</remarks>
        Public Overridable Function DW_CalcBubT(ByVal Vx As System.Array, ByVal P As Double, Optional ByVal Tref As Double = 0, Optional ByVal K As System.Array = Nothing, Optional ByVal ReuseK As Boolean = False) As Object
            Return Me.FlashBase.Flash_PV(Vx, P, 0, Tref, Me, ReuseK, K)
        End Function

        ''' <summary>
        ''' Does a Dew Pressure calculation for the specified vapor phase composition at the specified temperature.
        ''' </summary>
        ''' <param name="Vx">Vector of doubles containing vapor phase molar composition for each component in the mixture.</param>
        ''' <param name="T">Temperature in K</param>
        ''' <param name="Pref">Initial estimate for Pressure, in Pa</param>
        ''' <param name="K">Vector with initial estimates for K-values</param>
        ''' <param name="ReuseK">Boolean indicating wether to use the initial estimates for K-values or not.</param>
        ''' <returns>Returns the object vector {L, V, Vx, Vy, P, ecount, Ki} where L is the liquid phase molar fraction, 
        ''' V is the vapor phase molar fraction, Vx is the liquid phase molar composition vector, Vy is the vapor phase molar 
        ''' composition vector, P is the calculated Pressure in Pa, ecount is the number of iterations and Ki is a vector containing 
        ''' the calculated K-values.</returns>
        ''' <remarks>The composition vector must follow the same sequence as the components which were added in the material stream.</remarks>
        Public Overridable Function DW_CalcDewP(ByVal Vx As System.Array, ByVal T As Double, Optional ByVal Pref As Double = 0, Optional ByVal K As System.Array = Nothing, Optional ByVal ReuseK As Boolean = False) As Object
            Return Me.FlashBase.Flash_TV(Vx, T, 1, Pref, Me, ReuseK, K)
        End Function

        ''' <summary>
        ''' Does a Dew Temperature calculation for the specified vapor composition at the specified pressure.
        ''' </summary>
        ''' <param name="Vx">Vector of doubles containing vapor phase molar composition for each component in the mixture.</param>
        ''' <param name="P"></param>
        ''' <param name="Tref"></param>
        ''' <param name="K">Vector with initial estimates for K-values</param>
        ''' <param name="ReuseK">Boolean indicating wether to use the initial estimates for K-values or not.</param>
        ''' <returns>Returns the object vector {L, V, Vx, Vy, T, ecount, Ki} where L is the liquid phase molar fraction, 
        ''' V is the vapor phase molar fraction, Vx is the liquid phase molar composition vector, Vy is the vapor phase molar 
        ''' composition vector, T is the calculated Temperature in K, ecount is the number of iterations and Ki is a vector containing 
        ''' the calculated K-values.</returns>
        ''' <remarks>The composition vector must follow the same sequence as the components which were added in the material stream.</remarks>
        Public Overridable Function DW_CalcDewT(ByVal Vx As System.Array, ByVal P As Double, Optional ByVal Tref As Double = 0, Optional ByVal K As System.Array = Nothing, Optional ByVal ReuseK As Boolean = False) As Object
            Return Me.FlashBase.Flash_PV(Vx, P, 1, Tref, Me, ReuseK, K)
        End Function

        ''' <summary>
        ''' Calculates fugacity coefficients for the specified composition at the specified conditions.
        ''' </summary>
        ''' <param name="Vx">Vector of doubles containing the molar composition of the mixture.</param>
        ''' <param name="T">Temperature in K</param>
        ''' <param name="P">Pressure in Pa</param>
        ''' <param name="st">Mixture state (Liquid or Vapor)</param>
        ''' <returns>A vector of doubles containing fugacity coefficients for the components in the mixture.</returns>
        ''' <remarks>The composition vector must follow the same sequence as the components which were added in the material stream.</remarks>
        Public MustOverride Function DW_CalcFugCoeff(ByVal Vx As Array, ByVal T As Double, ByVal P As Double, ByVal st As State) As Double()

        Public MustOverride Function SupportsComponent(ByVal comp As DTL.ClassesBasicasTermodinamica.ConstantProperties) As Boolean

        Public MustOverride Sub DW_CalcPhaseProps(ByVal fase As Fase)

        Public Overridable Sub DW_CalcTwoPhaseProps(ByVal fase1 As Fase, ByVal fase2 As Fase)

            Dim T As Double

            T = Me.CurrentMaterialStream.Fases(0).SPMProperties.temperature
            Me.CurrentMaterialStream.Fases(0).TPMProperties.surfaceTension = Me.AUX_SURFTM(T)

        End Sub

        Public Function DW_CalcGibbsEnergy(ByVal Vx As System.Array, ByVal T As Double, ByVal P As Double) As Double

            Dim fugvap As Object = Nothing
            Dim fugliq As Object = Nothing

            If My.MyApplication._EnableParallelProcessing Then
                My.MyApplication.IsRunningParallelTasks = True
                Try
                    Dim task1 As Task = New Task(Sub()
                                                     fugliq = Me.DW_CalcFugCoeff(Vx, T, P, State.Liquid)
                                                 End Sub)
                    Dim task2 As Task = New Task(Sub()
                                                     fugvap = Me.DW_CalcFugCoeff(Vx, T, P, State.Vapor)
                                                 End Sub)
                    task1.Start()
                    task2.Start()
                    Task.WaitAll(task1, task2)
                Catch ae As AggregateException
                    For Each ex As Exception In ae.InnerExceptions
                        Throw ex
                    Next
                Finally
                End Try
                My.MyApplication.IsRunningParallelTasks = False
            Else
                fugliq = Me.DW_CalcFugCoeff(Vx, T, P, State.Liquid)
                fugvap = Me.DW_CalcFugCoeff(Vx, T, P, State.Vapor)
            End If

            Dim n As Integer = UBound(Vx)
            Dim i As Integer

            Dim g, gid, gexv, gexl As Double

            gid = 0.0#

            'If MathEx.Common.Sum(Vx) <> 0.0# Then
            '    gid = RET_Gid(298.15, T, P, Vx) * AUX_MMM(Vx) 'kJ/kmol
            'End If

            gexv = 0.0#
            gexl = 0.0#
            For i = 0 To n
                If Vx(i) <> 0.0# Then gexv += Vx(i) * Log(fugvap(i)) * 8.314 * T
                If Vx(i) <> 0.0# Then gexl += Vx(i) * Log(fugliq(i)) * 8.314 * T
            Next

            If gexv < gexl Then g = gid + gexv Else g = gid + gexl

            Return g 'kJ/kmol

        End Function

        Public Overridable Sub DW_CalcOverallProps()

            Dim HL, HV, HS, SL, SV, SS, DL, DV, DS, CPL, CPV, CPS, KL, KV, KS, CVL, CVV, CSV As Nullable(Of Double)
            Dim xl, xv, xs, wl, wv, ws, vl, vv, vs, result As Double

            xl = Me.CurrentMaterialStream.Fases(1).SPMProperties.molarfraction.GetValueOrDefault
            xv = Me.CurrentMaterialStream.Fases(2).SPMProperties.molarfraction.GetValueOrDefault
            xs = Me.CurrentMaterialStream.Fases(7).SPMProperties.molarfraction.GetValueOrDefault

            wl = Me.CurrentMaterialStream.Fases(1).SPMProperties.massfraction.GetValueOrDefault
            wv = Me.CurrentMaterialStream.Fases(2).SPMProperties.massfraction.GetValueOrDefault
            ws = Me.CurrentMaterialStream.Fases(7).SPMProperties.massfraction.GetValueOrDefault

            DL = Me.CurrentMaterialStream.Fases(1).SPMProperties.density.GetValueOrDefault
            DV = Me.CurrentMaterialStream.Fases(2).SPMProperties.density.GetValueOrDefault
            DS = Me.CurrentMaterialStream.Fases(7).SPMProperties.density.GetValueOrDefault

            Dim tl As Double = 0.0#
            Dim tv As Double = 0.0#
            Dim ts As Double = 0.0#

            If DL <> 0.0# Then tl = Me.CurrentMaterialStream.Fases(1).SPMProperties.massfraction.GetValueOrDefault / DL.GetValueOrDefault
            If DV <> 0.0# Then tv = Me.CurrentMaterialStream.Fases(2).SPMProperties.massfraction.GetValueOrDefault / DV.GetValueOrDefault
            If DS <> 0.0# Then ts = Me.CurrentMaterialStream.Fases(7).SPMProperties.massfraction.GetValueOrDefault / DS.GetValueOrDefault

            vl = tl / (tl + tv + ts)
            vv = tv / (tl + tv + ts)
            vs = ts / (tl + tv + ts)

            If xl = 1 Then
                vl = 1
                vv = 0
            ElseIf xl = 0 Then
                vl = 0
                vv = 1
            End If

            result = vl * DL.GetValueOrDefault + vv * DV.GetValueOrDefault + vs * DS.GetValueOrDefault
            If Double.IsNaN(result) Then
                If Double.IsNaN(DL) = False And Double.IsNaN(DV) = True Then
                    result = DL
                ElseIf Double.IsNaN(DL) = True And Double.IsNaN(DV) = False Then
                    result = DV
                Else
                    result = 0
                End If
            End If
            Me.CurrentMaterialStream.Fases(0).SPMProperties.density = result

            HL = Me.CurrentMaterialStream.Fases(1).SPMProperties.enthalpy.GetValueOrDefault
            HV = Me.CurrentMaterialStream.Fases(2).SPMProperties.enthalpy.GetValueOrDefault
            HS = Me.CurrentMaterialStream.Fases(7).SPMProperties.enthalpy.GetValueOrDefault

            result = wl * HL.GetValueOrDefault + wv * HV.GetValueOrDefault + ws * HS.GetValueOrDefault
            Me.CurrentMaterialStream.Fases(0).SPMProperties.enthalpy = result

            SL = Me.CurrentMaterialStream.Fases(1).SPMProperties.entropy.GetValueOrDefault
            SV = Me.CurrentMaterialStream.Fases(2).SPMProperties.entropy.GetValueOrDefault
            SS = Me.CurrentMaterialStream.Fases(7).SPMProperties.entropy.GetValueOrDefault

            result = wl * SL.GetValueOrDefault + wv * SV.GetValueOrDefault + ws * SS.GetValueOrDefault
            Me.CurrentMaterialStream.Fases(0).SPMProperties.entropy = result

            Me.CurrentMaterialStream.Fases(0).SPMProperties.compressibilityFactor = Nothing

            CPL = Me.CurrentMaterialStream.Fases(1).SPMProperties.heatCapacityCp.GetValueOrDefault
            CPV = Me.CurrentMaterialStream.Fases(2).SPMProperties.heatCapacityCp.GetValueOrDefault
            CPS = Me.CurrentMaterialStream.Fases(7).SPMProperties.heatCapacityCp.GetValueOrDefault

            result = wl * CPL.GetValueOrDefault + wv * CPV.GetValueOrDefault + ws * CPS.GetValueOrDefault
            Me.CurrentMaterialStream.Fases(0).SPMProperties.heatCapacityCp = result

            CVL = Me.CurrentMaterialStream.Fases(1).SPMProperties.heatCapacityCv.GetValueOrDefault
            CVV = Me.CurrentMaterialStream.Fases(2).SPMProperties.heatCapacityCv.GetValueOrDefault
            CSV = Me.CurrentMaterialStream.Fases(7).SPMProperties.heatCapacityCv.GetValueOrDefault

            result = wl * CVL.GetValueOrDefault + wv * CVV.GetValueOrDefault + ws * CSV.GetValueOrDefault
            Me.CurrentMaterialStream.Fases(0).SPMProperties.heatCapacityCv = result

            result = Me.AUX_MMM(Fase.Mixture)
            Me.CurrentMaterialStream.Fases(0).SPMProperties.molecularWeight = result

            result = Me.CurrentMaterialStream.Fases(0).SPMProperties.enthalpy.GetValueOrDefault * Me.CurrentMaterialStream.Fases(0).SPMProperties.molecularWeight.GetValueOrDefault
            Me.CurrentMaterialStream.Fases(0).SPMProperties.molar_enthalpy = result
            result = Me.CurrentMaterialStream.Fases(0).SPMProperties.entropy.GetValueOrDefault * Me.CurrentMaterialStream.Fases(0).SPMProperties.molecularWeight.GetValueOrDefault
            Me.CurrentMaterialStream.Fases(0).SPMProperties.molar_entropy = result

            KL = Me.CurrentMaterialStream.Fases(1).SPMProperties.thermalConductivity.GetValueOrDefault
            KV = Me.CurrentMaterialStream.Fases(2).SPMProperties.thermalConductivity.GetValueOrDefault
            KS = Me.CurrentMaterialStream.Fases(7).SPMProperties.thermalConductivity.GetValueOrDefault

            result = xl * KL.GetValueOrDefault + xv * KV.GetValueOrDefault + xs * KS.GetValueOrDefault
            Me.CurrentMaterialStream.Fases(0).SPMProperties.thermalConductivity = result

            Me.CurrentMaterialStream.Fases(0).SPMProperties.viscosity = Nothing

            Me.CurrentMaterialStream.Fases(0).SPMProperties.kinematic_viscosity = Nothing

        End Sub

        Public Overridable Sub DW_CalcLiqMixtureProps()

            Dim hl, hl1, hl2, hl3, hw, sl, sl1, sl2, sl3, sw, dl, dl1, dl2, dl3, dw As Double
            Dim cpl, cpl1, cpl2, cpl3, cpw, cvl, cvl1, cvl2, cvl3, cvw As Double
            Dim kl, kl1, kl2, kl3, kw, vil, vil1, vil2, vil3, viw As Double
            Dim xl, xl1, xl2, xl3, xw, wl, wl1, wl2, wl3, ww As Double
            Dim xlf, xlf1, xlf2, xlf3, xwf, wlf, wlf1, wlf2, wlf3, wwf As Double
            Dim cml, cml1, cml2, cml3, cmw, cwl, cwl1, cwl2, cwl3, cww As Double

            xl1 = Me.CurrentMaterialStream.Fases(3).SPMProperties.molarfraction.GetValueOrDefault
            xl2 = Me.CurrentMaterialStream.Fases(4).SPMProperties.molarfraction.GetValueOrDefault
            xl3 = Me.CurrentMaterialStream.Fases(5).SPMProperties.molarfraction.GetValueOrDefault
            xw = Me.CurrentMaterialStream.Fases(6).SPMProperties.molarfraction.GetValueOrDefault

            xl = xl1 + xl2 + xl3 + xw
            Me.CurrentMaterialStream.Fases(1).SPMProperties.molarfraction = xl

            wl1 = Me.CurrentMaterialStream.Fases(3).SPMProperties.massfraction.GetValueOrDefault
            wl2 = Me.CurrentMaterialStream.Fases(4).SPMProperties.massfraction.GetValueOrDefault
            wl3 = Me.CurrentMaterialStream.Fases(5).SPMProperties.massfraction.GetValueOrDefault
            ww = Me.CurrentMaterialStream.Fases(6).SPMProperties.massfraction.GetValueOrDefault

            wl = wl1 + wl2 + wl3 + ww
            Me.CurrentMaterialStream.Fases(1).SPMProperties.massfraction = wl

            xlf1 = Me.CurrentMaterialStream.Fases(3).SPMProperties.molarflow.GetValueOrDefault
            xlf2 = Me.CurrentMaterialStream.Fases(4).SPMProperties.molarflow.GetValueOrDefault
            xlf3 = Me.CurrentMaterialStream.Fases(5).SPMProperties.molarflow.GetValueOrDefault
            xwf = Me.CurrentMaterialStream.Fases(6).SPMProperties.molarflow.GetValueOrDefault

            xlf = xlf1 + xlf2 + xlf3 + xwf
            Me.CurrentMaterialStream.Fases(1).SPMProperties.molarflow = xlf

            wlf1 = Me.CurrentMaterialStream.Fases(3).SPMProperties.massflow.GetValueOrDefault
            wlf2 = Me.CurrentMaterialStream.Fases(4).SPMProperties.massflow.GetValueOrDefault
            wlf3 = Me.CurrentMaterialStream.Fases(5).SPMProperties.massflow.GetValueOrDefault
            wwf = Me.CurrentMaterialStream.Fases(6).SPMProperties.massflow.GetValueOrDefault

            wlf = wlf1 + wlf2 + wlf3 + wwf
            Me.CurrentMaterialStream.Fases(1).SPMProperties.massflow = wlf

            For Each c As Substancia In Me.CurrentMaterialStream.Fases(1).Componentes.Values
                cml1 = Me.CurrentMaterialStream.Fases(3).Componentes(c.Nome).FracaoMolar.GetValueOrDefault
                cml2 = Me.CurrentMaterialStream.Fases(4).Componentes(c.Nome).FracaoMolar.GetValueOrDefault
                cml3 = Me.CurrentMaterialStream.Fases(5).Componentes(c.Nome).FracaoMolar.GetValueOrDefault
                cmw = Me.CurrentMaterialStream.Fases(6).Componentes(c.Nome).FracaoMolar.GetValueOrDefault
                cwl1 = Me.CurrentMaterialStream.Fases(3).Componentes(c.Nome).FracaoMassica.GetValueOrDefault
                cwl2 = Me.CurrentMaterialStream.Fases(4).Componentes(c.Nome).FracaoMassica.GetValueOrDefault
                cwl3 = Me.CurrentMaterialStream.Fases(5).Componentes(c.Nome).FracaoMassica.GetValueOrDefault
                cww = Me.CurrentMaterialStream.Fases(6).Componentes(c.Nome).FracaoMassica.GetValueOrDefault
                cml = (xl1 * cml1 + xl2 * cml2 + xl3 * cml3 + xw * cmw) / xl
                cwl = (wl1 * cwl1 + wl2 * cwl2 + wl3 * cwl3 + ww * cww) / wl
                c.FracaoMolar = cml
                c.FracaoMassica = cwl
            Next

            If wl1 > 0 Then dl1 = Me.CurrentMaterialStream.Fases(3).SPMProperties.density.GetValueOrDefault Else dl1 = 1
            If wl2 > 0 Then dl2 = Me.CurrentMaterialStream.Fases(4).SPMProperties.density.GetValueOrDefault Else dl2 = 1
            If wl3 > 0 Then dl3 = Me.CurrentMaterialStream.Fases(5).SPMProperties.density.GetValueOrDefault Else dl3 = 1
            If ww > 0 Then dw = Me.CurrentMaterialStream.Fases(6).SPMProperties.density.GetValueOrDefault Else dw = 1

            dl = wl1 / dl1 + wl2 / dl2 + wl3 / dl3 + ww / dw
            dl = wl / dl
            Me.CurrentMaterialStream.Fases(1).SPMProperties.density = dl


            If Double.IsNaN(wlf / dl) Then
                Me.CurrentMaterialStream.Fases(1).SPMProperties.volumetric_flow = 0.0#
            Else
                Me.CurrentMaterialStream.Fases(1).SPMProperties.volumetric_flow = wlf / dl
            End If

            If wl = 0 Then wl = 1.0#

            hl1 = Me.CurrentMaterialStream.Fases(3).SPMProperties.enthalpy.GetValueOrDefault
            hl2 = Me.CurrentMaterialStream.Fases(4).SPMProperties.enthalpy.GetValueOrDefault
            hl3 = Me.CurrentMaterialStream.Fases(5).SPMProperties.enthalpy.GetValueOrDefault
            hw = Me.CurrentMaterialStream.Fases(6).SPMProperties.enthalpy.GetValueOrDefault

            hl = hl1 * wl1 + hl2 * wl2 + hl3 * wl3 + hw * ww

            Me.CurrentMaterialStream.Fases(1).SPMProperties.enthalpy = hl / wl

            sl1 = Me.CurrentMaterialStream.Fases(3).SPMProperties.entropy.GetValueOrDefault
            sl2 = Me.CurrentMaterialStream.Fases(4).SPMProperties.entropy.GetValueOrDefault
            sl3 = Me.CurrentMaterialStream.Fases(5).SPMProperties.entropy.GetValueOrDefault
            sw = Me.CurrentMaterialStream.Fases(6).SPMProperties.entropy.GetValueOrDefault

            sl = sl1 * wl1 + sl2 * wl2 + sl3 * wl3 + sw * ww

            Me.CurrentMaterialStream.Fases(1).SPMProperties.entropy = sl / wl

            cpl1 = Me.CurrentMaterialStream.Fases(3).SPMProperties.heatCapacityCp.GetValueOrDefault
            cpl2 = Me.CurrentMaterialStream.Fases(4).SPMProperties.heatCapacityCp.GetValueOrDefault
            cpl3 = Me.CurrentMaterialStream.Fases(5).SPMProperties.heatCapacityCp.GetValueOrDefault
            cpw = Me.CurrentMaterialStream.Fases(6).SPMProperties.heatCapacityCp.GetValueOrDefault

            cpl = cpl1 * wl1 + cpl2 * wl2 + cpl3 * wl3 + cpw * ww

            Me.CurrentMaterialStream.Fases(1).SPMProperties.heatCapacityCp = cpl / wl

            cvl1 = Me.CurrentMaterialStream.Fases(3).SPMProperties.heatCapacityCv.GetValueOrDefault
            cvl2 = Me.CurrentMaterialStream.Fases(4).SPMProperties.heatCapacityCv.GetValueOrDefault
            cvl3 = Me.CurrentMaterialStream.Fases(5).SPMProperties.heatCapacityCv.GetValueOrDefault
            cvw = Me.CurrentMaterialStream.Fases(6).SPMProperties.heatCapacityCv.GetValueOrDefault

            cvl = cvl1 * wl1 + cvl2 * wl2 + cvl3 * wl3 + cvw * ww

            Me.CurrentMaterialStream.Fases(1).SPMProperties.heatCapacityCv = cvl / wl

            Dim result As Double

            result = Me.AUX_MMM(Fase.Liquid)
            Me.CurrentMaterialStream.Fases(1).SPMProperties.molecularWeight = result

            result = Me.CurrentMaterialStream.Fases(1).SPMProperties.enthalpy.GetValueOrDefault * Me.CurrentMaterialStream.Fases(1).SPMProperties.molecularWeight.GetValueOrDefault
            Me.CurrentMaterialStream.Fases(1).SPMProperties.molar_enthalpy = result

            result = Me.CurrentMaterialStream.Fases(1).SPMProperties.entropy.GetValueOrDefault * Me.CurrentMaterialStream.Fases(1).SPMProperties.molecularWeight.GetValueOrDefault
            Me.CurrentMaterialStream.Fases(1).SPMProperties.molar_entropy = result

            kl1 = Me.CurrentMaterialStream.Fases(3).SPMProperties.thermalConductivity.GetValueOrDefault
            kl2 = Me.CurrentMaterialStream.Fases(4).SPMProperties.thermalConductivity.GetValueOrDefault
            kl3 = Me.CurrentMaterialStream.Fases(5).SPMProperties.thermalConductivity.GetValueOrDefault
            kw = Me.CurrentMaterialStream.Fases(6).SPMProperties.thermalConductivity.GetValueOrDefault

            kl = kl1 * wl1 + kl2 * wl2 + kl3 * kl3 + kw * ww

            Me.CurrentMaterialStream.Fases(1).SPMProperties.thermalConductivity = kl / wl

            vil1 = Me.CurrentMaterialStream.Fases(3).SPMProperties.viscosity.GetValueOrDefault
            vil2 = Me.CurrentMaterialStream.Fases(4).SPMProperties.viscosity.GetValueOrDefault
            vil3 = Me.CurrentMaterialStream.Fases(5).SPMProperties.viscosity.GetValueOrDefault
            viw = Me.CurrentMaterialStream.Fases(6).SPMProperties.viscosity.GetValueOrDefault

            vil = vil1 * wl1 + vil2 * vil2 + kl3 * vil3 + viw * ww

            Me.CurrentMaterialStream.Fases(1).SPMProperties.viscosity = vil / wl

            Me.CurrentMaterialStream.Fases(1).SPMProperties.kinematic_viscosity = vil / dl

            Me.CurrentMaterialStream.Fases(1).SPMProperties.compressibilityFactor = 0.0#

        End Sub

        Public Overridable Sub DW_CalcEquilibrium(ByVal spec1 As FlashSpec, ByVal spec2 As FlashSpec)

            Me.CurrentMaterialStream.AtEquilibrium = False

            Dim P, T, H, S, xv, xl, xl2, xs As Double
            Dim result As Object = Nothing
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim n As Integer = Me.CurrentMaterialStream.Fases(0).Componentes.Count
            Dim i As Integer = 0

            'for TVF/PVF/PH/PS flashes
            xv = Me.CurrentMaterialStream.Fases(2).SPMProperties.molarfraction.GetValueOrDefault
            H = Me.CurrentMaterialStream.Fases(0).SPMProperties.enthalpy.GetValueOrDefault
            S = Me.CurrentMaterialStream.Fases(0).SPMProperties.entropy.GetValueOrDefault

            Me.DW_ZerarPhaseProps(Fase.Vapor)
            Me.DW_ZerarPhaseProps(Fase.Liquid)
            Me.DW_ZerarPhaseProps(Fase.Liquid1)
            Me.DW_ZerarPhaseProps(Fase.Liquid2)
            Me.DW_ZerarPhaseProps(Fase.Liquid3)
            Me.DW_ZerarPhaseProps(Fase.Aqueous)
            Me.DW_ZerarPhaseProps(Fase.Solid)
            Me.DW_ZerarComposicoes(Fase.Vapor)
            Me.DW_ZerarComposicoes(Fase.Liquid)
            Me.DW_ZerarComposicoes(Fase.Liquid1)
            Me.DW_ZerarComposicoes(Fase.Liquid2)
            Me.DW_ZerarComposicoes(Fase.Liquid3)
            Me.DW_ZerarComposicoes(Fase.Aqueous)
            Me.DW_ZerarComposicoes(Fase.Solid)

            Select Case spec1

                Case FlashSpec.T

                    Select Case spec2

                        Case FlashSpec.P

                            T = Me.CurrentMaterialStream.Fases(0).SPMProperties.temperature.GetValueOrDefault
                            P = Me.CurrentMaterialStream.Fases(0).SPMProperties.pressure.GetValueOrDefault

                            Dim ige As Double = 0
                            Dim fge As Double = 0
                            Dim dge As Double = 0

                            If Not Me.FlashAlgorithm = FlashMethod.NestedLoopsSLE _
                            And Not Me.FlashAlgorithm = FlashMethod.NestedLoopsSLE_SS Then

                                ige = Me.DW_CalcGibbsEnergy(RET_VMOL(Fase.Mixture), T, P)

                            End If
                           
                            result = Me.FlashBase.Flash_PT(RET_VMOL(Fase.Mixture), P, T, Me)

                            xl = result(0)
                            xv = result(1)
                            xl2 = result(5)
                            xs = result(7)

                            Dim Vx = result(2)
                            Dim Vy = result(3)
                            Dim Vx2 = result(6)
                            Dim Vs = result(8)

                            'identify phase
                            If Me.ComponentName.Contains("SRK") Or Me.ComponentName.Contains("PR") Then
                                If Not Me.AUX_IS_SINGLECOMP(Fase.Mixture) Then
                                    Dim newphase, eos As String
                                    If Me.ComponentName.Contains("SRK") Then eos = "SRK" Else eos = "PR"
                                    If xv = 1.0# Or xl = 1.0# Then
                                        If xv = 1.0# Then
                                            newphase = Auxiliary.FlashAlgorithms.FlashAlgorithm.IdentifyPhase(Vy, P, T, Me, eos)
                                            If newphase = "L" Then
                                                xv = 0.0#
                                                xl = 1.0#
                                                Vx = Vy
                                            End If
                                        Else
                                            newphase = Auxiliary.FlashAlgorithms.FlashAlgorithm.IdentifyPhase(Vx, P, T, Me, eos)
                                            If newphase = "V" Then
                                                xv = 1.0#
                                                xl = 0.0#
                                                Vy = Vx
                                            End If
                                        End If
                                    Else
                                        If xl2 = 0.0# Then
                                            newphase = Auxiliary.FlashAlgorithms.FlashAlgorithm.IdentifyPhase(Vy, P, T, Me, eos)
                                            If newphase = "L" Then
                                                xl2 = xv
                                                xv = 0.0#
                                                Vx2 = Vy
                                            End If
                                            newphase = Auxiliary.FlashAlgorithms.FlashAlgorithm.IdentifyPhase(Vx, P, T, Me, eos)
                                            If newphase = "V" Then
                                                xv = 1.0#
                                                xl = 0.0#
                                                xl2 = 0.0#
                                                Vy = RET_VMOL(Fase.Mixture)
                                            End If
                                        ElseIf xv = 0.0# Then
                                            newphase = Auxiliary.FlashAlgorithms.FlashAlgorithm.IdentifyPhase(Vx2, P, T, Me, eos)
                                            If newphase = "V" Then
                                                xv = xl2
                                                xl2 = 0.0#
                                                Vy = Vx2
                                            End If
                                        End If
                                    End If
                                End If
                            End If

                            If Not Me.FlashAlgorithm = FlashMethod.NestedLoopsSLE _
                            And Not Me.FlashAlgorithm = FlashMethod.NestedLoopsSLE_SS Then

                                fge = xl * Me.DW_CalcGibbsEnergy(Vx, T, P)
                                fge += xl2 * Me.DW_CalcGibbsEnergy(Vx2, T, P)
                                fge += xv * Me.DW_CalcGibbsEnergy(Vy, T, P)

                                dge = fge - ige

                                Dim dgtol As Double = 0.01

                                If dge > 0.0# And Math.Abs(dge / ige * 100) > Math.Abs(dgtol) Then
                                    Throw New Exception(DTL.App.GetLocalString("InvalidFlashResult") & "(DGE = " & dge & " kJ/kg, " & Format(dge / ige * 100, "0.00") & "%)")
                                End If

                            End If

                            'do a density calculation check to order liquid phases from lighter to heavier

                            If xl2 <> 0.0# And xl = 0.0# Then
                                xl = result(5)
                                xl2 = 0.0#
                                Vx = result(6)
                                Vx2 = result(2)
                            ElseIf xl2 <> 0.0# And xl <> 0.0# Then
                                Dim dens1, dens2, xl0, xl20, Vx0(), Vx20() As Double
                                dens1 = Me.AUX_LIQDENS(T, Vx, P, 0, False)
                                dens2 = Me.AUX_LIQDENS(T, Vx2, P, 0, False)
                                If dens2 < dens1 Then
                                    xl0 = xl
                                    xl20 = xl2
                                    Vx0 = Vx
                                    Vx20 = Vx2
                                    xl = xl20
                                    xl2 = xl0
                                    Vx = Vx20
                                    Vx2 = Vx0
                                End If
                            End If

                            Me.CurrentMaterialStream.Fases(3).SPMProperties.molarfraction = xl
                            Me.CurrentMaterialStream.Fases(4).SPMProperties.molarfraction = xl2
                            Me.CurrentMaterialStream.Fases(2).SPMProperties.molarfraction = xv
                            Me.CurrentMaterialStream.Fases(7).SPMProperties.molarfraction = xs

                            Dim FCL = Me.DW_CalcFugCoeff(Vx, T, P, State.Liquid)
                            Dim FCL2 = Me.DW_CalcFugCoeff(Vx2, T, P, State.Liquid)
                            Dim FCV = Me.DW_CalcFugCoeff(Vy, T, P, State.Vapor)
                            Dim FCS = Me.DW_CalcFugCoeff(Vs, T, P, State.Solid)

                            i = 0
                            For Each subst In Me.CurrentMaterialStream.Fases(3).Componentes.Values
                                subst.FracaoMolar = Vx(i)
                                subst.FracaoMassica = Me.AUX_CONVERT_MOL_TO_MASS(Vx)(i)
                                subst.FugacityCoeff = FCL(i)
                                subst.ActivityCoeff = 0
                                subst.PartialVolume = 0
                                subst.PartialPressure = 0
                                i += 1
                            Next
                            i = 0
                            For Each subst In Me.CurrentMaterialStream.Fases(4).Componentes.Values
                                subst.FracaoMolar = Vx2(i)
                                subst.FracaoMassica = Me.AUX_CONVERT_MOL_TO_MASS(Vx2)(i)
                                subst.FugacityCoeff = FCL2(i)
                                subst.ActivityCoeff = 0
                                subst.PartialVolume = 0
                                subst.PartialPressure = 0
                                i += 1
                            Next
                            i = 0
                            For Each subst In Me.CurrentMaterialStream.Fases(2).Componentes.Values
                                subst.FracaoMolar = Vy(i)
                                subst.FracaoMassica = Me.AUX_CONVERT_MOL_TO_MASS(Vy)(i)
                                subst.FugacityCoeff = FCV(i)
                                subst.ActivityCoeff = 0
                                subst.PartialVolume = 0
                                subst.PartialPressure = 0
                                i += 1
                            Next
                            i = 0
                            For Each subst In Me.CurrentMaterialStream.Fases(7).Componentes.Values
                                subst.FracaoMolar = Vs(i)
                                subst.FracaoMassica = Me.AUX_CONVERT_MOL_TO_MASS(Vs)(i)
                                subst.FugacityCoeff = FCS(i)
                                subst.ActivityCoeff = 0
                                subst.PartialVolume = 0
                                subst.PartialPressure = 0
                                i += 1
                            Next

                            Me.CurrentMaterialStream.Fases(3).SPMProperties.massfraction = xl * Me.AUX_MMM(Fase.Liquid1) / (xl * Me.AUX_MMM(Fase.Liquid1) + xl2 * Me.AUX_MMM(Fase.Liquid2) + xv * Me.AUX_MMM(Fase.Vapor) + xs * Me.AUX_MMM(Fase.Solid))
                            Me.CurrentMaterialStream.Fases(4).SPMProperties.massfraction = xl2 * Me.AUX_MMM(Fase.Liquid2) / (xl * Me.AUX_MMM(Fase.Liquid1) + xl2 * Me.AUX_MMM(Fase.Liquid2) + xv * Me.AUX_MMM(Fase.Vapor) + xs * Me.AUX_MMM(Fase.Solid))
                            Me.CurrentMaterialStream.Fases(2).SPMProperties.massfraction = xv * Me.AUX_MMM(Fase.Vapor) / (xl * Me.AUX_MMM(Fase.Liquid1) + xl2 * Me.AUX_MMM(Fase.Liquid2) + xv * Me.AUX_MMM(Fase.Vapor) + xs * Me.AUX_MMM(Fase.Solid))
                            Me.CurrentMaterialStream.Fases(7).SPMProperties.massfraction = xs * Me.AUX_MMM(Fase.Solid) / (xl * Me.AUX_MMM(Fase.Liquid1) + xl2 * Me.AUX_MMM(Fase.Liquid2) + xv * Me.AUX_MMM(Fase.Vapor) + xs * Me.AUX_MMM(Fase.Solid))

                            Dim constprops As New List(Of ConstantProperties)
                            For Each su As Substancia In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                                constprops.Add(su.ConstantProperties)
                            Next

                            Dim HM, HV, HL, HL2, HS As Double

                            If xl <> 0 Then HL = Me.DW_CalcEnthalpy(Vx, T, P, State.Liquid)
                            If xl2 <> 0 Then HL2 = Me.DW_CalcEnthalpy(Vx2, T, P, State.Liquid)
                            If xv <> 0 Then HV = Me.DW_CalcEnthalpy(Vy, T, P, State.Vapor)
                            If xs <> 0 Then HS = Me.DW_CalcSolidEnthalpy(T, Vs, constprops)
                            HM = Me.CurrentMaterialStream.Fases(4).SPMProperties.massfraction.GetValueOrDefault * HL2 + Me.CurrentMaterialStream.Fases(3).SPMProperties.massfraction.GetValueOrDefault * HL + Me.CurrentMaterialStream.Fases(2).SPMProperties.massfraction.GetValueOrDefault * HV + Me.CurrentMaterialStream.Fases(7).SPMProperties.massfraction.GetValueOrDefault * HS

                            H = HM

                            Dim SM, SV, SL, SL2, SS As Double

                            If xl <> 0 Then SL = Me.DW_CalcEntropy(Vx, T, P, State.Liquid)
                            If xl2 <> 0 Then SL2 = Me.DW_CalcEntropy(Vx2, T, P, State.Liquid)
                            If xv <> 0 Then SV = Me.DW_CalcEntropy(Vy, T, P, State.Vapor)
                            If xs <> 0 Then SS = Me.DW_CalcSolidEnthalpy(T, Vs, constprops) / (T - 298.15)
                            SM = Me.CurrentMaterialStream.Fases(4).SPMProperties.massfraction.GetValueOrDefault * SL2 + Me.CurrentMaterialStream.Fases(3).SPMProperties.massfraction.GetValueOrDefault * SL + Me.CurrentMaterialStream.Fases(2).SPMProperties.massfraction.GetValueOrDefault * SV + Me.CurrentMaterialStream.Fases(7).SPMProperties.massfraction.GetValueOrDefault * SS

                            S = SM

                        Case FlashSpec.H

                            Throw New Exception(DTL.App.GetLocalString("PropPack_FlashTHNotSupported"))

                        Case FlashSpec.S

                            Throw New Exception(DTL.App.GetLocalString("PropPack_FlashTSNotSupported"))

                        Case FlashSpec.VAP

                            Dim KI(n) As Double
                            Dim HM, HV, HL, HL2 As Double
                            Dim SM, SV, SL, SL2 As Double

                            i = 0
                            Do
                                KI(i) = 0
                                i = i + 1
                            Loop Until i = n + 1

                            T = Me.CurrentMaterialStream.Fases(0).SPMProperties.temperature.GetValueOrDefault
                            P = Me.CurrentMaterialStream.Fases(0).SPMProperties.pressure.GetValueOrDefault

                            If Double.IsNaN(P) Or Double.IsInfinity(P) Then P = 0.0#

                            Dim Vx, Vx2, Vy As Double()

                            If Me.AUX_IS_SINGLECOMP(Fase.Mixture) Then

                                Dim Psat As Double
                                Dim vz As Object = Me.RET_VMOL(Fase.Mixture)

                                Psat = Me.AUX_PVAPM(T)

                                HL = Me.DW_CalcEnthalpy(vz, T, Psat, State.Liquid)
                                HV = Me.DW_CalcEnthalpy(vz, T, Psat, State.Vapor)
                                SL = Me.DW_CalcEntropy(vz, T, Psat, State.Liquid)
                                SV = Me.DW_CalcEntropy(vz, T, Psat, State.Vapor)
                                H = xv * HV + (1 - xv) * HL
                                S = xv * SV + (1 - xv) * SL
                                P = Psat
                                xl = 1 - xv
                                xl2 = 0.0#

                                Vx = vz
                                Vy = vz
                                Vx2 = vz

                            Else

                                result = Me.FlashBase.Flash_TV(RET_VMOL(Fase.Mixture), T, xv, P, Me)

                                P = result(4)

                                xl = result(0)
                                xv = result(1)
                                xl2 = result(7)

                                Vx = result(2)
                                Vy = result(3)
                                Vx2 = result(8)

                            End If

                            Me.CurrentMaterialStream.Fases(3).SPMProperties.molarfraction = xl
                            Me.CurrentMaterialStream.Fases(4).SPMProperties.molarfraction = xl2
                            Me.CurrentMaterialStream.Fases(2).SPMProperties.molarfraction = xv

                            Dim FCL = Me.DW_CalcFugCoeff(Vx, T, P, State.Liquid)
                            Dim FCL2 = Me.DW_CalcFugCoeff(Vx2, T, P, State.Liquid)
                            Dim FCV = Me.DW_CalcFugCoeff(Vy, T, P, State.Vapor)

                            i = 0
                            For Each subst In Me.CurrentMaterialStream.Fases(3).Componentes.Values
                                subst.FracaoMolar = Vx(i)
                                subst.FracaoMassica = Me.AUX_CONVERT_MOL_TO_MASS(Vx)(i)
                                subst.FugacityCoeff = FCL(i)
                                subst.ActivityCoeff = 0
                                subst.PartialVolume = 0
                                subst.PartialPressure = 0
                                i += 1
                            Next
                            i = 0
                            For Each subst In Me.CurrentMaterialStream.Fases(4).Componentes.Values
                                subst.FracaoMolar = Vx2(i)
                                subst.FracaoMassica = Me.AUX_CONVERT_MOL_TO_MASS(Vx2)(i)
                                subst.FugacityCoeff = FCL2(i)
                                subst.ActivityCoeff = 0
                                subst.PartialVolume = 0
                                subst.PartialPressure = 0
                                i += 1
                            Next
                            i = 0
                            For Each subst In Me.CurrentMaterialStream.Fases(2).Componentes.Values
                                subst.FracaoMolar = Vy(i)
                                subst.FracaoMassica = Me.AUX_CONVERT_MOL_TO_MASS(Vy)(i)
                                subst.FugacityCoeff = FCV(i)
                                subst.ActivityCoeff = 0
                                subst.PartialVolume = 0
                                subst.PartialPressure = 0
                                i += 1
                            Next

                            Me.CurrentMaterialStream.Fases(3).SPMProperties.massfraction = xl * Me.AUX_MMM(Fase.Liquid1) / (xl * Me.AUX_MMM(Fase.Liquid1) + xl2 * Me.AUX_MMM(Fase.Liquid2) + xv * Me.AUX_MMM(Fase.Vapor))
                            Me.CurrentMaterialStream.Fases(4).SPMProperties.massfraction = xl2 * Me.AUX_MMM(Fase.Liquid2) / (xl * Me.AUX_MMM(Fase.Liquid1) + xl2 * Me.AUX_MMM(Fase.Liquid2) + xv * Me.AUX_MMM(Fase.Vapor))
                            Me.CurrentMaterialStream.Fases(2).SPMProperties.massfraction = xv * Me.AUX_MMM(Fase.Vapor) / (xl * Me.AUX_MMM(Fase.Liquid1) + xl2 * Me.AUX_MMM(Fase.Liquid2) + xv * Me.AUX_MMM(Fase.Vapor))

                            If xl <> 0 Then HL = Me.DW_CalcEnthalpy(Vx, T, P, State.Liquid)
                            If xl2 <> 0 Then HL2 = Me.DW_CalcEnthalpy(Vx2, T, P, State.Liquid)
                            If xv <> 0 Then HV = Me.DW_CalcEnthalpy(Vy, T, P, State.Vapor)
                            HM = Me.CurrentMaterialStream.Fases(4).SPMProperties.massfraction.GetValueOrDefault * HL2 + Me.CurrentMaterialStream.Fases(3).SPMProperties.massfraction.GetValueOrDefault * HL + Me.CurrentMaterialStream.Fases(2).SPMProperties.massfraction.GetValueOrDefault * HV

                            H = HM


                            If xl <> 0 Then SL = Me.DW_CalcEntropy(Vx, T, P, State.Liquid)
                            If xl2 <> 0 Then SL2 = Me.DW_CalcEntropy(Vx2, T, P, State.Liquid)
                            If xv <> 0 Then SV = Me.DW_CalcEntropy(Vy, T, P, State.Vapor)
                            SM = Me.CurrentMaterialStream.Fases(4).SPMProperties.massfraction.GetValueOrDefault * SL2 + Me.CurrentMaterialStream.Fases(3).SPMProperties.massfraction.GetValueOrDefault * SL + Me.CurrentMaterialStream.Fases(2).SPMProperties.massfraction.GetValueOrDefault * SV

                            S = SM

                    End Select

                Case FlashSpec.P

                    Select Case spec2

                        Case FlashSpec.H

                            T = Me.CurrentMaterialStream.Fases(0).SPMProperties.temperature.GetValueOrDefault
                            P = Me.CurrentMaterialStream.Fases(0).SPMProperties.pressure.GetValueOrDefault

                            If Double.IsNaN(H) Or Double.IsInfinity(H) Then H = Me.CurrentMaterialStream.Fases(0).SPMProperties.molar_enthalpy.GetValueOrDefault / Me.CurrentMaterialStream.Fases(0).SPMProperties.molecularWeight.GetValueOrDefault
                            If Double.IsNaN(T) Or Double.IsInfinity(T) Then T = 0.0#

                            If Me.AUX_IS_SINGLECOMP(Fase.Mixture) And Me.ComponentName <> "FPROPS" Then

                                Dim brentsolverT As New BrentOpt.Brent
                                brentsolverT.DefineFuncDelegate(AddressOf EnthalpyTx)

                                Dim hl, hv, sl, sv, Tsat As Double
                                Dim vz As Object = Me.RET_VMOL(Fase.Mixture)

                                P = Me.CurrentMaterialStream.Fases(0).SPMProperties.pressure.GetValueOrDefault

                                Tsat = 0.0#
                                For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                                    Tsat += subst.FracaoMolar * Me.AUX_TSATi(P, subst.Nome)
                                Next

                                hl = Me.DW_CalcEnthalpy(vz, Tsat, P, State.Liquid)
                                hv = Me.DW_CalcEnthalpy(vz, Tsat, P, State.Vapor)
                                sl = Me.DW_CalcEntropy(vz, Tsat, P, State.Liquid)
                                sv = Me.DW_CalcEntropy(vz, Tsat, P, State.Vapor)
                                If H <= hl Then
                                    xv = 0
                                    LoopVarState = State.Liquid
                                ElseIf H >= hv Then
                                    xv = 1
                                    LoopVarState = State.Vapor
                                Else
                                    xv = (H - hl) / (hv - hl)
                                End If
                                If Tsat > Me.AUX_TCM(Fase.Mixture) Then
                                    xv = 1.0#
                                    LoopVarState = State.Vapor
                                End If
                                xl = 1 - xv

                                If xv <> 0.0# And xv <> 1.0# Then
                                    T = Tsat
                                    S = xv * sv + (1 - xv) * sl
                                Else
                                    LoopVarF = H
                                    LoopVarX = P
                                    T = brentsolverT.BrentOpt(Me.AUX_TFM(Fase.Mixture), 2000, 20, 0.0001, 1000, Nothing)
                                    If xv = 0.0# Then
                                        S = Me.DW_CalcEntropy(vz, T, P, State.Liquid)
                                    Else
                                        S = Me.DW_CalcEntropy(vz, T, P, State.Vapor)
                                    End If
                                End If

                                If T <= Me.AUX_TFM(Fase.Mixture) Then

                                    'solid only.

                                    xv = 0.0#
                                    xl = 0.0#
                                    xs = 1.0#

                                    Dim constprops As New List(Of ConstantProperties)
                                    For Each su As Substancia In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                                        constprops.Add(su.ConstantProperties)
                                    Next

                                    S = Me.DW_CalcSolidEnthalpy(T, vz, constprops) / (T - 298.15)

                                    Me.CurrentMaterialStream.Fases(3).SPMProperties.molarfraction = xl
                                    Me.CurrentMaterialStream.Fases(2).SPMProperties.molarfraction = xv
                                    Me.CurrentMaterialStream.Fases(3).SPMProperties.massfraction = xl
                                    Me.CurrentMaterialStream.Fases(2).SPMProperties.massfraction = xv
                                    Me.CurrentMaterialStream.Fases(7).SPMProperties.massfraction = xs
                                    Me.CurrentMaterialStream.Fases(7).SPMProperties.molarfraction = xs
                                    Me.CurrentMaterialStream.Fases(4).SPMProperties.massfraction = 0.0#
                                    Me.CurrentMaterialStream.Fases(4).SPMProperties.molarfraction = 0.0#

                                    i = 0
                                    For Each subst In Me.CurrentMaterialStream.Fases(7).Componentes.Values
                                        subst.FracaoMolar = vz(i)
                                        subst.FugacityCoeff = 1
                                        subst.ActivityCoeff = 1
                                        subst.PartialVolume = 0
                                        subst.PartialPressure = P
                                        subst.FracaoMassica = vz(i)
                                        i += 1
                                    Next

                                Else

                                    Me.CurrentMaterialStream.Fases(3).SPMProperties.molarfraction = xl
                                    Me.CurrentMaterialStream.Fases(2).SPMProperties.molarfraction = xv
                                    Me.CurrentMaterialStream.Fases(3).SPMProperties.massfraction = xl
                                    Me.CurrentMaterialStream.Fases(2).SPMProperties.massfraction = xv

                                    i = 0
                                    For Each subst In Me.CurrentMaterialStream.Fases(3).Componentes.Values
                                        subst.FracaoMolar = vz(i)
                                        subst.FugacityCoeff = 1
                                        subst.ActivityCoeff = 1
                                        subst.PartialVolume = 0
                                        subst.PartialPressure = P
                                        subst.FracaoMassica = vz(i)
                                        i += 1
                                    Next
                                    i = 0
                                    For Each subst In Me.CurrentMaterialStream.Fases(2).Componentes.Values
                                        subst.FracaoMolar = vz(i)
                                        subst.FugacityCoeff = 1
                                        subst.ActivityCoeff = 1
                                        subst.PartialVolume = 0
                                        subst.PartialPressure = P
                                        subst.FracaoMassica = vz(i)
                                        i += 1
                                    Next

                                End If

                            Else

redirect:                       result = Me.FlashBase.Flash_PH(RET_VMOL(Fase.Mixture), P, H, T, Me)

                                T = result(4)

                                xl = result(0)
                                xv = result(1)
                                xl2 = result(7)
                                xs = result(9)

                                Me.CurrentMaterialStream.Fases(3).SPMProperties.molarfraction = xl
                                Me.CurrentMaterialStream.Fases(4).SPMProperties.molarfraction = xl2
                                Me.CurrentMaterialStream.Fases(2).SPMProperties.molarfraction = xv
                                Me.CurrentMaterialStream.Fases(7).SPMProperties.molarfraction = xs

                                Dim Vx = result(2)
                                Dim Vy = result(3)
                                Dim Vx2 = result(8)
                                Dim Vs = result(10)

                                Dim FCL = Me.DW_CalcFugCoeff(Vx, T, P, State.Liquid)
                                Dim FCL2 = Me.DW_CalcFugCoeff(Vx2, T, P, State.Liquid)
                                Dim FCV = Me.DW_CalcFugCoeff(Vy, T, P, State.Vapor)
                                Dim FCS = Me.DW_CalcFugCoeff(Vs, T, P, State.Solid)

                                i = 0
                                For Each subst In Me.CurrentMaterialStream.Fases(3).Componentes.Values
                                    subst.FracaoMolar = Vx(i)
                                    subst.FracaoMassica = Me.AUX_CONVERT_MOL_TO_MASS(Vx)(i)
                                    subst.FugacityCoeff = FCL(i)
                                    subst.ActivityCoeff = 0
                                    subst.PartialVolume = 0
                                    subst.PartialPressure = 0
                                    i += 1
                                Next
                                i = 1
                                i = 0
                                For Each subst In Me.CurrentMaterialStream.Fases(4).Componentes.Values
                                    subst.FracaoMolar = Vx2(i)
                                    subst.FracaoMassica = Me.AUX_CONVERT_MOL_TO_MASS(Vx2)(i)
                                    subst.FugacityCoeff = FCL2(i)
                                    subst.ActivityCoeff = 0
                                    subst.PartialVolume = 0
                                    subst.PartialPressure = 0
                                    i += 1
                                Next
                                i = 0
                                For Each subst In Me.CurrentMaterialStream.Fases(2).Componentes.Values
                                    subst.FracaoMolar = Vy(i)
                                    subst.FracaoMassica = Me.AUX_CONVERT_MOL_TO_MASS(Vy)(i)
                                    subst.FugacityCoeff = FCV(i)
                                    subst.ActivityCoeff = 0
                                    subst.PartialVolume = 0
                                    subst.PartialPressure = 0
                                    i += 1
                                Next
                                i = 0
                                For Each subst In Me.CurrentMaterialStream.Fases(7).Componentes.Values
                                    subst.FracaoMolar = Vs(i)
                                    subst.FracaoMassica = Me.AUX_CONVERT_MOL_TO_MASS(Vs)(i)
                                    subst.FugacityCoeff = FCS(i)
                                    subst.ActivityCoeff = 0
                                    subst.PartialVolume = 0
                                    subst.PartialPressure = 0
                                    i += 1
                                Next

                                Me.CurrentMaterialStream.Fases(3).SPMProperties.massfraction = xl * Me.AUX_MMM(Fase.Liquid1) / (xl * Me.AUX_MMM(Fase.Liquid1) + xl2 * Me.AUX_MMM(Fase.Liquid2) + xv * Me.AUX_MMM(Fase.Vapor) + xs * Me.AUX_MMM(Fase.Solid))
                                Me.CurrentMaterialStream.Fases(4).SPMProperties.massfraction = xl2 * Me.AUX_MMM(Fase.Liquid2) / (xl * Me.AUX_MMM(Fase.Liquid1) + xl2 * Me.AUX_MMM(Fase.Liquid2) + xv * Me.AUX_MMM(Fase.Vapor) + xs * Me.AUX_MMM(Fase.Solid))
                                Me.CurrentMaterialStream.Fases(2).SPMProperties.massfraction = xv * Me.AUX_MMM(Fase.Vapor) / (xl * Me.AUX_MMM(Fase.Liquid1) + xl2 * Me.AUX_MMM(Fase.Liquid2) + xv * Me.AUX_MMM(Fase.Vapor) + xs * Me.AUX_MMM(Fase.Solid))
                                Me.CurrentMaterialStream.Fases(7).SPMProperties.massfraction = xs * Me.AUX_MMM(Fase.Solid) / (xl * Me.AUX_MMM(Fase.Liquid1) + xl2 * Me.AUX_MMM(Fase.Liquid2) + xv * Me.AUX_MMM(Fase.Vapor) + xs * Me.AUX_MMM(Fase.Solid))

                                Dim constprops As New List(Of ConstantProperties)
                                For Each su As Substancia In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                                    constprops.Add(su.ConstantProperties)
                                Next

                                Dim SM, SV, SL, SL2, SS As Double

                                If xl <> 0 Then SL = Me.DW_CalcEntropy(Vx, T, P, State.Liquid)
                                If xl2 <> 0 Then SL2 = Me.DW_CalcEntropy(Vx2, T, P, State.Liquid)
                                If xv <> 0 Then SV = Me.DW_CalcEntropy(Vy, T, P, State.Vapor)
                                If xs <> 0 Then SS = Me.DW_CalcSolidEnthalpy(T, Vs, constprops) / (T - 298.15)
                                SM = Me.CurrentMaterialStream.Fases(4).SPMProperties.massfraction.GetValueOrDefault * SL2 + Me.CurrentMaterialStream.Fases(3).SPMProperties.massfraction.GetValueOrDefault * SL + Me.CurrentMaterialStream.Fases(2).SPMProperties.massfraction.GetValueOrDefault * SV + Me.CurrentMaterialStream.Fases(7).SPMProperties.massfraction.GetValueOrDefault * SS

                                S = SM

                            End If

                        Case FlashSpec.S

                            T = Me.CurrentMaterialStream.Fases(0).SPMProperties.temperature.GetValueOrDefault
                            P = Me.CurrentMaterialStream.Fases(0).SPMProperties.pressure.GetValueOrDefault

                            If Double.IsNaN(S) Or Double.IsInfinity(S) Then S = Me.CurrentMaterialStream.Fases(0).SPMProperties.molar_entropy.GetValueOrDefault / Me.CurrentMaterialStream.Fases(0).SPMProperties.molecularWeight.GetValueOrDefault

                            If Me.AUX_IS_SINGLECOMP(Fase.Mixture) And Me.ComponentName <> "FPROPS" And Me.ComponentName <> "CoolProp" Then

                                Dim brentsolverT As New BrentOpt.Brent
                                brentsolverT.DefineFuncDelegate(AddressOf EntropyTx)

                                Dim hl, hv, sl, sv, Tsat As Double
                                Dim vz As Object = Me.RET_VMOL(Fase.Mixture)

                                P = Me.CurrentMaterialStream.Fases(0).SPMProperties.pressure.GetValueOrDefault

                                Tsat = 0.0#
                                For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                                    Tsat += subst.FracaoMolar * Me.AUX_TSATi(P, subst.Nome)
                                Next

                                hl = Me.DW_CalcEnthalpy(vz, Tsat, P, State.Liquid)
                                hv = Me.DW_CalcEnthalpy(vz, Tsat, P, State.Vapor)
                                sl = Me.DW_CalcEntropy(vz, Tsat, P, State.Liquid)
                                sv = Me.DW_CalcEntropy(vz, Tsat, P, State.Vapor)
                                If S <= sl Then
                                    xv = 0
                                    LoopVarState = State.Liquid
                                ElseIf S >= sv Then
                                    xv = 1
                                    LoopVarState = State.Vapor
                                Else
                                    xv = (S - sl) / (sv - sl)
                                End If
                                If Tsat > Me.AUX_TCM(Fase.Mixture) Then
                                    xv = 1.0#
                                    LoopVarState = State.Vapor
                                End If
                                xl = 1 - xv

                                If xv <> 0.0# And xv <> 1.0# Then
                                    T = Tsat
                                    H = xv * hv + (1 - xv) * hl
                                Else
                                    LoopVarF = S
                                    LoopVarX = P
                                    T = brentsolverT.BrentOpt(Me.AUX_TFM(Fase.Mixture), 2000, 20, 0.0001, 1000, Nothing)
                                    If xv = 0.0# Then
                                        H = Me.DW_CalcEnthalpy(vz, T, P, State.Liquid)
                                    Else
                                        H = Me.DW_CalcEnthalpy(vz, T, P, State.Vapor)
                                    End If
                                End If

                                Me.CurrentMaterialStream.Fases(3).SPMProperties.molarfraction = xl
                                Me.CurrentMaterialStream.Fases(2).SPMProperties.molarfraction = xv
                                Me.CurrentMaterialStream.Fases(3).SPMProperties.massfraction = xl
                                Me.CurrentMaterialStream.Fases(2).SPMProperties.massfraction = xv

                                i = 0
                                For Each subst In Me.CurrentMaterialStream.Fases(3).Componentes.Values
                                    subst.FracaoMolar = vz(i)
                                    subst.FugacityCoeff = 1
                                    subst.ActivityCoeff = 1
                                    subst.PartialVolume = 0
                                    subst.PartialPressure = P
                                    subst.FracaoMassica = vz(i)
                                    i += 1
                                Next
                                i = 0
                                For Each subst In Me.CurrentMaterialStream.Fases(2).Componentes.Values
                                    subst.FracaoMolar = vz(i)
                                    subst.FugacityCoeff = 1
                                    subst.ActivityCoeff = 1
                                    subst.PartialVolume = 0
                                    subst.PartialPressure = P
                                    subst.FracaoMassica = vz(i)
                                    i += 1
                                Next

                            Else

redirect2:                      result = Me.FlashBase.Flash_PS(RET_VMOL(Fase.Mixture), P, S, T, Me)

                                T = result(4)

                                xl = result(0)
                                xv = result(1)
                                xl2 = result(7)

                                Me.CurrentMaterialStream.Fases(3).SPMProperties.molarfraction = xl
                                Me.CurrentMaterialStream.Fases(4).SPMProperties.molarfraction = xl2
                                Me.CurrentMaterialStream.Fases(2).SPMProperties.molarfraction = xv

                                Dim Vx = result(2)
                                Dim Vy = result(3)
                                Dim Vx2 = result(8)

                                Dim FCL = Me.DW_CalcFugCoeff(Vx, T, P, State.Liquid)
                                Dim FCL2 = Me.DW_CalcFugCoeff(Vx2, T, P, State.Liquid)
                                Dim FCV = Me.DW_CalcFugCoeff(Vy, T, P, State.Vapor)

                                i = 0
                                For Each subst In Me.CurrentMaterialStream.Fases(3).Componentes.Values
                                    subst.FracaoMolar = Vx(i)
                                    subst.FugacityCoeff = FCL(i)
                                    subst.ActivityCoeff = 0
                                    subst.PartialVolume = 0
                                    subst.PartialPressure = 0
                                    i += 1
                                Next
                                i = 1
                                For Each subst In Me.CurrentMaterialStream.Fases(3).Componentes.Values
                                    subst.FracaoMassica = Me.AUX_CONVERT_MOL_TO_MASS(subst.Nome, 3)
                                    i += 1
                                Next
                                i = 0
                                For Each subst In Me.CurrentMaterialStream.Fases(4).Componentes.Values
                                    subst.FracaoMolar = Vx2(i)
                                    subst.FugacityCoeff = FCL2(i)
                                    subst.ActivityCoeff = 0
                                    subst.PartialVolume = 0
                                    subst.PartialPressure = 0
                                    i += 1
                                Next
                                i = 1
                                For Each subst In Me.CurrentMaterialStream.Fases(4).Componentes.Values
                                    subst.FracaoMassica = Me.AUX_CONVERT_MOL_TO_MASS(subst.Nome, 4)
                                    i += 1
                                Next
                                i = 0
                                For Each subst In Me.CurrentMaterialStream.Fases(2).Componentes.Values
                                    subst.FracaoMolar = Vy(i)
                                    subst.FugacityCoeff = FCV(i)
                                    subst.ActivityCoeff = 0
                                    subst.PartialVolume = 0
                                    subst.PartialPressure = 0
                                    i += 1
                                Next

                                i = 1
                                For Each subst In Me.CurrentMaterialStream.Fases(2).Componentes.Values
                                    subst.FracaoMassica = Me.AUX_CONVERT_MOL_TO_MASS(subst.Nome, 2)
                                    i += 1
                                Next

                                Me.CurrentMaterialStream.Fases(3).SPMProperties.massfraction = xl * Me.AUX_MMM(Fase.Liquid1) / (xl * Me.AUX_MMM(Fase.Liquid1) + xl2 * Me.AUX_MMM(Fase.Liquid2) + xv * Me.AUX_MMM(Fase.Vapor))
                                Me.CurrentMaterialStream.Fases(4).SPMProperties.massfraction = xl2 * Me.AUX_MMM(Fase.Liquid2) / (xl * Me.AUX_MMM(Fase.Liquid1) + xl2 * Me.AUX_MMM(Fase.Liquid2) + xv * Me.AUX_MMM(Fase.Vapor))
                                Me.CurrentMaterialStream.Fases(2).SPMProperties.massfraction = xv * Me.AUX_MMM(Fase.Vapor) / (xl * Me.AUX_MMM(Fase.Liquid1) + xl2 * Me.AUX_MMM(Fase.Liquid2) + xv * Me.AUX_MMM(Fase.Vapor))

                                Dim HM, HV, HL, HL2 As Double

                                If xl <> 0 Then HL = Me.DW_CalcEnthalpy(Vx, T, P, State.Liquid)
                                If xl2 <> 0 Then HL2 = Me.DW_CalcEnthalpy(Vx2, T, P, State.Liquid)
                                If xv <> 0 Then HV = Me.DW_CalcEnthalpy(Vy, T, P, State.Vapor)
                                HM = Me.CurrentMaterialStream.Fases(4).SPMProperties.massfraction.GetValueOrDefault * HL2 + Me.CurrentMaterialStream.Fases(3).SPMProperties.massfraction.GetValueOrDefault * HL + Me.CurrentMaterialStream.Fases(2).SPMProperties.massfraction.GetValueOrDefault * HV

                                H = HM

                            End If

                        Case FlashSpec.VAP

                            Dim KI(n) As Double
                            Dim HM, HV, HL, HL2 As Double
                            Dim SM, SV, SL, SL2 As Double

                            i = 0
                            Do
                                KI(i) = 0
                                i = i + 1
                            Loop Until i = n + 1

                            T = Me.CurrentMaterialStream.Fases(0).SPMProperties.temperature.GetValueOrDefault
                            P = Me.CurrentMaterialStream.Fases(0).SPMProperties.pressure.GetValueOrDefault

                            If Double.IsNaN(T) Or Double.IsInfinity(T) Then T = 0.0#

                            Dim Vx, Vx2, Vy As Double()

                            If Me.AUX_IS_SINGLECOMP(Fase.Mixture) Then

                                Dim Tsat As Double
                                Dim vz As Object = Me.RET_VMOL(Fase.Mixture)

                                Tsat = 0.0#
                                For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                                    Tsat += subst.FracaoMolar * Me.AUX_TSATi(P, subst.Nome)
                                Next

                                HL = Me.DW_CalcEnthalpy(vz, Tsat, P, State.Liquid)
                                HV = Me.DW_CalcEnthalpy(vz, Tsat, P, State.Vapor)
                                SL = Me.DW_CalcEntropy(vz, Tsat, P, State.Liquid)
                                SV = Me.DW_CalcEntropy(vz, Tsat, P, State.Vapor)
                                H = xv * HV + (1 - xv) * HL
                                S = xv * SV + (1 - xv) * SL
                                T = Tsat
                                xl = 1 - xv
                                xl2 = 0.0#

                                Vx = vz
                                Vy = vz
                                Vx2 = vz

                            Else

                                result = Me.FlashBase.Flash_PV(RET_VMOL(Fase.Mixture), P, xv, T, Me)

                                T = result(4)

                                xl = result(0)
                                xv = result(1)
                                xl2 = result(7)

                                Vx = result(2)
                                Vy = result(3)
                                Vx2 = result(8)

                            End If

                            Me.CurrentMaterialStream.Fases(3).SPMProperties.molarfraction = xl
                            Me.CurrentMaterialStream.Fases(4).SPMProperties.molarfraction = xl2
                            Me.CurrentMaterialStream.Fases(2).SPMProperties.molarfraction = xv

                            Dim FCL = Me.DW_CalcFugCoeff(Vx, T, P, State.Liquid)
                            Dim FCL2 = Me.DW_CalcFugCoeff(Vx2, T, P, State.Liquid)
                            Dim FCV = Me.DW_CalcFugCoeff(Vy, T, P, State.Vapor)

                            i = 0
                            For Each subst In Me.CurrentMaterialStream.Fases(3).Componentes.Values
                                subst.FracaoMolar = Vx(i)
                                subst.FugacityCoeff = FCL(i)
                                subst.ActivityCoeff = 0
                                subst.PartialVolume = 0
                                subst.PartialPressure = 0
                                i += 1
                            Next
                            For Each subst In Me.CurrentMaterialStream.Fases(3).Componentes.Values
                                subst.FracaoMassica = Me.AUX_CONVERT_MOL_TO_MASS(subst.Nome, 3)
                            Next
                            i = 0
                            For Each subst In Me.CurrentMaterialStream.Fases(4).Componentes.Values
                                subst.FracaoMolar = Vx2(i)
                                subst.FugacityCoeff = FCL2(i)
                                subst.ActivityCoeff = 0
                                subst.PartialVolume = 0
                                subst.PartialPressure = 0
                                i += 1
                            Next
                            For Each subst In Me.CurrentMaterialStream.Fases(4).Componentes.Values
                                subst.FracaoMassica = Me.AUX_CONVERT_MOL_TO_MASS(subst.Nome, 4)
                            Next
                            i = 0
                            For Each subst In Me.CurrentMaterialStream.Fases(2).Componentes.Values
                                subst.FracaoMolar = Vy(i)
                                subst.FugacityCoeff = FCV(i)
                                subst.ActivityCoeff = 0
                                subst.PartialVolume = 0
                                subst.PartialPressure = 0
                                i += 1
                            Next
                            For Each subst In Me.CurrentMaterialStream.Fases(2).Componentes.Values
                                subst.FracaoMassica = Me.AUX_CONVERT_MOL_TO_MASS(subst.Nome, 2)
                            Next

                            Me.CurrentMaterialStream.Fases(3).SPMProperties.massfraction = xl * Me.AUX_MMM(Fase.Liquid1) / (xl * Me.AUX_MMM(Fase.Liquid1) + xl2 * Me.AUX_MMM(Fase.Liquid2) + xv * Me.AUX_MMM(Fase.Vapor))
                            Me.CurrentMaterialStream.Fases(4).SPMProperties.massfraction = xl2 * Me.AUX_MMM(Fase.Liquid2) / (xl * Me.AUX_MMM(Fase.Liquid1) + xl2 * Me.AUX_MMM(Fase.Liquid2) + xv * Me.AUX_MMM(Fase.Vapor))
                            Me.CurrentMaterialStream.Fases(2).SPMProperties.massfraction = xv * Me.AUX_MMM(Fase.Vapor) / (xl * Me.AUX_MMM(Fase.Liquid1) + xl2 * Me.AUX_MMM(Fase.Liquid2) + xv * Me.AUX_MMM(Fase.Vapor))

                            If xl <> 0 Then HL = Me.DW_CalcEnthalpy(Vx, T, P, State.Liquid)
                            If xl2 <> 0 Then HL2 = Me.DW_CalcEnthalpy(Vx2, T, P, State.Liquid)
                            If xv <> 0 Then HV = Me.DW_CalcEnthalpy(Vy, T, P, State.Vapor)
                            HM = Me.CurrentMaterialStream.Fases(4).SPMProperties.massfraction.GetValueOrDefault * HL2 + Me.CurrentMaterialStream.Fases(3).SPMProperties.massfraction.GetValueOrDefault * HL + Me.CurrentMaterialStream.Fases(2).SPMProperties.massfraction.GetValueOrDefault * HV

                            H = HM

                            If xl <> 0 Then SL = Me.DW_CalcEntropy(Vx, T, P, State.Liquid)
                            If xl2 <> 0 Then SL2 = Me.DW_CalcEntropy(Vx2, T, P, State.Liquid)
                            If xv <> 0 Then SV = Me.DW_CalcEntropy(Vy, T, P, State.Vapor)
                            SM = Me.CurrentMaterialStream.Fases(4).SPMProperties.massfraction.GetValueOrDefault * SL2 + Me.CurrentMaterialStream.Fases(3).SPMProperties.massfraction.GetValueOrDefault * SL + Me.CurrentMaterialStream.Fases(2).SPMProperties.massfraction.GetValueOrDefault * SV

                            S = SM

                    End Select

            End Select

            Dim summf As Double = 0, sumwf As Double = 0
            For Each pi As PhaseInfo In Me.PhaseMappings.Values
                If Not pi.PhaseLabel = "Disabled" Then
                    summf += Me.CurrentMaterialStream.Fases(pi.DWPhaseIndex).SPMProperties.molarfraction.GetValueOrDefault
                    sumwf += Me.CurrentMaterialStream.Fases(pi.DWPhaseIndex).SPMProperties.massfraction.GetValueOrDefault
                End If
            Next
            If Abs(summf - 1) > 0.0001 Then
                For Each pi As PhaseInfo In Me.PhaseMappings.Values
                    If Not pi.PhaseLabel = "Disabled" Then
                        If Not Me.CurrentMaterialStream.Fases(pi.DWPhaseIndex).SPMProperties.molarfraction.HasValue Then
                            Me.CurrentMaterialStream.Fases(pi.DWPhaseIndex).SPMProperties.molarfraction = 1 - summf
                            Me.CurrentMaterialStream.Fases(pi.DWPhaseIndex).SPMProperties.massfraction = 1 - sumwf
                        End If
                    End If
                Next
            End If

            With Me.CurrentMaterialStream

                .Fases(0).SPMProperties.temperature = T
                .Fases(0).SPMProperties.pressure = P
                .Fases(0).SPMProperties.enthalpy = H
                .Fases(0).SPMProperties.entropy = S

            End With

            Me.CurrentMaterialStream.AtEquilibrium = True

        End Sub

        Private Function EnthalpyTx(ByVal x As Double, ByVal otherargs As Object) As Double

            Dim er As Double = LoopVarF - Me.DW_CalcEnthalpy(Me.RET_VMOL(Fase.Mixture), x, LoopVarX, LoopVarState)
            Return er

        End Function

        Private Function EnthalpyPx(ByVal x As Double, ByVal otherargs As Object) As Double

            Dim er As Double = LoopVarF - Me.DW_CalcEnthalpy(Me.RET_VMOL(Fase.Mixture), LoopVarX, x, LoopVarState)
            Return er

        End Function

        Private Function EntropyTx(ByVal x As Double, ByVal otherargs As Object) As Double

            Dim er As Double = LoopVarF - Me.DW_CalcEntropy(Me.RET_VMOL(Fase.Mixture), x, LoopVarX, LoopVarState)
            Return er

        End Function

        Private Function EntropyPx(ByVal x As Double, ByVal otherargs As Object) As Double

            Dim er As Double = LoopVarF - Me.DW_CalcEntropy(Me.RET_VMOL(Fase.Mixture), LoopVarX, x, LoopVarState)
            Return er

        End Function


        Public MustOverride Sub DW_CalcVazaoMolar()

        Public MustOverride Sub DW_CalcVazaoMassica()

        Public MustOverride Sub DW_CalcVazaoVolumetrica()

        Public MustOverride Function DW_CalcMassaEspecifica_ISOL(ByVal fase1 As Fase, ByVal T As Double, ByVal P As Double, Optional ByVal Pvp As Double = 0) As Double

        Public MustOverride Function DW_CalcViscosidadeDinamica_ISOL(ByVal fase1 As Fase, ByVal T As Double, ByVal P As Double) As Double


        Public MustOverride Function DW_CalcEnergiaMistura_ISOL(ByVal T As Double, ByVal P As Double) As Double

        Public MustOverride Function DW_CalcCp_ISOL(ByVal fase1 As DTL.SimulationObjects.PropertyPackages.Fase, ByVal T As Double, ByVal P As Double) As Double

        Public MustOverride Function DW_CalcCv_ISOL(ByVal fase1 As DTL.SimulationObjects.PropertyPackages.Fase, ByVal T As Double, ByVal P As Double) As Double

        Public MustOverride Function DW_CalcK_ISOL(ByVal fase1 As DTL.SimulationObjects.PropertyPackages.Fase, ByVal T As Double, ByVal P As Double) As Double

        Public MustOverride Function DW_CalcMM_ISOL(ByVal fase1 As DTL.SimulationObjects.PropertyPackages.Fase, ByVal T As Double, ByVal P As Double) As Double

        Public MustOverride Function DW_CalcPVAP_ISOL(ByVal T As Double) As Double

        Public MustOverride Sub DW_CalcCompPartialVolume(ByVal phase As DTL.SimulationObjects.PropertyPackages.Fase, ByVal T As Double, ByVal P As Double)

#End Region

#Region "   Commmon Functions"

        Public Overloads Sub DW_CalcKvalue()

            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia

            For Each phase As DTL.ClassesBasicasTermodinamica.Fase In Me.CurrentMaterialStream.Fases.Values
                For Each subst In phase.Componentes.Values
                    subst.Kvalue = Me.CurrentMaterialStream.Fases(2).Componentes(subst.Nome).FracaoMolar.GetValueOrDefault / phase.Componentes(subst.Nome).FracaoMolar.GetValueOrDefault
                    subst.lnKvalue = Log(subst.Kvalue)
                Next
            Next

        End Sub

        Public Sub DW_CalcCompMolarFlow(ByVal phaseID As Integer)

            If Not phaseID = -1 Then
                With Me.CurrentMaterialStream.Fases(phaseID)
                    For Each subs As Substancia In .Componentes.Values
                        subs.MolarFlow = .SPMProperties.molarflow.GetValueOrDefault * subs.FracaoMolar.GetValueOrDefault
                    Next
                End With
            Else
                For Each phase As DTL.ClassesBasicasTermodinamica.Fase In Me.CurrentMaterialStream.Fases.Values
                    With phase
                        For Each subs As Substancia In .Componentes.Values
                            subs.MolarFlow = .SPMProperties.molarflow.GetValueOrDefault * subs.FracaoMolar.GetValueOrDefault
                        Next
                    End With
                Next
            End If


        End Sub

        Public Sub DW_CalcCompFugCoeff(ByVal f As Fase)

            Dim fc As Object
            Dim vmol As Object = Me.RET_VMOL(f)
            Dim P, T As Double
            P = Me.CurrentMaterialStream.Fases(0).SPMProperties.pressure.GetValueOrDefault
            T = Me.CurrentMaterialStream.Fases(0).SPMProperties.temperature.GetValueOrDefault
            Select Case f
                Case Fase.Vapor
                    fc = Me.DW_CalcFugCoeff(vmol, T, P, State.Vapor)
                Case Else
                    fc = Me.DW_CalcFugCoeff(vmol, T, P, State.Liquid)
            End Select
            Dim i As Integer = 0
            For Each subs As Substancia In Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(f)).Componentes.Values
                subs.FugacityCoeff = fc(i)
                subs.ActivityCoeff = fc(i) * P / Me.AUX_PVAPi(i, T)
                i += 1
            Next

        End Sub

        Public Sub DW_CalcCompMassFlow(ByVal phaseID As Integer)

            If Not phaseID = -1 Then
                With Me.CurrentMaterialStream.Fases(phaseID)
                    For Each subs As Substancia In .Componentes.Values
                        subs.MassFlow = .SPMProperties.massflow.GetValueOrDefault * subs.FracaoMassica.GetValueOrDefault
                    Next
                End With
            Else
                For Each phase As DTL.ClassesBasicasTermodinamica.Fase In Me.CurrentMaterialStream.Fases.Values
                    With phase
                        For Each subs As Substancia In .Componentes.Values
                            subs.MassFlow = .SPMProperties.massflow.GetValueOrDefault * subs.FracaoMassica.GetValueOrDefault
                        Next
                    End With
                Next
            End If

        End Sub

        Public Sub DW_CalcCompVolFlow(ByVal phaseID As Integer)

            Dim TotalMolarFlow, TotalVolFlow, MolarFrac, PartialVol, VolFrac, VolFlow, Sum As Double

            Dim T, P As Double
            T = Me.CurrentMaterialStream.Fases(0).SPMProperties.temperature.GetValueOrDefault
            P = Me.CurrentMaterialStream.Fases(0).SPMProperties.pressure.GetValueOrDefault
            Me.DW_CalcCompPartialVolume(Me.RET_PHASECODE(phaseID), T, P)

            If Not phaseID = -1 Then
                With Me.CurrentMaterialStream.Fases(phaseID)

                    Sum = 0

                    For Each subs As Substancia In .Componentes.Values
                        TotalMolarFlow = .SPMProperties.molarflow.GetValueOrDefault
                        TotalVolFlow = .SPMProperties.volumetric_flow.GetValueOrDefault
                        MolarFrac = subs.FracaoMolar.GetValueOrDefault
                        PartialVol = subs.PartialVolume.GetValueOrDefault
                        VolFlow = TotalMolarFlow * MolarFrac * PartialVol
                        If TotalVolFlow > 0 Then
                            VolFrac = VolFlow / TotalVolFlow
                        Else
                            VolFrac = 0
                        End If
                        subs.VolumetricFraction = VolFrac
                        Sum += VolFrac
                    Next

                    'Normalization is still needed due to minor deviations in the partial molar volume estimation. Summation of partial flow rates
                    'should match phase flow rate.
                    For Each subs As Substancia In .Componentes.Values
                        If Sum > 0 Then
                            subs.VolumetricFraction = subs.VolumetricFraction.GetValueOrDefault / Sum
                        Else
                            subs.VolumetricFraction = 0
                        End If
                        'Corrects volumetric flow rate after normalization of fractions.
                        subs.VolumetricFlow = subs.VolumetricFraction.GetValueOrDefault * TotalVolFlow
                    Next

                End With
            Else
                For Each phase As DTL.ClassesBasicasTermodinamica.Fase In Me.CurrentMaterialStream.Fases.Values
                    With phase

                        Sum = 0

                        For Each subs As Substancia In .Componentes.Values
                            TotalMolarFlow = .SPMProperties.molarflow.GetValueOrDefault
                            TotalVolFlow = .SPMProperties.volumetric_flow.GetValueOrDefault
                            MolarFrac = subs.FracaoMolar.GetValueOrDefault
                            PartialVol = subs.PartialVolume.GetValueOrDefault
                            VolFlow = TotalMolarFlow * MolarFrac * PartialVol
                            If TotalVolFlow > 0 Then
                                VolFrac = VolFlow / TotalVolFlow
                            Else
                                VolFrac = 0
                            End If
                            subs.VolumetricFraction = VolFrac
                            Sum += VolFrac
                        Next

                        'Normalization is still needed due to minor deviations in the partial molar volume estimation. Summation of partial flow rates
                        'should match phase flow rate.
                        For Each subs As Substancia In .Componentes.Values
                            If Sum > 0 Then
                                subs.VolumetricFraction = subs.VolumetricFraction.GetValueOrDefault / Sum
                            Else
                                subs.VolumetricFraction = 0
                            End If
                            'Corrects volumetric flow rate after normalization of fractions.
                            subs.VolumetricFlow = subs.VolumetricFraction.GetValueOrDefault * TotalVolFlow
                        Next

                    End With
                Next
            End If

            'Totalization for the mixture "phase" should be made separatly, since the concept o partial molar volume is non sense for the whole mixture.
            For Each Subs As Substancia In CurrentMaterialStream.Fases(0).Componentes.Values
                Subs.VolumetricFraction = 0
                Subs.VolumetricFlow = 0
            Next

            'Summation of volumetric flow rates over all phases.
            For APhaseID As Integer = 1 To Me.CurrentMaterialStream.Fases.Count - 1
                For Each Subs As String In CurrentMaterialStream.Fases(APhaseID).Componentes.Keys
                    CurrentMaterialStream.Fases(0).Componentes(Subs).VolumetricFlow = CurrentMaterialStream.Fases(0).Componentes(Subs).VolumetricFlow.GetValueOrDefault + CurrentMaterialStream.Fases(APhaseID).Componentes(Subs).VolumetricFlow.GetValueOrDefault
                Next
            Next

            'Calculate volumetric fractions for the mixture.
            For Each Subs As Substancia In CurrentMaterialStream.Fases(0).Componentes.Values
                Subs.VolumetricFraction = Subs.VolumetricFlow.GetValueOrDefault / CurrentMaterialStream.Fases(0).SPMProperties.volumetric_flow.GetValueOrDefault
            Next

        End Sub

        Public Sub DW_ZerarPhaseProps(ByVal fase As Fase)

            Dim phaseID As Integer

            phaseID = Me.RET_PHASEID(fase)

            If Me.CurrentMaterialStream.Fases.ContainsKey(phaseID) Then

                If phaseID <> 0 Then

                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.density = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.enthalpy = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.entropy = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.molar_enthalpy = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.molar_entropy = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.enthalpyF = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.entropyF = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.molar_enthalpyF = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.molar_entropyF = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.compressibilityFactor = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.heatCapacityCp = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.heatCapacityCv = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.molecularWeight = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.thermalConductivity = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.speedOfSound = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.volumetric_flow = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.jouleThomsonCoefficient = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.excessEnthalpy = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.excessEntropy = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.compressibility = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.bubbleTemperature = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.bubblePressure = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.dewTemperature = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.dewPressure = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.viscosity = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.kinematic_viscosity = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.molarflow = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.massflow = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.massfraction = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.molarfraction = Nothing

                Else

                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.density = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.enthalpy = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.entropy = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.molar_enthalpy = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.molar_entropy = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.enthalpyF = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.entropyF = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.molar_enthalpyF = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.molar_entropyF = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.compressibilityFactor = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.heatCapacityCp = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.heatCapacityCv = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.molecularWeight = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.thermalConductivity = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.speedOfSound = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.volumetric_flow = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.jouleThomsonCoefficient = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.excessEnthalpy = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.excessEntropy = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.compressibility = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.bubbleTemperature = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.bubblePressure = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.dewTemperature = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.dewPressure = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.viscosity = Nothing
                    Me.CurrentMaterialStream.Fases(phaseID).SPMProperties.kinematic_viscosity = Nothing

                End If

            End If

        End Sub

        Public Sub DW_ZerarTwoPhaseProps(ByVal fase1 As Fase, ByVal fase2 As Fase)

            Me.CurrentMaterialStream.Fases(0).TPMProperties.kvalue = Nothing
            Me.CurrentMaterialStream.Fases(0).TPMProperties.logKvalue = Nothing
            Me.CurrentMaterialStream.Fases(0).TPMProperties.surfaceTension = Nothing

        End Sub

        Public Sub DW_ZerarOverallProps()

        End Sub

        Public Sub DW_ZerarVazaoMolar()
            With Me.CurrentMaterialStream
                .Fases(0).SPMProperties.molarflow = 0
            End With
        End Sub

        Public Sub DW_ZerarVazaoVolumetrica()
            With Me.CurrentMaterialStream
                .Fases(0).SPMProperties.volumetric_flow = 0
            End With
        End Sub

        Public Sub DW_ZerarVazaoMassica()
            With Me.CurrentMaterialStream
                .Fases(0).SPMProperties.massflow = 0
            End With
        End Sub

        Public Sub DW_ZerarComposicoes(ByVal fase As Fase)

            Dim phaseID As Integer
            phaseID = Me.RET_PHASEID(fase)

            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia

            For Each subst In Me.CurrentMaterialStream.Fases(phaseID).Componentes.Values
                subst.FracaoMolar = Nothing
                subst.FracaoMassica = Nothing
            Next

        End Sub

        Public Function AUX_TCM(ByVal fase As Fase) As Double

            Dim Tc As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia

            For Each subst In Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).Componentes.Values
                Tc += subst.FracaoMolar.GetValueOrDefault * subst.ConstantProperties.Critical_Temperature
            Next

            Return Tc

        End Function
        Public Function AUX_TBM(ByVal fase As Fase) As Double

            Dim Tb As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia

            For Each subst In Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).Componentes.Values
                Tb += subst.FracaoMolar.GetValueOrDefault * subst.ConstantProperties.Normal_Boiling_Point
            Next

            Return Tb

        End Function

        Public Function AUX_TFM(ByVal fase As Fase) As Double

            Dim Tf As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia

            For Each subst In Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).Componentes.Values
                Tf += subst.FracaoMolar.GetValueOrDefault * subst.ConstantProperties.TemperatureOfFusion
            Next

            Return Tf

        End Function

        Public Function AUX_WM(ByVal fase As Fase) As Double

            Dim val As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia

            For Each subst In Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).Componentes.Values
                val += subst.FracaoMolar.GetValueOrDefault * subst.ConstantProperties.Acentric_Factor
            Next

            Return val

        End Function

        Public Function AUX_ZCM(ByVal fase As Fase) As Double

            Dim val As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia

            For Each subst In Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).Componentes.Values
                val += subst.FracaoMolar.GetValueOrDefault * subst.ConstantProperties.Critical_Compressibility
            Next

            Return val

        End Function

        Public Function AUX_ZRAM(ByVal fase As Fase) As Double

            Dim val As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia

            For Each subst In Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).Componentes.Values
                If subst.ConstantProperties.Z_Rackett <> 0 Then
                    val += subst.FracaoMolar.GetValueOrDefault * subst.ConstantProperties.Z_Rackett
                Else
                    val += subst.FracaoMolar.GetValueOrDefault * (0.29056 - 0.08775 * subst.ConstantProperties.Acentric_Factor)
                End If
            Next

            Return val

        End Function

        Public Function AUX_VCM(ByVal fase As Fase) As Double

            Dim val, vc As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia

            For Each subst In Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).Componentes.Values
                vc = subst.ConstantProperties.Critical_Volume
                If vc = 0.0# Then
                    vc = m_props.Vc(subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Critical_Pressure, subst.ConstantProperties.Acentric_Factor)
                End If
                val += subst.FracaoMolar.GetValueOrDefault * vc
            Next

            Return val / 1000

        End Function

        Public Function AUX_PCM(ByVal fase As Fase) As Double

            Dim val As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia

            For Each subst In Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).Componentes.Values
                val += subst.FracaoMolar.GetValueOrDefault * subst.ConstantProperties.Critical_Pressure
            Next

            Return val

        End Function

        Public Function AUX_PVAPM(ByVal T) As Double

            Dim val As Double = 0
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia

            For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                val += subst.FracaoMolar.GetValueOrDefault * Me.AUX_PVAPi(subst.Nome, T)
            Next

            Return val

        End Function

        Public Function AUX_KIJ(ByVal sub1 As String, ByVal sub2 As String) As Double

            Dim Vc1 As Double = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Critical_Volume
            Dim Vc2 As Double = Me.CurrentMaterialStream.Fases(0).Componentes(sub2).ConstantProperties.Critical_Volume

            Dim tmp As Double = 1 - 8 * (Vc1 * Vc2) ^ 0.5 / ((Vc1 ^ (1 / 3) + Vc2 ^ (1 / 3)) ^ 3)

            Return tmp

            Return 0

        End Function

        Public Function AUX_MMM(ByVal fase As Fase) As Double

            Dim val As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia

            For Each subst In Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).Componentes.Values
                val += subst.FracaoMolar.GetValueOrDefault * subst.ConstantProperties.Molar_Weight
            Next

            Return val

        End Function

        Private Function AUX_Rackett_PHIi(ByVal sub1 As String, ByVal fase As Fase) As Double

            Dim val, vc As Double
            vc = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Critical_Volume
            If vc = 0.0# Then vc = m_props.Vc(Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).Componentes(sub1).ConstantProperties.Critical_Temperature, Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).Componentes(sub1).ConstantProperties.Critical_Pressure, Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).Componentes(sub1).ConstantProperties.Acentric_Factor)

            val = Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).Componentes(sub1).FracaoMolar.GetValueOrDefault * vc

            val = val / Me.AUX_VCM(fase) / 1000

            Return val

        End Function

        Private Function AUX_Rackett_Kij(ByVal sub1 As String, ByVal sub2 As String) As Double

            Dim Vc1 As Double = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Critical_Volume
            Dim Vc2 As Double = Me.CurrentMaterialStream.Fases(0).Componentes(sub2).ConstantProperties.Critical_Volume

            If Vc1 = 0.0# Then Vc1 = m_props.Vc(Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Critical_Temperature, Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Critical_Pressure, Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Acentric_Factor)
            If Vc2 = 0.0# Then Vc2 = m_props.Vc(Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Critical_Temperature, Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Critical_Pressure, Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Acentric_Factor)

            Dim tmp As Double = 8 * (Vc1 * Vc2) ^ 0.5 / ((Vc1 ^ (1 / 3) + Vc2 ^ (1 / 3)) ^ 3)

            Return tmp

        End Function

        Private Function AUX_Rackett_Tcij(ByVal sub1 As String, ByVal sub2 As String) As Double

            Dim Tc1 As Double = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Critical_Temperature
            Dim Tc2 As Double = Me.CurrentMaterialStream.Fases(0).Componentes(sub2).ConstantProperties.Critical_Temperature

            Dim tmp As Double = Me.AUX_Rackett_Kij(sub1, sub2) * (Tc1 * Tc2) ^ 0.5

            Return tmp

        End Function

        Private Function AUX_Rackett_Tcm(ByVal fase As Fase) As Double

            Dim Tc As Double
            Dim subst1, subst2 As DTL.ClassesBasicasTermodinamica.Substancia

            For Each subst1 In Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).Componentes.Values
                For Each subst2 In Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).Componentes.Values
                    Tc += Me.AUX_Rackett_PHIi(subst1.Nome, fase) * Me.AUX_Rackett_PHIi(subst2.Nome, fase) * Me.AUX_Rackett_Tcij(subst1.Nome, subst2.Nome)
                Next
            Next

            Return Tc

        End Function

        Public Function AUX_CPi(ByVal sub1 As String, ByVal T As Double)

            If Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.IsPF = 1 Then

                With Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties

                    Return Me.m_props.Cpig_lk(.PF_Watson_K, .Acentric_Factor, T) '* .Molar_Weight

                End With

            Else

                If Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.OriginalDB = "DWSIM" Or _
                Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.OriginalDB = "" Then
                    Dim A, B, C, D, E, result As Double
                    A = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_A
                    B = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_B
                    C = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_C
                    D = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_D
                    E = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_E
                    'Cp = A + B*T + C*T^2 + D*T^3 + E*T^4 where Cp in kJ/kg-mol , T in K 
                    result = A + B * T + C * T ^ 2 + D * T ^ 3 + E * T ^ 4
                    Return result / Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Molar_Weight
                ElseIf Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.OriginalDB = "CheResources" Then
                    Dim A, B, C, D, E, result As Double
                    A = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_A
                    B = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_B
                    C = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_C
                    D = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_D
                    E = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_E
                    'CAL/MOL.K [CP=A+(B*T)+(C*T^2)+(D*T^3)], T in K
                    result = A + B * T + C * T ^ 2 + D * T ^ 3
                    Return result / Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Molar_Weight * 4.1868
                ElseIf Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.OriginalDB = "ChemSep" Or _
                Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.OriginalDB = "User" Then
                    Dim A, B, C, D, E, result As Double
                    Dim eqno As String = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.IdealgasCpEquation
                    Dim mw As Double = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Molar_Weight
                    A = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_A
                    B = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_B
                    C = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_C
                    D = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_D
                    E = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_E
                    '<rppc name="Ideal gas heat capacity (RPP)"  units="J/kmol/K" >
                    result = Me.CalcCSTDepProp(eqno, A, B, C, D, E, T, 0) / 1000 / mw 'kJ/kg.K
                    Return result
                ElseIf Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.OriginalDB = "Biodiesel" Then
                    Dim A, B, C, D, E, result As Double
                    Dim eqno As String = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.IdealgasCpEquation
                    A = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_A
                    B = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_B
                    C = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_C
                    D = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_D
                    E = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_E
                    result = Me.CalcCSTDepProp(eqno, A, B, C, D, E, T, 0) 'kJ/kg.K
                    Return result
                Else
                    Return 0
                End If

            End If

        End Function

        Public Function AUX_CPm(ByVal fase As Fase, ByVal T As Double) As Double

            Dim val As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia

            For Each subst In Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).Componentes.Values
                val += subst.FracaoMassica.GetValueOrDefault * Me.AUX_CPi(subst.Nome, T)
            Next

            Return val

        End Function

        Public Overridable Function AUX_HFm25(ByVal fase As Fase) As Double

            Dim val As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia

            For Each subst In Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).Componentes.Values
                val += subst.FracaoMolar.GetValueOrDefault * subst.ConstantProperties.IG_Enthalpy_of_Formation_25C * subst.ConstantProperties.Molar_Weight
            Next

            Dim mw As Double = Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).SPMProperties.molecularWeight.GetValueOrDefault

            If mw <> 0.0# Then Return val / mw Else Return 0.0#

        End Function

        Public Overridable Function AUX_SFm25(ByVal fase As Fase) As Double

            Dim val As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia

            For Each subst In Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).Componentes.Values
                val += subst.FracaoMolar.GetValueOrDefault * subst.ConstantProperties.IG_Entropy_of_Formation_25C * subst.ConstantProperties.Molar_Weight
            Next

            Dim mw As Double = Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).SPMProperties.molecularWeight.GetValueOrDefault

            If mw <> 0.0# Then Return val / mw Else Return 0.0#

        End Function

        Public Function AUX_KHenry(ByVal CompName As String, ByVal T As Double) As Double

            Dim KHx As Double
            Dim MW As Double = 18 'mol weight of water [g/mol]
            Dim DW As Double = 996 'density of water at 298.15 K [Kg/m3]
            Dim KHCP As Double = 0.0000064 'nitrogen [mol/m3/Pa]
            Dim C As Double = 1600 'nitrogen
            Dim CAS As String

            CAS = Me.CurrentMaterialStream.Fases(0).Componentes(CompName).ConstantProperties.CAS_Number

            If m_Henry.ContainsKey(CAS) Then
                KHCP = m_Henry(CAS).KHcp
                C = m_Henry(CAS).C
            End If
            KHx = 1 / (KHCP * MW / DW / 1000 * Exp(C * (1 / T - 1 / 298.15)))

            Return KHx '[Pa]

        End Function

        Public Function AUX_PVAPi(ByVal sub1 As String, ByVal T As Double)

            If Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.IsPF = 1 Then

                With Me.CurrentMaterialStream.Fases(0).Componentes(sub1)

                    Return Me.m_props.Pvp_leekesler(T, .ConstantProperties.Critical_Temperature, .ConstantProperties.Critical_Pressure, .ConstantProperties.Acentric_Factor)

                End With

            Else


                If Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.OriginalDB = "DWSIM" Or _
                Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.OriginalDB = "" Then
                    Dim A, B, C, D, E, result As Double
                    A = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Vapor_Pressure_Constant_A
                    B = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Vapor_Pressure_Constant_B
                    C = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Vapor_Pressure_Constant_C
                    D = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Vapor_Pressure_Constant_D
                    E = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Vapor_Pressure_Constant_E
                    result = Math.Exp(A + B / T + C * Math.Log(T) + D * T ^ E)
                    Return result
                ElseIf Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.OriginalDB = "CheResources" Then
                    Dim A, B, C, result As Double
                    A = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Vapor_Pressure_Constant_A
                    B = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Vapor_Pressure_Constant_B
                    C = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Vapor_Pressure_Constant_C
                    '[LN(P)=A-B/(T+C), P(mmHG) T(K)]
                    result = Math.Exp(A - B / (T + C)) * 133.322368 'mmHg to Pascal
                    Return result
                ElseIf Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.OriginalDB = "ChemSep" Or _
                Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.OriginalDB = "User" Then
                    Dim A, B, C, D, E, result As Double
                    Dim eqno As String = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.VaporPressureEquation
                    Dim mw As Double = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Molar_Weight
                    A = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Vapor_Pressure_Constant_A
                    B = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Vapor_Pressure_Constant_B
                    C = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Vapor_Pressure_Constant_C
                    D = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Vapor_Pressure_Constant_D
                    E = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Vapor_Pressure_Constant_E
                    '<vp_c name="Vapour pressure"  units="Pa" >
                    result = Me.CalcCSTDepProp(eqno, A, B, C, D, E, T, 0) 'Pa
                    If eqno = "0" Then
                        With Me.CurrentMaterialStream.Fases(0).Componentes(sub1)
                            result = Me.m_props.Pvp_leekesler(T, .ConstantProperties.Critical_Temperature, .ConstantProperties.Critical_Pressure, .ConstantProperties.Acentric_Factor)
                        End With
                    End If
                    Return result
                ElseIf Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.OriginalDB = "Biodiesel" Then
                    Dim A, B, C, D, E, result As Double
                    Dim eqno As String = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.VaporPressureEquation
                    A = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Vapor_Pressure_Constant_A
                    B = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Vapor_Pressure_Constant_B
                    C = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Vapor_Pressure_Constant_C
                    D = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Vapor_Pressure_Constant_D
                    E = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Vapor_Pressure_Constant_E
                    result = Me.CalcCSTDepProp(eqno, A, B, C, D, E, T, 0) 'kPa
                    Return result * 1000
                Else
                    Return 0
                End If

            End If

        End Function

        Public Function AUX_PVAPi(ByVal index As Integer, ByVal T As Double)

            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim nome As String = ""
            Dim i As Integer = 0

            For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                If i = index Then nome = subst.Nome
                i += 1
            Next
            If Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.IsPF = 1 Then
                With Me.CurrentMaterialStream.Fases(0).Componentes(nome)
                    Return Me.m_props.Pvp_leekesler(T, .ConstantProperties.Critical_Temperature, .ConstantProperties.Critical_Pressure, .ConstantProperties.Acentric_Factor)
                End With
            Else
                If Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.OriginalDB = "DWSIM" Or _
                Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.OriginalDB = "" Then
                    Dim A, B, C, D, E, result As Double
                    A = Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.Vapor_Pressure_Constant_A
                    B = Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.Vapor_Pressure_Constant_B
                    C = Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.Vapor_Pressure_Constant_C
                    D = Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.Vapor_Pressure_Constant_D
                    E = Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.Vapor_Pressure_Constant_E
                    result = Math.Exp(A + B / T + C * Math.Log(T) + D * T ^ E)
                    Return result
                ElseIf Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.OriginalDB = "CheResources" Then
                    Dim A, B, C, result As Double
                    A = Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.Vapor_Pressure_Constant_A
                    B = Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.Vapor_Pressure_Constant_B
                    C = Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.Vapor_Pressure_Constant_C
                    '[LN(P)=A-B/(T+C), P(mmHG) T(K)]
                    result = Math.Exp(A - B / (T + C)) * 133.322368 'mmHg to Pascal
                    Return result
                ElseIf Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.OriginalDB = "ChemSep" Or _
                Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.OriginalDB = "User" Then
                    Dim A, B, C, D, E, result As Double
                    Dim eqno As String = Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.VaporPressureEquation
                    Dim mw As Double = Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.Molar_Weight
                    A = Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.Vapor_Pressure_Constant_A
                    B = Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.Vapor_Pressure_Constant_B
                    C = Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.Vapor_Pressure_Constant_C
                    D = Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.Vapor_Pressure_Constant_D
                    E = Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.Vapor_Pressure_Constant_E
                    '<vp_c name="Vapour pressure"  units="Pa" >
                    result = Me.CalcCSTDepProp(eqno, A, B, C, D, E, T, 0) 'Pa.s
                    If eqno = "0" Then
                        With Me.CurrentMaterialStream.Fases(0).Componentes(nome)
                            result = Me.m_props.Pvp_leekesler(T, .ConstantProperties.Critical_Temperature, .ConstantProperties.Critical_Pressure, .ConstantProperties.Acentric_Factor)
                        End With
                    End If
                    Return result
                ElseIf Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.OriginalDB = "Biodiesel" Then
                    Dim A, B, C, D, E, result As Double
                    Dim eqno As String = Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.VaporPressureEquation
                    A = Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.Vapor_Pressure_Constant_A
                    B = Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.Vapor_Pressure_Constant_B
                    C = Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.Vapor_Pressure_Constant_C
                    D = Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.Vapor_Pressure_Constant_D
                    E = Me.CurrentMaterialStream.Fases(0).Componentes(nome).ConstantProperties.Vapor_Pressure_Constant_E
                    result = Me.CalcCSTDepProp(eqno, A, B, C, D, E, T, 0) 'kPa
                    Return result * 1000
                Else
                    Return 0
                End If
            End If

        End Function

        Public Function AUX_TSATi(ByVal PVAP As Double, ByVal subst As String) As Double

            Dim i As Integer

            Dim Tinf, Tsup As Double

            Dim fT, fT_inf, nsub, delta_T As Double

            Tinf = 100
            Tsup = 2000

            nsub = 10

            delta_T = (Tsup - Tinf) / nsub

            i = 0
            Do
                i = i + 1
                fT = PVAP - Me.AUX_PVAPi(subst, Tinf)
                Tinf = Tinf + delta_T
                fT_inf = PVAP - Me.AUX_PVAPi(subst, Tinf)
            Loop Until fT * fT_inf < 0 Or fT_inf > fT Or i >= 100
            Tsup = Tinf
            Tinf = Tinf - delta_T

            'método de Brent para encontrar Vc

            Dim aaa, bbb, ccc, ddd, eee, min11, min22, faa, fbb, fcc, ppp, qqq, rrr, sss, tol11, xmm As Double
            Dim ITMAX2 As Integer = 100
            Dim iter2 As Integer

            aaa = Tinf
            bbb = Tsup
            ccc = Tsup

            faa = PVAP - Me.AUX_PVAPi(subst, aaa)
            fbb = PVAP - Me.AUX_PVAPi(subst, bbb)
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
                tol11 = 0.0001
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
                fbb = PVAP - Me.AUX_PVAPi(subst, bbb)
                iter2 += 1
            Loop Until iter2 = ITMAX2

Final3:

            Return bbb

        End Function

        Public Function AUX_TSATi(ByVal PVAP As Double, ByVal index As Integer) As Double

            Dim i As Integer

            Dim Tinf, Tsup As Double

            Dim fT, fT_inf, nsub, delta_T As Double

            Tinf = 100
            Tsup = 2000

            nsub = 10

            delta_T = (Tsup - Tinf) / nsub

            i = 0
            Do
                i = i + 1
                fT = PVAP - Me.AUX_PVAPi(index, Tinf)
                Tinf = Tinf + delta_T
                fT_inf = PVAP - Me.AUX_PVAPi(index, Tinf)
            Loop Until fT * fT_inf < 0 Or fT_inf > fT Or i >= 100
            Tsup = Tinf
            Tinf = Tinf - delta_T

            'método de Brent para encontrar Vc

            Dim aaa, bbb, ccc, ddd, eee, min11, min22, faa, fbb, fcc, ppp, qqq, rrr, sss, tol11, xmm As Double
            Dim ITMAX2 As Integer = 100
            Dim iter2 As Integer

            aaa = Tinf
            bbb = Tsup
            ccc = Tsup

            faa = PVAP - Me.AUX_PVAPi(index, aaa)
            fbb = PVAP - Me.AUX_PVAPi(index, bbb)
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
                tol11 = 0.0001
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
                fbb = PVAP - Me.AUX_PVAPi(index, bbb)
                iter2 += 1
            Loop Until iter2 = ITMAX2

Final3:

            Return bbb

        End Function

        Public Function AUX_LIQVISCi(ByVal sub1 As String, ByVal T As Double)

            If T / Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Critical_Temperature < 1 Then

                If Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.IsPF = 1 Then

                    With Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties

                        Return Me.m_props.oilvisc_twu(T, .PF_Tv1, .PF_Tv2, .PF_v1, .PF_v2) * Me.m_props.liq_dens_rackett(T, .Critical_Temperature, .Critical_Pressure, .Acentric_Factor, .Molar_Weight, .Z_Rackett)

                    End With

                Else
                    If Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.OriginalDB = "DWSIM" Or _
                    Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.OriginalDB = "" Then
                        Dim A, B, C, D, E, result As Double
                        A = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Liquid_Viscosity_Const_A
                        B = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Liquid_Viscosity_Const_B
                        C = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Liquid_Viscosity_Const_C
                        D = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Liquid_Viscosity_Const_D
                        E = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Liquid_Viscosity_Const_E
                        result = Math.Exp(A + B / T + C * Math.Log(T) + D * T ^ E)
                        Return result
                    ElseIf Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.OriginalDB = "CheResources" Then
                        Dim B, C, result As Double
                        B = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Liquid_Viscosity_Const_B
                        C = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Liquid_Viscosity_Const_C
                        '[LOG(V)=B*(1/T-1/C), T(K) V(CP)]
                        result = Exp(B * (1 / T - 1 / C)) * 0.001
                        Return result
                    ElseIf Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.OriginalDB = "ChemSep" Or _
                    Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.OriginalDB = "User" Then
                        Dim A, B, C, D, E, result As Double
                        Dim eqno As String = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.LiquidViscosityEquation
                        Dim mw As Double = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Molar_Weight
                        A = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Liquid_Viscosity_Const_A
                        B = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Liquid_Viscosity_Const_B
                        C = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Liquid_Viscosity_Const_C
                        D = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Liquid_Viscosity_Const_D
                        E = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Liquid_Viscosity_Const_E
                        '<lvsc name="Liquid viscosity"  units="Pa.s" >
                        result = Me.CalcCSTDepProp(eqno, A, B, C, D, E, T, 0) 'Pa.s
                        If eqno = "0" Or eqno = "" Then
                            Dim Tc, Pc, w As Double
                            Tc = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Critical_Temperature
                            Pc = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Critical_Pressure
                            w = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Acentric_Factor
                            mw = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Molar_Weight
                            result = Me.m_props.viscl_letsti(T, Tc, Pc, w, mw)
                        End If
                        Return result
                    ElseIf Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.OriginalDB = "Biodiesel" Then
                        Dim result As Double
                        Dim Tc, Pc, w, Mw As Double
                        Tc = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Critical_Temperature
                        Pc = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Critical_Pressure
                        w = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Acentric_Factor
                        Mw = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Molar_Weight
                        result = Me.m_props.viscl_letsti(T, Tc, Pc, w, Mw)
                        Return result
                    Else
                        Return 0
                    End If

                End If

            Else

                Return 0.0#

            End If

        End Function

        Public Function AUX_LIQVISCm(ByVal T As Double, Optional ByVal phaseid As Integer = 3) As Double

            Dim val, val2, logvisc, result As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia

            val = 0
            val2 = 0
            For Each subst In Me.CurrentMaterialStream.Fases(phaseid).Componentes.Values
                'logvisc = Math.Log(Me.AUX_LIQVISCi(subst.Nome, T))
                If Not Double.IsInfinity(logvisc) Then
                    val += subst.FracaoMolar.GetValueOrDefault * Me.AUX_LIQVISCi(subst.Nome, T)
                Else
                    val2 += subst.FracaoMolar.GetValueOrDefault
                End If
            Next

            result = (val / (1 - val2))

            Return result

        End Function

        Public Overridable Function AUX_SOLIDDENS() As Double

            Dim val As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim zerodens As Double = 0
            Dim db As String
            Dim T As Double = Me.CurrentMaterialStream.Fases(0).SPMProperties.temperature.GetValueOrDefault

            For Each subst In Me.CurrentMaterialStream.Fases(7).Componentes.Values
                db = subst.ConstantProperties.OriginalDB
                If db = "ChemSep" Or (db = "User" And subst.ConstantProperties.SolidDensityEquation <> "") Then
                    Dim A, B, C, D, E, result As Double
                    Dim eqno As String = subst.ConstantProperties.SolidDensityEquation
                    Dim mw As Double = subst.ConstantProperties.Molar_Weight
                    A = subst.ConstantProperties.Solid_Density_Const_A
                    B = subst.ConstantProperties.Solid_Density_Const_B
                    C = subst.ConstantProperties.Solid_Density_Const_C
                    D = subst.ConstantProperties.Solid_Density_Const_D
                    E = subst.ConstantProperties.Solid_Density_Const_E
                    result = Me.CalcCSTDepProp(eqno, A, B, C, D, E, T, 0) 'kmol/m3
                    val += subst.FracaoMassica.GetValueOrDefault * 1 / (result * mw)
                Else
                    If subst.ConstantProperties.SolidDensityAtTs <> 0.0# Then
                        val += subst.FracaoMassica.GetValueOrDefault * 1 / subst.ConstantProperties.SolidDensityAtTs
                    Else
                        zerodens += subst.FracaoMassica.GetValueOrDefault
                    End If
                End If
            Next

            Return 1 / val / (1 - zerodens)

        End Function

        Public Overridable Function AUX_SOLIDDENSi(cprop As ConstantProperties, T As Double) As Double

            Dim val As Double
            Dim zerodens As Double = 0

            If cprop.OriginalDB = "ChemSep" Or (cprop.OriginalDB = "User" And cprop.SolidDensityEquation <> "") Then
                Dim A, B, C, D, E, result As Double
                Dim eqno As String = cprop.SolidDensityEquation
                Dim mw As Double = cprop.Molar_Weight
                A = cprop.Solid_Density_Const_A
                B = cprop.Solid_Density_Const_B
                C = cprop.Solid_Density_Const_C
                D = cprop.Solid_Density_Const_D
                E = cprop.Solid_Density_Const_E
                result = Me.CalcCSTDepProp(eqno, A, B, C, D, E, T, 0) 'kmol/m3
                val = 1 / (result * mw)
            Else
                If cprop.SolidDensityAtTs <> 0.0# Then
                    val = 1 / cprop.SolidDensityAtTs
                Else
                    val = 1.0E+20
                End If
            End If

            Return 1 / val

        End Function

        Public Overridable Function AUX_SolidHeatCapacity(ByVal cprop As ConstantProperties, ByVal T As Double) As Double

            Dim val As Double

            If cprop.OriginalDB = "ChemSep" Or (cprop.OriginalDB = "User" And cprop.SolidDensityEquation <> "") Then
                Dim A, B, C, D, E, result As Double
                Dim eqno As String = cprop.SolidHeatCapacityEquation
                Dim mw As Double = cprop.Molar_Weight
                A = cprop.Solid_Heat_Capacity_Const_A
                B = cprop.Solid_Heat_Capacity_Const_B
                C = cprop.Solid_Heat_Capacity_Const_C
                D = cprop.Solid_Heat_Capacity_Const_D
                E = cprop.Solid_Heat_Capacity_Const_E
                result = Me.CalcCSTDepProp(eqno, A, B, C, D, E, T, 0) 'J/kmol/K
                val = result / 1000 / mw 'kJ/kg.K
            Else
                val = 3 ' replacement if no params available
            End If

            Return val

        End Function

        Public Function AUX_SOLIDCP(ByVal Vxm As Array, ByVal cprops As List(Of ConstantProperties), ByVal T As Double) As Double

            Dim n As Integer = UBound(Vxm)
            Dim val As Double = 0
            For i As Integer = 0 To n
                val += Vxm(i) * AUX_SolidHeatCapacity(cprops(i), T)
            Next
            Return val

        End Function

        Public Function AUX_SURFTM(ByVal T As Double) As Double

            If m_props Is Nothing Then m_props = New DTL.SimulationObjects.PropertyPackages.Auxiliary.PROPS

            Dim val As Double = 0
            Dim nbp As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim ftotal As Double = 1

            For Each subst In Me.CurrentMaterialStream.Fases(1).Componentes.Values
                If T / subst.ConstantProperties.Critical_Temperature < 1 Then
                    With subst.ConstantProperties
                        If .SurfaceTensionEquation <> "" And .SurfaceTensionEquation <> "0" And Not .IsIon And Not .IsSalt Then
                            subst.TDProperties.surfaceTension = Me.CalcCSTDepProp(.SurfaceTensionEquation, .Surface_Tension_Const_A, .Surface_Tension_Const_B, .Surface_Tension_Const_C, .Surface_Tension_Const_D, .Surface_Tension_Const_E, T, .Critical_Temperature)
                        ElseIf .IsIon Or .IsSalt Then
                            subst.TDProperties.surfaceTension = 0.0#
                        Else
                            nbp = subst.ConstantProperties.Normal_Boiling_Point
                            If nbp = 0 Then nbp = 0.7 * subst.ConstantProperties.Critical_Temperature
                            subst.TDProperties.surfaceTension = Me.m_props.sigma_bb(T, nbp, subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Critical_Pressure)
                        End If
                    End With
                Else
                    subst.TDProperties.surfaceTension = 0
                    ftotal -= subst.FracaoMolar.GetValueOrDefault
                End If
                val += subst.FracaoMolar.GetValueOrDefault * subst.TDProperties.surfaceTension.GetValueOrDefault / ftotal
            Next

            Return val

        End Function

        Public Function AUX_SURFTi(ByVal constprop As ConstantProperties, ByVal T As Double) As Double

            If m_props Is Nothing Then m_props = New DTL.SimulationObjects.PropertyPackages.Auxiliary.PROPS

            Dim val As Double = 0
            Dim nbp As Double
            Dim ftotal As Double = 1

            If T / constprop.Critical_Temperature < 1 Then
                With constprop
                    If .SurfaceTensionEquation <> "" And .SurfaceTensionEquation <> "0" And Not .IsIon And Not .IsSalt Then
                        val = Me.CalcCSTDepProp(.SurfaceTensionEquation, .Surface_Tension_Const_A, .Surface_Tension_Const_B, .Surface_Tension_Const_C, .Surface_Tension_Const_D, .Surface_Tension_Const_E, T, .Critical_Temperature)
                    ElseIf .IsIon Or .IsSalt Then
                        val = 0.0#
                    Else
                        nbp = constprop.Normal_Boiling_Point
                        If nbp = 0 Then nbp = 0.7 * constprop.Critical_Temperature
                        val = Me.m_props.sigma_bb(T, nbp, constprop.Critical_Temperature, constprop.Critical_Pressure)
                    End If
                End With
            Else
            End If

            Return val

        End Function

        Public Function AUX_CONDTL(ByVal T As Double, Optional ByVal phaseid As Integer = 3) As Double

            If m_props Is Nothing Then m_props = New DTL.SimulationObjects.PropertyPackages.Auxiliary.PROPS

            Dim val As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim vcl(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1)
            Dim i As Integer = 0
            For Each subst In Me.CurrentMaterialStream.Fases(phaseid).Componentes.Values
                If Me.Parameters.ContainsKey("PP_USEEXPLIQTHERMALCOND") Then
                    If CInt(Me.Parameters("PP_USEEXPLIQTHERMALCOND")) = 1 Then
                        If subst.ConstantProperties.LiquidThermalConductivityEquation <> "" Then
                            vcl(i) = Me.CalcCSTDepProp(subst.ConstantProperties.LiquidThermalConductivityEquation, subst.ConstantProperties.Liquid_Thermal_Conductivity_Const_A, subst.ConstantProperties.Liquid_Thermal_Conductivity_Const_B, subst.ConstantProperties.Liquid_Thermal_Conductivity_Const_C, subst.ConstantProperties.Liquid_Thermal_Conductivity_Const_D, subst.ConstantProperties.Liquid_Thermal_Conductivity_Const_E, T, subst.ConstantProperties.Critical_Temperature)
                        Else
                            vcl(i) = Me.m_props.condl_latini(T, subst.ConstantProperties.Normal_Boiling_Point, subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Molar_Weight, "")
                        End If
                    Else
                        vcl(i) = Me.m_props.condl_latini(T, subst.ConstantProperties.Normal_Boiling_Point, subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Molar_Weight, "")
                    End If
                Else
                    vcl(i) = Me.m_props.condl_latini(T, subst.ConstantProperties.Normal_Boiling_Point, subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Molar_Weight, "")
                End If
                i = i + 1
            Next
            val = Me.m_props.condlm_li(Me.RET_VVC, vcl, Me.RET_VMOL(Me.RET_PHASECODE(phaseid)))
            Return val

        End Function

        Public Function AUX_CONDTG(ByVal T As Double, ByVal P As Double) As Double

            If m_props Is Nothing Then m_props = New DTL.SimulationObjects.PropertyPackages.Auxiliary.PROPS

            Dim val As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim i As Integer = 0
            For Each subst In Me.CurrentMaterialStream.Fases(2).Componentes.Values
                If subst.ConstantProperties.VaporThermalConductivityEquation <> "" Then
                    val += subst.FracaoMolar.GetValueOrDefault * Me.CalcCSTDepProp(subst.ConstantProperties.VaporThermalConductivityEquation, subst.ConstantProperties.Vapor_Thermal_Conductivity_Const_A, subst.ConstantProperties.Vapor_Thermal_Conductivity_Const_B, subst.ConstantProperties.Vapor_Thermal_Conductivity_Const_C, subst.ConstantProperties.Vapor_Thermal_Conductivity_Const_D, subst.ConstantProperties.Vapor_Thermal_Conductivity_Const_E, T, subst.ConstantProperties.Critical_Temperature)
                Else
                    val += subst.FracaoMolar.GetValueOrDefault * Me.m_props.condtg_elyhanley(T, Me.AUX_TCM(Fase.Vapor), Me.AUX_VCM(Fase.Vapor), Me.AUX_ZCM(Fase.Vapor), Me.AUX_WM(Fase.Vapor), Me.AUX_MMM(Fase.Vapor), Me.DW_CalcCv_ISOL(Fase.Vapor, T, P) * Me.AUX_MMM(Fase.Vapor))
                End If
                i = i + 1
            Next

            Return val

        End Function

        Public Function AUX_LIQTHERMCONDi(ByVal cprop As ConstantProperties, ByVal T As Double) As Double

            Dim val As Double

            If cprop.LiquidThermalConductivityEquation <> "" And cprop.LiquidThermalConductivityEquation <> "0" And Not cprop.IsIon And Not cprop.IsSalt Then
                val = Me.CalcCSTDepProp(cprop.LiquidThermalConductivityEquation, cprop.Liquid_Thermal_Conductivity_Const_A, cprop.Liquid_Thermal_Conductivity_Const_B, cprop.Liquid_Thermal_Conductivity_Const_C, cprop.Liquid_Thermal_Conductivity_Const_D, cprop.Liquid_Thermal_Conductivity_Const_E, T, cprop.Critical_Temperature)
            ElseIf cprop.IsIon Or cprop.IsSalt Then
                val = 0.0#
            Else
                val = Me.m_props.condl_latini(T, cprop.Normal_Boiling_Point, cprop.Critical_Temperature, cprop.Molar_Weight, "")
            End If

            Return val

        End Function

        Public Function AUX_VAPTHERMCONDi(ByVal cprop As ConstantProperties, ByVal T As Double, ByVal P As Double) As Double

            Dim val As Double

            If cprop.VaporThermalConductivityEquation <> "" And cprop.VaporThermalConductivityEquation <> "0" And Not cprop.IsIon And Not cprop.IsSalt Then
                val = Me.CalcCSTDepProp(cprop.VaporThermalConductivityEquation, cprop.Vapor_Thermal_Conductivity_Const_A, cprop.Vapor_Thermal_Conductivity_Const_B, cprop.Vapor_Thermal_Conductivity_Const_C, cprop.Vapor_Thermal_Conductivity_Const_D, cprop.Vapor_Thermal_Conductivity_Const_E, T, cprop.Critical_Temperature)
            ElseIf cprop.IsIon Or cprop.IsSalt Then
                val = 0.0#
            Else
                val = Me.m_props.condtg_elyhanley(T, cprop.Critical_Temperature, cprop.Critical_Volume / 1000, cprop.Critical_Compressibility, cprop.Acentric_Factor, cprop.Molar_Weight, Me.AUX_CPi(cprop.Name, T) * cprop.Molar_Weight - 8.314)
            End If

            Return val

        End Function

        Public Function AUX_VAPVISCm(ByVal T As Double, ByVal RHO As Double, ByVal MM As Double) As Double

            If m_props Is Nothing Then m_props = New DTL.SimulationObjects.PropertyPackages.Auxiliary.PROPS

            Dim val As Double = 0.0#

            For Each subst As Substancia In Me.CurrentMaterialStream.Fases(2).Componentes.Values
                val += subst.FracaoMolar.GetValueOrDefault * Me.AUX_VAPVISCi(subst.ConstantProperties, T)
            Next

            val = Me.m_props.viscg_jossi_stiel_thodos(val, T, MM / RHO / 1000, AUX_TCM(Fase.Vapor), AUX_PCM(Fase.Vapor), AUX_VCM(Fase.Vapor), AUX_MMM(Fase.Vapor))

            Return val

        End Function

        Public Function AUX_VAPVISCi(ByVal cprop As ConstantProperties, ByVal T As Double) As Double

            Dim val As Double

            If cprop.VaporViscosityEquation <> "" And cprop.VaporViscosityEquation <> "0" And Not cprop.IsIon And Not cprop.IsSalt Then
                val = Me.CalcCSTDepProp(cprop.VaporViscosityEquation, cprop.Vapor_Viscosity_Const_A, cprop.Vapor_Viscosity_Const_B, cprop.Vapor_Viscosity_Const_C, cprop.Vapor_Viscosity_Const_D, cprop.Vapor_Viscosity_Const_E, T, cprop.Critical_Temperature)
            ElseIf cprop.IsIon Or cprop.IsSalt Then
                val = 0.0#
            Else
                val = Me.m_props.viscg_lucas(T, cprop.Critical_Temperature, cprop.Critical_Pressure, cprop.Acentric_Factor, cprop.Molar_Weight)
            End If

            Return val

        End Function

        Public Overridable Function AUX_LIQDENS(ByVal T As Double, Optional ByVal P As Double = 0, Optional ByVal Pvp As Double = 0, Optional ByVal phaseid As Integer = 3, Optional ByVal FORCE_EOS As Boolean = False) As Double

            If m_props Is Nothing Then m_props = New DTL.SimulationObjects.PropertyPackages.Auxiliary.PROPS

            Dim val As Double
            Dim m_pr2 As New DTL.SimulationObjects.PropertyPackages.Auxiliary.PengRobinson

            If phaseid = 1 Then
                If T / Me.AUX_TCM(Fase.Liquid) > 1 Then

                    Dim Z = m_pr2.Z_PR(T, P, RET_VMOL(Fase.Liquid), RET_VKij(), RET_VTC, RET_VPC, RET_VW, "L")
                    val = 1 / (8.314 * Z * T / P)
                    val = val * Me.AUX_MMM(Fase.Liquid) / 1000

                Else

                    If Me.Parameters.ContainsKey("PP_IDEAL_MIXRULE_LIQDENS") Then
                        If CInt(Me.Parameters("PP_IDEAL_MIXRULE_LIQDENS")) = 1 Then
                            Dim i As Integer
                            Dim vk(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As Double
                            i = 0
                            For Each subst As Substancia In Me.CurrentMaterialStream.Fases(phaseid).Componentes.Values
                                If Me.Parameters.ContainsKey("PP_USEEXPLIQDENS") Then
                                    If CInt(Me.Parameters("PP_USEEXPLIQDENS")) = 1 Then
                                        If subst.ConstantProperties.LiquidDensityEquation <> "" And subst.ConstantProperties.LiquidDensityEquation <> "0" Then
                                            vk(i) = Me.CalcCSTDepProp(subst.ConstantProperties.LiquidDensityEquation, subst.ConstantProperties.Liquid_Density_Const_A, subst.ConstantProperties.Liquid_Density_Const_B, subst.ConstantProperties.Liquid_Density_Const_C, subst.ConstantProperties.Liquid_Density_Const_D, subst.ConstantProperties.Liquid_Density_Const_E, T, subst.ConstantProperties.Critical_Temperature)
                                            vk(i) = subst.ConstantProperties.Molar_Weight * vk(i)
                                        Else
                                            vk(i) = Me.m_props.liq_dens_rackett(T, subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Critical_Pressure, subst.ConstantProperties.Acentric_Factor, subst.ConstantProperties.Molar_Weight, subst.ConstantProperties.Z_Rackett, P, Me.AUX_PVAPi(subst.Nome, T))
                                        End If
                                    Else
                                        vk(i) = Me.m_props.liq_dens_rackett(T, subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Critical_Pressure, subst.ConstantProperties.Acentric_Factor, subst.ConstantProperties.Molar_Weight, subst.ConstantProperties.Z_Rackett, P, Me.AUX_PVAPi(subst.Nome, T))
                                    End If
                                Else
                                    vk(i) = Me.m_props.liq_dens_rackett(T, subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Critical_Pressure, subst.ConstantProperties.Acentric_Factor, subst.ConstantProperties.Molar_Weight, subst.ConstantProperties.Z_Rackett, P, Me.AUX_PVAPi(subst.Nome, T))
                                End If
                                vk(i) = subst.FracaoMassica / vk(i)
                                i = i + 1
                            Next
                            val = 1 / MathEx.Common.Sum(vk)
                        Else
                            val = Me.m_props.liq_dens_rackett(T, Me.AUX_Rackett_Tcm(Fase.Liquid), Me.AUX_PCM(Fase.Liquid), Me.AUX_WM(Fase.Liquid), Me.AUX_MMM(Fase.Liquid), Me.AUX_ZRAM(Fase.Liquid), P, Me.AUX_PVAPM(T))
                        End If
                    Else
                        val = Me.m_props.liq_dens_rackett(T, Me.AUX_Rackett_Tcm(Fase.Liquid), Me.AUX_PCM(Fase.Liquid), Me.AUX_WM(Fase.Liquid), Me.AUX_MMM(Fase.Liquid), Me.AUX_ZRAM(Fase.Liquid), P, Me.AUX_PVAPM(T))
                    End If

                End If

            ElseIf phaseid = 2 Then

                If T / Me.AUX_TCM(Fase.Vapor) > 1 Then

                    Dim Z = m_pr2.Z_PR(T, P, RET_VMOL(Fase.Vapor), RET_VKij(), RET_VTC, RET_VPC, RET_VW, "L")
                    val = 1 / (8.314 * Z * T / P)
                    val = val * Me.AUX_MMM(Fase.Vapor) / 1000

                Else

                    Dim vk(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As Double
                    val = Me.m_props.liq_dens_rackett(T, Me.AUX_Rackett_Tcm(Fase.Vapor), Me.AUX_PCM(Fase.Vapor), Me.AUX_WM(Fase.Vapor), Me.AUX_MMM(Fase.Vapor), Me.AUX_ZRAM(Fase.Vapor), P, Me.AUX_PVAPM(T))

                End If
            ElseIf phaseid = 3 Then
                If T / Me.AUX_TCM(Fase.Liquid1) > 1 Then

                    Dim Z = m_pr2.Z_PR(T, P, RET_VMOL(Fase.Liquid1), RET_VKij(), RET_VTC, RET_VPC, RET_VW, "L")
                    val = 1 / (8.314 * Z * T / P)
                    val = val * Me.AUX_MMM(Fase.Liquid1) / 1000

                Else

                    If Me.Parameters.ContainsKey("PP_IDEAL_MIXRULE_LIQDENS") Then
                        If CInt(Me.Parameters("PP_IDEAL_MIXRULE_LIQDENS")) = 1 Then
                            Dim i As Integer
                            Dim vk(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As Double
                            i = 0
                            For Each subst As Substancia In Me.CurrentMaterialStream.Fases(phaseid).Componentes.Values
                                If Me.Parameters.ContainsKey("PP_USEEXPLIQDENS") Then
                                    If CInt(Me.Parameters("PP_USEEXPLIQDENS")) = 1 Then
                                        If subst.ConstantProperties.LiquidDensityEquation <> "" And subst.ConstantProperties.LiquidDensityEquation <> "0" Then
                                            vk(i) = Me.CalcCSTDepProp(subst.ConstantProperties.LiquidDensityEquation, subst.ConstantProperties.Liquid_Density_Const_A, subst.ConstantProperties.Liquid_Density_Const_B, subst.ConstantProperties.Liquid_Density_Const_C, subst.ConstantProperties.Liquid_Density_Const_D, subst.ConstantProperties.Liquid_Density_Const_E, T, subst.ConstantProperties.Critical_Temperature)
                                            vk(i) = subst.ConstantProperties.Molar_Weight * vk(i)
                                        Else
                                            vk(i) = Me.m_props.liq_dens_rackett(T, subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Critical_Pressure, subst.ConstantProperties.Acentric_Factor, subst.ConstantProperties.Molar_Weight, subst.ConstantProperties.Z_Rackett, P, Me.AUX_PVAPi(subst.Nome, T))
                                        End If
                                    Else
                                        vk(i) = Me.m_props.liq_dens_rackett(T, subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Critical_Pressure, subst.ConstantProperties.Acentric_Factor, subst.ConstantProperties.Molar_Weight, subst.ConstantProperties.Z_Rackett, P, Me.AUX_PVAPi(subst.Nome, T))
                                    End If
                                Else
                                    vk(i) = Me.m_props.liq_dens_rackett(T, subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Critical_Pressure, subst.ConstantProperties.Acentric_Factor, subst.ConstantProperties.Molar_Weight, subst.ConstantProperties.Z_Rackett, P, Me.AUX_PVAPi(subst.Nome, T))
                                End If
                                vk(i) = subst.FracaoMassica / vk(i)
                                i = i + 1
                            Next
                            val = 1 / MathEx.Common.Sum(vk)
                        Else
                            val = Me.m_props.liq_dens_rackett(T, Me.AUX_Rackett_Tcm(Fase.Liquid1), Me.AUX_PCM(Fase.Liquid1), Me.AUX_WM(Fase.Liquid1), Me.AUX_MMM(Fase.Liquid1), Me.AUX_ZRAM(Fase.Liquid1), P, Me.AUX_PVAPM(T))
                        End If
                    Else
                        val = Me.m_props.liq_dens_rackett(T, Me.AUX_Rackett_Tcm(Fase.Liquid1), Me.AUX_PCM(Fase.Liquid1), Me.AUX_WM(Fase.Liquid1), Me.AUX_MMM(Fase.Liquid1), Me.AUX_ZRAM(Fase.Liquid1), P, Me.AUX_PVAPM(T))
                    End If

                End If
            ElseIf phaseid = 4 Then
                If T / Me.AUX_TCM(Fase.Liquid2) > 1 Then

                    Dim Z = m_pr2.Z_PR(T, P, RET_VMOL(Fase.Liquid2), RET_VKij(), RET_VTC, RET_VPC, RET_VW, "L")
                    val = 1 / (8.314 * Z * T / P)
                    val = val * Me.AUX_MMM(Fase.Liquid2) / 1000

                Else

                    If Me.Parameters.ContainsKey("PP_IDEAL_MIXRULE_LIQDENS") Then
                        If CInt(Me.Parameters("PP_IDEAL_MIXRULE_LIQDENS")) = 1 Then
                            Dim i As Integer
                            Dim vk(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As Double
                            i = 0
                            For Each subst As Substancia In Me.CurrentMaterialStream.Fases(phaseid).Componentes.Values
                                If Me.Parameters.ContainsKey("PP_USEEXPLIQDENS") Then
                                    If CInt(Me.Parameters("PP_USEEXPLIQDENS")) = 1 Then
                                        If subst.ConstantProperties.LiquidDensityEquation <> "" And subst.ConstantProperties.LiquidDensityEquation <> "0" Then
                                            vk(i) = Me.CalcCSTDepProp(subst.ConstantProperties.LiquidDensityEquation, subst.ConstantProperties.Liquid_Density_Const_A, subst.ConstantProperties.Liquid_Density_Const_B, subst.ConstantProperties.Liquid_Density_Const_C, subst.ConstantProperties.Liquid_Density_Const_D, subst.ConstantProperties.Liquid_Density_Const_E, T, subst.ConstantProperties.Critical_Temperature)
                                            vk(i) = subst.ConstantProperties.Molar_Weight * vk(i)
                                        Else
                                            vk(i) = Me.m_props.liq_dens_rackett(T, subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Critical_Pressure, subst.ConstantProperties.Acentric_Factor, subst.ConstantProperties.Molar_Weight, subst.ConstantProperties.Z_Rackett, P, Me.AUX_PVAPi(subst.Nome, T))
                                        End If
                                    Else
                                        vk(i) = Me.m_props.liq_dens_rackett(T, subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Critical_Pressure, subst.ConstantProperties.Acentric_Factor, subst.ConstantProperties.Molar_Weight, subst.ConstantProperties.Z_Rackett, P, Me.AUX_PVAPi(subst.Nome, T))
                                    End If
                                Else
                                    vk(i) = Me.m_props.liq_dens_rackett(T, subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Critical_Pressure, subst.ConstantProperties.Acentric_Factor, subst.ConstantProperties.Molar_Weight, subst.ConstantProperties.Z_Rackett, P, Me.AUX_PVAPi(subst.Nome, T))
                                End If
                                vk(i) = subst.FracaoMassica / vk(i)
                                i = i + 1
                            Next
                            val = 1 / MathEx.Common.Sum(vk)
                        Else
                            val = Me.m_props.liq_dens_rackett(T, Me.AUX_Rackett_Tcm(Fase.Liquid2), Me.AUX_PCM(Fase.Liquid2), Me.AUX_WM(Fase.Liquid2), Me.AUX_MMM(Fase.Liquid2), Me.AUX_ZRAM(Fase.Liquid2), P, Me.AUX_PVAPM(T))
                        End If
                    Else
                        val = Me.m_props.liq_dens_rackett(T, Me.AUX_Rackett_Tcm(Fase.Liquid2), Me.AUX_PCM(Fase.Liquid2), Me.AUX_WM(Fase.Liquid2), Me.AUX_MMM(Fase.Liquid2), Me.AUX_ZRAM(Fase.Liquid2), P, Me.AUX_PVAPM(T))
                    End If

                End If
            ElseIf phaseid = 5 Then
                If T / Me.AUX_TCM(Fase.Liquid3) > 1 Then

                    Dim Z = m_pr2.Z_PR(T, P, RET_VMOL(Fase.Liquid3), RET_VKij(), RET_VTC, RET_VPC, RET_VW, "L")
                    val = 1 / (8.314 * Z * T / P)
                    val = val * Me.AUX_MMM(Fase.Liquid3) / 1000

                Else

                    If Me.Parameters.ContainsKey("PP_IDEAL_MIXRULE_LIQDENS") Then
                        If CInt(Me.Parameters("PP_IDEAL_MIXRULE_LIQDENS")) = 1 Then
                            Dim i As Integer
                            Dim vk(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As Double
                            i = 0
                            For Each subst As Substancia In Me.CurrentMaterialStream.Fases(phaseid).Componentes.Values
                                If Me.Parameters.ContainsKey("PP_USEEXPLIQDENS") Then
                                    If CInt(Me.Parameters("PP_USEEXPLIQDENS")) = 1 Then
                                        If subst.ConstantProperties.LiquidDensityEquation <> "" And subst.ConstantProperties.LiquidDensityEquation <> "0" Then
                                            vk(i) = Me.CalcCSTDepProp(subst.ConstantProperties.LiquidDensityEquation, subst.ConstantProperties.Liquid_Density_Const_A, subst.ConstantProperties.Liquid_Density_Const_B, subst.ConstantProperties.Liquid_Density_Const_C, subst.ConstantProperties.Liquid_Density_Const_D, subst.ConstantProperties.Liquid_Density_Const_E, T, subst.ConstantProperties.Critical_Temperature)
                                            vk(i) = subst.ConstantProperties.Molar_Weight * vk(i)
                                        Else
                                            vk(i) = Me.m_props.liq_dens_rackett(T, subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Critical_Pressure, subst.ConstantProperties.Acentric_Factor, subst.ConstantProperties.Molar_Weight, subst.ConstantProperties.Z_Rackett, P, Me.AUX_PVAPi(subst.Nome, T))
                                        End If
                                    Else
                                        vk(i) = Me.m_props.liq_dens_rackett(T, subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Critical_Pressure, subst.ConstantProperties.Acentric_Factor, subst.ConstantProperties.Molar_Weight, subst.ConstantProperties.Z_Rackett, P, Me.AUX_PVAPi(subst.Nome, T))
                                    End If
                                Else
                                    vk(i) = Me.m_props.liq_dens_rackett(T, subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Critical_Pressure, subst.ConstantProperties.Acentric_Factor, subst.ConstantProperties.Molar_Weight, subst.ConstantProperties.Z_Rackett, P, Me.AUX_PVAPi(subst.Nome, T))
                                End If
                                vk(i) = subst.FracaoMassica / vk(i)
                                i = i + 1
                            Next
                            val = 1 / MathEx.Common.Sum(vk)
                        Else
                            val = Me.m_props.liq_dens_rackett(T, Me.AUX_Rackett_Tcm(Fase.Liquid3), Me.AUX_PCM(Fase.Liquid3), Me.AUX_WM(Fase.Liquid3), Me.AUX_MMM(Fase.Liquid3), Me.AUX_ZRAM(Fase.Liquid3), P, Me.AUX_PVAPM(T))
                        End If
                    Else
                        val = Me.m_props.liq_dens_rackett(T, Me.AUX_Rackett_Tcm(Fase.Liquid3), Me.AUX_PCM(Fase.Liquid3), Me.AUX_WM(Fase.Liquid3), Me.AUX_MMM(Fase.Liquid3), Me.AUX_ZRAM(Fase.Liquid3), P, Me.AUX_PVAPM(T))
                    End If

                End If
            ElseIf phaseid = 6 Then
                If T / Me.AUX_TCM(Fase.Aqueous) > 1 Then

                    Dim Z = m_pr2.Z_PR(T, P, RET_VMOL(Fase.Aqueous), RET_VKij(), RET_VTC, RET_VPC, RET_VW, "L")
                    val = 1 / (8.314 * Z * T / P)
                    val = val * Me.AUX_MMM(Fase.Aqueous) / 1000

                Else

                    If Me.Parameters.ContainsKey("PP_IDEAL_MIXRULE_LIQDENS") Then
                        If CInt(Me.Parameters("PP_IDEAL_MIXRULE_LIQDENS")) = 1 Then
                            Dim i As Integer
                            Dim vk(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As Double
                            i = 0
                            For Each subst As Substancia In Me.CurrentMaterialStream.Fases(phaseid).Componentes.Values
                                If Me.Parameters.ContainsKey("PP_USEEXPLIQDENS") Then
                                    If CInt(Me.Parameters("PP_USEEXPLIQDENS")) = 1 Then
                                        If subst.ConstantProperties.LiquidDensityEquation <> "" And subst.ConstantProperties.LiquidDensityEquation <> "0" Then
                                            vk(i) = Me.CalcCSTDepProp(subst.ConstantProperties.LiquidDensityEquation, subst.ConstantProperties.Liquid_Density_Const_A, subst.ConstantProperties.Liquid_Density_Const_B, subst.ConstantProperties.Liquid_Density_Const_C, subst.ConstantProperties.Liquid_Density_Const_D, subst.ConstantProperties.Liquid_Density_Const_E, T, subst.ConstantProperties.Critical_Temperature)
                                            vk(i) = subst.ConstantProperties.Molar_Weight * vk(i)
                                        Else
                                            vk(i) = Me.m_props.liq_dens_rackett(T, subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Critical_Pressure, subst.ConstantProperties.Acentric_Factor, subst.ConstantProperties.Molar_Weight, subst.ConstantProperties.Z_Rackett, P, Me.AUX_PVAPi(subst.Nome, T))
                                        End If
                                    Else
                                        vk(i) = Me.m_props.liq_dens_rackett(T, subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Critical_Pressure, subst.ConstantProperties.Acentric_Factor, subst.ConstantProperties.Molar_Weight, subst.ConstantProperties.Z_Rackett, P, Me.AUX_PVAPi(subst.Nome, T))
                                    End If
                                Else
                                    vk(i) = Me.m_props.liq_dens_rackett(T, subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Critical_Pressure, subst.ConstantProperties.Acentric_Factor, subst.ConstantProperties.Molar_Weight, subst.ConstantProperties.Z_Rackett, P, Me.AUX_PVAPi(subst.Nome, T))
                                End If
                                vk(i) = subst.FracaoMassica / vk(i)
                                i = i + 1
                            Next
                            val = 1 / MathEx.Common.Sum(vk)
                        Else
                            val = Me.m_props.liq_dens_rackett(T, Me.AUX_Rackett_Tcm(Fase.Aqueous), Me.AUX_PCM(Fase.Aqueous), Me.AUX_WM(Fase.Aqueous), Me.AUX_MMM(Fase.Aqueous), Me.AUX_ZRAM(Fase.Aqueous), P, Me.AUX_PVAPM(T))
                        End If
                    Else
                        val = Me.m_props.liq_dens_rackett(T, Me.AUX_Rackett_Tcm(Fase.Aqueous), Me.AUX_PCM(Fase.Aqueous), Me.AUX_WM(Fase.Aqueous), Me.AUX_MMM(Fase.Aqueous), Me.AUX_ZRAM(Fase.Aqueous), P, Me.AUX_PVAPM(T))
                    End If

                End If
            End If

            m_pr2 = Nothing

            Return val

        End Function

        Public Overridable Function AUX_LIQDENS(ByVal T As Double, ByVal Vx As Array, Optional ByVal P As Double = 0, Optional ByVal Pvp As Double = 0, Optional ByVal FORCE_EOS As Boolean = False) As Double

            If m_props Is Nothing Then m_props = New DTL.SimulationObjects.PropertyPackages.Auxiliary.PROPS

            Dim val As Double
            Dim m_pr2 As New DTL.SimulationObjects.PropertyPackages.Auxiliary.PengRobinson

            Dim i As Integer
            Dim vk(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As Double
            i = 0
            For Each subst As Substancia In Me.CurrentMaterialStream.Fases(1).Componentes.Values
                If Me.Parameters.ContainsKey("PP_USEEXPLIQDENS") Then
                    If CInt(Me.Parameters("PP_USEEXPLIQDENS")) = 1 Then
                        If subst.ConstantProperties.LiquidDensityEquation <> "" And subst.ConstantProperties.LiquidDensityEquation <> "0" And Not subst.ConstantProperties.IsIon And Not subst.ConstantProperties.IsSalt Then
                            vk(i) = Me.CalcCSTDepProp(subst.ConstantProperties.LiquidDensityEquation, subst.ConstantProperties.Liquid_Density_Const_A, subst.ConstantProperties.Liquid_Density_Const_B, subst.ConstantProperties.Liquid_Density_Const_C, subst.ConstantProperties.Liquid_Density_Const_D, subst.ConstantProperties.Liquid_Density_Const_E, T, subst.ConstantProperties.Critical_Temperature)
                            vk(i) = subst.ConstantProperties.Molar_Weight * vk(i)
                        Else
                            vk(i) = Me.m_props.liq_dens_rackett(T, subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Critical_Pressure, subst.ConstantProperties.Acentric_Factor, subst.ConstantProperties.Molar_Weight, subst.ConstantProperties.Z_Rackett, P, Me.AUX_PVAPi(subst.Nome, T))
                        End If
                    Else
                        vk(i) = Me.m_props.liq_dens_rackett(T, subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Critical_Pressure, subst.ConstantProperties.Acentric_Factor, subst.ConstantProperties.Molar_Weight, subst.ConstantProperties.Z_Rackett, P, Me.AUX_PVAPi(subst.Nome, T))
                    End If
                Else
                    vk(i) = Me.m_props.liq_dens_rackett(T, subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Critical_Pressure, subst.ConstantProperties.Acentric_Factor, subst.ConstantProperties.Molar_Weight, subst.ConstantProperties.Z_Rackett, P, Me.AUX_PVAPi(subst.Nome, T))
                End If
                If T > subst.ConstantProperties.Critical_Temperature Then
                    vk(i) = 1.0E+20
                End If
                If Not Double.IsNaN(vk(i)) Then vk(i) = Vx(i) / vk(i) Else vk(i) = 0.0#
                i = i + 1
            Next
            val = 1 / MathEx.Common.Sum(vk)

            m_pr2 = Nothing

            Return val 'kg/m3

        End Function

        Public Overridable Function AUX_LIQDENSi(ByVal subst As Substancia, ByVal T As Double) As Double

            If m_props Is Nothing Then m_props = New DTL.SimulationObjects.PropertyPackages.Auxiliary.PROPS

            Dim val As Double

            If subst.ConstantProperties.LiquidDensityEquation <> "" And subst.ConstantProperties.LiquidDensityEquation <> "0" And Not subst.ConstantProperties.IsIon And Not subst.ConstantProperties.IsSalt Then
                val = Me.CalcCSTDepProp(subst.ConstantProperties.LiquidDensityEquation, subst.ConstantProperties.Liquid_Density_Const_A, subst.ConstantProperties.Liquid_Density_Const_B, subst.ConstantProperties.Liquid_Density_Const_C, subst.ConstantProperties.Liquid_Density_Const_D, subst.ConstantProperties.Liquid_Density_Const_E, T, subst.ConstantProperties.Critical_Temperature)
                val = subst.ConstantProperties.Molar_Weight * val
            Else
                val = Me.m_props.liq_dens_rackett(T, subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Critical_Pressure, subst.ConstantProperties.Acentric_Factor, subst.ConstantProperties.Molar_Weight, subst.ConstantProperties.Z_Rackett, 101325, Me.AUX_PVAPi(subst.Nome, T))
            End If

            Return val 'kg/m3

        End Function

        Public Overridable Function AUX_LIQDENSi(ByVal cprop As ConstantProperties, ByVal T As Double) As Double

            Dim val As Double

            If cprop.LiquidDensityEquation <> "" And cprop.LiquidDensityEquation <> "0" And Not cprop.IsIon And Not cprop.IsSalt Then
                val = Me.CalcCSTDepProp(cprop.LiquidDensityEquation, cprop.Liquid_Density_Const_A, cprop.Liquid_Density_Const_B, cprop.Liquid_Density_Const_C, cprop.Liquid_Density_Const_D, cprop.Liquid_Density_Const_E, T, cprop.Critical_Temperature)
                val = cprop.Molar_Weight * val
            Else
                val = Me.m_props.liq_dens_rackett(T, cprop.Critical_Temperature, cprop.Critical_Pressure, cprop.Acentric_Factor, cprop.Molar_Weight, cprop.Z_Rackett, 101325, Me.AUX_PVAPi(cprop.Name, T))
            End If

            Return val 'kg/m3

        End Function

        Public Function AUX_LIQCPm(ByVal T As Double, ByVal phaseid As Double) As Double

            Dim val As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia

            val = 0
            For Each subst In Me.CurrentMaterialStream.Fases(phaseid).Componentes.Values
                val += subst.FracaoMassica.GetValueOrDefault * Me.AUX_LIQ_Cpi(subst.ConstantProperties, T)
            Next

            Return val

        End Function

        Public Overridable Function AUX_LIQ_Cpi(ByVal cprop As ConstantProperties, ByVal T As Double) As Double

            Dim val As Double

            If cprop.LiquidHeatCapacityEquation <> "" And cprop.LiquidHeatCapacityEquation <> "0" And Not cprop.IsIon And Not cprop.IsSalt Then
                val = Me.CalcCSTDepProp(cprop.LiquidHeatCapacityEquation, cprop.Liquid_Heat_Capacity_Const_A, cprop.Liquid_Heat_Capacity_Const_B, cprop.Liquid_Heat_Capacity_Const_C, cprop.Liquid_Heat_Capacity_Const_D, cprop.Liquid_Heat_Capacity_Const_E, T, cprop.Critical_Temperature)
                val = val / 1000 / cprop.Molar_Weight 'kJ/kg.K
            Else
                'estimate using Rownlinson/Bondi correlation
                val = Me.m_props.Cpl_rb(AUX_CPi(cprop.Name, T), T, cprop.Critical_Temperature, cprop.Acentric_Factor, cprop.Molar_Weight) 'kJ/kg.K
            End If

            Return val

        End Function


        Public MustOverride Function AUX_VAPDENS(ByVal T As Double, ByVal P As Double) As Double

        Public Function AUX_INT_CPDTi(ByVal T1 As Double, ByVal T2 As Double, ByVal subst As String)

            Dim A, B, C, D, E, result As Double
            A = Me.CurrentMaterialStream.Fases(0).Componentes(subst).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_A
            B = Me.CurrentMaterialStream.Fases(0).Componentes(subst).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_B
            C = Me.CurrentMaterialStream.Fases(0).Componentes(subst).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_C
            D = Me.CurrentMaterialStream.Fases(0).Componentes(subst).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_D
            E = Me.CurrentMaterialStream.Fases(0).Componentes(subst).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_E

            result = A * (T2 - T1) + B / 2 * (T2 ^ 2 - T1 ^ 2) + C / 3 * (T2 ^ 3 - T1 ^ 3) + D / 4 * (T2 ^ 4 - T1 ^ 4) + E / 5 * (T2 ^ 5 - T1 ^ 5)

            Dim deltaT As Double = (T2 - T1) / 10
            Dim result2, Ti As Double
            Dim i As Integer = 0
            Dim integral As Double = 0

            Ti = T1 + deltaT
            For i = 0 To 9
                integral += Me.AUX_CPi(subst, Ti) * deltaT
                Ti += deltaT
            Next

            'result = Me.IntegralSimpsonCp(T1, T2, 0.001, subst)

            result = result / Me.CurrentMaterialStream.Fases(0).Componentes(subst).ConstantProperties.Molar_Weight

            result2 = integral '/ Me.CurrentMaterialStream.Fases(0).Componentes(subst).ConstantProperties.Molar_Weight

            Return result2

        End Function

        Public Function AUX_INT_CPDT_Ti(ByVal T1 As Double, ByVal T2 As Double, ByVal subst As String)

            Dim A, B, C, D, E, result As Double
            A = Me.CurrentMaterialStream.Fases(0).Componentes(subst).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_A
            B = Me.CurrentMaterialStream.Fases(0).Componentes(subst).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_B
            C = Me.CurrentMaterialStream.Fases(0).Componentes(subst).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_C
            D = Me.CurrentMaterialStream.Fases(0).Componentes(subst).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_D
            E = Me.CurrentMaterialStream.Fases(0).Componentes(subst).ConstantProperties.Ideal_Gas_Heat_Capacity_Const_E

            result = A * Log(T2 / T1) + B * (T2 - T1) + C / 2 * (T2 ^ 2 - T1 ^ 2) + D / 3 * (T2 ^ 3 - T1 ^ 3) + E / 4 * (T2 ^ 4 - T1 ^ 4)

            Dim deltaT As Double = (T2 - T1) / 10
            Dim result2, Ti As Double
            Dim i As Integer = 0
            Dim integral As Double = 0

            Ti = T1 + deltaT
            For i = 0 To 9
                integral += Me.AUX_CPi(subst, Ti) * deltaT / (Ti - deltaT) '* Log(Ti / (Ti - deltaT))
                Ti += deltaT
            Next

            'result = Me.IntegralSimpsonCp(T1, T2, 0.001, subst)

            result = result / Me.CurrentMaterialStream.Fases(0).Componentes(subst).ConstantProperties.Molar_Weight

            result2 = integral '/ Me.CurrentMaterialStream.Fases(0).Componentes(subst).ConstantProperties.Molar_Weight

            Return result2

        End Function

        Public Function AUX_INT_CPDTm(ByVal T1 As Double, ByVal T2 As Double, ByVal fase As Fase)

            Dim val As Double = 0
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia

            Select Case fase

                Case DTL.SimulationObjects.PropertyPackages.Fase.Mixture

                    For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                        If subst.FracaoMolar.GetValueOrDefault <> 0.0# And subst.FracaoMassica.GetValueOrDefault = 0.0# Then
                            subst.FracaoMassica = Me.AUX_CONVERT_MOL_TO_MASS(subst.Nome, Me.RET_PHASEID(fase))
                        End If
                        val += subst.FracaoMassica.GetValueOrDefault * Me.AUX_INT_CPDTi(T1, T2, subst.Nome)
                    Next

                Case DTL.SimulationObjects.PropertyPackages.Fase.Liquid

                    For Each subst In Me.CurrentMaterialStream.Fases(1).Componentes.Values
                        If subst.FracaoMolar.GetValueOrDefault <> 0.0# And subst.FracaoMassica.GetValueOrDefault = 0.0# Then
                            subst.FracaoMassica = Me.AUX_CONVERT_MOL_TO_MASS(subst.Nome, Me.RET_PHASEID(fase))
                        End If
                        val += subst.FracaoMassica.GetValueOrDefault * Me.AUX_INT_CPDTi(T1, T2, subst.Nome)
                    Next

                Case DTL.SimulationObjects.PropertyPackages.Fase.Liquid1

                    For Each subst In Me.CurrentMaterialStream.Fases(3).Componentes.Values
                        If subst.FracaoMolar.GetValueOrDefault <> 0.0# And subst.FracaoMassica.GetValueOrDefault = 0.0# Then
                            subst.FracaoMassica = Me.AUX_CONVERT_MOL_TO_MASS(subst.Nome, Me.RET_PHASEID(fase))
                        End If
                        val += subst.FracaoMassica.GetValueOrDefault * Me.AUX_INT_CPDTi(T1, T2, subst.Nome)
                    Next

                Case DTL.SimulationObjects.PropertyPackages.Fase.Liquid2

                    For Each subst In Me.CurrentMaterialStream.Fases(4).Componentes.Values
                        If subst.FracaoMolar.GetValueOrDefault <> 0.0# And subst.FracaoMassica.GetValueOrDefault = 0.0# Then
                            subst.FracaoMassica = Me.AUX_CONVERT_MOL_TO_MASS(subst.Nome, Me.RET_PHASEID(fase))
                        End If
                        val += subst.FracaoMassica.GetValueOrDefault * Me.AUX_INT_CPDTi(T1, T2, subst.Nome)
                    Next

                Case DTL.SimulationObjects.PropertyPackages.Fase.Liquid3

                    For Each subst In Me.CurrentMaterialStream.Fases(5).Componentes.Values
                        If subst.FracaoMolar.GetValueOrDefault <> 0.0# And subst.FracaoMassica.GetValueOrDefault = 0.0# Then
                            subst.FracaoMassica = Me.AUX_CONVERT_MOL_TO_MASS(subst.Nome, Me.RET_PHASEID(fase))
                        End If
                        val += subst.FracaoMassica.GetValueOrDefault * Me.AUX_INT_CPDTi(T1, T2, subst.Nome)
                    Next

                Case DTL.SimulationObjects.PropertyPackages.Fase.Vapor

                    For Each subst In Me.CurrentMaterialStream.Fases(2).Componentes.Values
                        If subst.FracaoMolar.GetValueOrDefault <> 0.0# And subst.FracaoMassica.GetValueOrDefault = 0.0# Then
                            subst.FracaoMassica = Me.AUX_CONVERT_MOL_TO_MASS(subst.Nome, Me.RET_PHASEID(fase))
                        End If
                        val += subst.FracaoMassica.GetValueOrDefault * Me.AUX_INT_CPDTi(T1, T2, subst.Nome)
                    Next

            End Select

            Return val

        End Function

        Public Function AUX_INT_CPDTi_L(ByVal T1 As Double, ByVal T2 As Double, ByVal subst As String) As Double

            Dim deltaT As Double = (T2 - T1) / 10
            Dim Ti, Tc As Double
            Dim i As Integer = 0
            Dim integral As Double = 0

            Ti = T1 + deltaT / 2
            Tc = Me.CurrentMaterialStream.Fases(0).Componentes(subst).ConstantProperties.Critical_Temperature
            For i = 0 To 9
                If Ti > Tc Then
                    integral += Me.AUX_CPi(subst, Ti) * deltaT
                Else
                    integral += Me.AUX_LIQ_Cpi(Me.CurrentMaterialStream.Fases(0).Componentes(subst).ConstantProperties, Ti) * deltaT
                End If

                Ti += deltaT
            Next

            Return integral 'KJ/Kg

        End Function

        Public Function AUX_INT_CPDTm_L(ByVal T1 As Double, ByVal T2 As Double, ByVal Vw As Object)

            Dim val As Double
            Dim i As Integer = 0
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                val += Vw(i) * Me.AUX_INT_CPDTi_L(T1, T2, subst.Nome)
                i += 1
            Next

            Return val

        End Function

        Public Function AUX_INT_CPDT_Tm(ByVal T1 As Double, ByVal T2 As Double, ByVal fase As Fase)

            Dim val As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia

            For Each subst In Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).Componentes.Values
                If subst.FracaoMolar.GetValueOrDefault <> 0.0# And subst.FracaoMassica.GetValueOrDefault = 0.0# Then
                    subst.FracaoMassica = Me.AUX_CONVERT_MOL_TO_MASS(subst.Nome, Me.RET_PHASEID(fase))
                End If
                val += subst.FracaoMassica.GetValueOrDefault * Me.AUX_INT_CPDT_Ti(T1, T2, subst.Nome)
            Next

            Return val

        End Function

        Public Function AUX_DELGF_T(ByVal T1 As Double, ByVal T2 As Double, ByVal id As String, Optional ByVal mode2 As Boolean = False) As Double

            Dim dA As Double = 0
            Dim dB As Double = 0
            Dim dC As Double = 0
            Dim dD As Double = 0
            Dim dE As Double = 0
            Dim dHr As Double = 0
            Dim dGr As Double = 0
            Dim term3 As Double = 0

            Dim int1, int2 As Double
            Dim R = 8.314

            With Me.CurrentMaterialStream.Fases(0).Componentes(id).ConstantProperties

                dHr = .IG_Enthalpy_of_Formation_25C
                dGr = .IG_Gibbs_Energy_of_Formation_25C

                If mode2 Then
                    If .IsIon Or .IsSalt Then
                        term3 = .Electrolyte_Cp0 * 1000 / .Molar_Weight * (Log(T2 / T1) + (T1 / T2) - 1) / R
                    Else
                        int1 = Me.AUX_INT_CPDTi(T1, T2, id)
                        int2 = Me.AUX_INT_CPDT_Ti(T1, T2, id)
                        term3 = int1 / (R * T2) - int2 / R
                    End If
                Else
                    int1 = Me.AUX_INT_CPDTi(T1, T2, id)
                    int2 = Me.AUX_INT_CPDT_Ti(T1, T2, id)
                    term3 = int1 / (R * T2) - int2 / R
                End If


            End With

            Dim result As Double

            If mode2 Then
                result = dGr / (R * T1) + dHr / R * (1 / T2 - 1 / T1) + term3
            Else
                result = (dGr - dHr) / (R * T1) + dHr / (R * T2) + term3
            End If

            Return result

        End Function

        Public Function AUX_DELGig_RT(ByVal T1 As Double, ByVal T2 As Double, ByVal ID() As String, ByVal stcoeff() As Double, ByVal baseID As Integer, Optional ByVal mode2 As Boolean = False) As Double

            Dim n As Integer = ID.Length

            Dim int1(n - 1), int2(n - 1), sint1, sint2, dgfT(n - 1), sumdgft As Double

            Dim dHr As Double = 0
            Dim dGr As Double = 0

            sint1 = 0
            sint2 = 0

            With Me.CurrentMaterialStream

                Dim i As Integer = 0
                sumdgft = 0
                Do
                    dHr += stcoeff(i) * .Fases(0).Componentes(ID(i)).ConstantProperties.IG_Enthalpy_of_Formation_25C * .Fases(0).Componentes(ID(i)).ConstantProperties.Molar_Weight
                    dGr += stcoeff(i) * .Fases(0).Componentes(ID(i)).ConstantProperties.IG_Gibbs_Energy_of_Formation_25C * .Fases(0).Componentes(ID(i)).ConstantProperties.Molar_Weight
                    int1(i) = stcoeff(i) * Me.AUX_INT_CPDTi(T1, T2, ID(i)) / stcoeff(baseID) * .Fases(0).Componentes(ID(i)).ConstantProperties.Molar_Weight
                    int2(i) = stcoeff(i) * Me.AUX_INT_CPDT_Ti(T1, T2, ID(i)) / stcoeff(baseID) * .Fases(0).Componentes(ID(i)).ConstantProperties.Molar_Weight
                    sint1 += int1(i)
                    sint2 += int2(i)
                    dgfT(i) = stcoeff(i) * Me.AUX_DELGF_T(T1, T2, ID(i), mode2) * .Fases(0).Componentes(ID(i)).ConstantProperties.Molar_Weight
                    sumdgft += dgfT(i) '/ stcoeff(baseID)
                    i = i + 1
                Loop Until i = n

            End With
            dHr /= Abs(stcoeff(baseID))
            dGr /= Abs(stcoeff(baseID))

            Dim R = 8.314

            Dim result = (dGr - dHr) / (R * T1) + dHr / (R * T2) + sint1 / (R * T2) - sint2 / R

            Return sumdgft

        End Function

        Public Function AUX_DELHig_RT(ByVal T1 As Double, ByVal T2 As Double, ByVal ID() As String, ByVal stcoeff() As Double, ByVal baseID As Integer) As Double

            Dim n As Integer = ID.Length

            Dim int1(n - 1), sint1 As Double

            Dim dHr As Double = 0

            sint1 = 0

            With Me.CurrentMaterialStream

                Dim i As Integer = 0
                Do
                    dHr += stcoeff(i) * .Fases(0).Componentes(ID(i)).ConstantProperties.IG_Enthalpy_of_Formation_25C * .Fases(0).Componentes(ID(i)).ConstantProperties.Molar_Weight
                    int1(i) = stcoeff(i) * Me.AUX_INT_CPDTi(T1, T2, ID(i)) / Abs(stcoeff(baseID)) * .Fases(0).Componentes(ID(i)).ConstantProperties.Molar_Weight
                    sint1 += int1(i)
                    i = i + 1
                Loop Until i = n

            End With
            dHr /= Abs(stcoeff(baseID))

            Dim result = dHr + sint1

            Return result / (8.314 * T2)

        End Function

        Public Function AUX_CONVERT_MOL_TO_MASS(ByVal subst As String, ByVal phasenumber As Integer) As Double

            Dim mol_x_mm As Double
            Dim sub1 As DTL.ClassesBasicasTermodinamica.Substancia
            For Each sub1 In Me.CurrentMaterialStream.Fases(phasenumber).Componentes.Values
                mol_x_mm += sub1.FracaoMolar.GetValueOrDefault * sub1.ConstantProperties.Molar_Weight
            Next

            sub1 = Me.CurrentMaterialStream.Fases(phasenumber).Componentes(subst)
            If mol_x_mm <> 0.0# Then
                Return sub1.FracaoMolar.GetValueOrDefault * sub1.ConstantProperties.Molar_Weight / mol_x_mm
            Else
                Return 0.0#
            End If

        End Function

        Public Function AUX_CONVERT_MASS_TO_MOL(ByVal subst As String, ByVal phasenumber As Integer) As Double

            Dim mass_div_mm As Double
            Dim sub1 As DTL.ClassesBasicasTermodinamica.Substancia
            For Each sub1 In Me.CurrentMaterialStream.Fases(phasenumber).Componentes.Values
                mass_div_mm += sub1.FracaoMassica.GetValueOrDefault / sub1.ConstantProperties.Molar_Weight
            Next

            sub1 = Me.CurrentMaterialStream.Fases(phasenumber).Componentes(subst)
            Return sub1.FracaoMassica.GetValueOrDefault / sub1.ConstantProperties.Molar_Weight / mass_div_mm

        End Function

        Public Overridable Function DW_CalcSolidEnthalpy(ByVal T As Double, ByVal Vx As Double(), cprops As List(Of ConstantProperties)) As Double

            Dim n As Integer = UBound(Vx)
            Dim i As Integer
            Dim HS As Double = 0.0#
            Dim Cpi As Double

            For i = 0 To n
                If cprops(i).OriginalDB = "ChemSep" Or cprops(i).OriginalDB = "User" Then
                    Dim A, B, C, D, E, result As Double
                    Dim eqno As String = cprops(i).SolidHeatCapacityEquation
                    Dim mw As Double = cprops(i).Molar_Weight
                    A = cprops(i).Solid_Heat_Capacity_Const_A
                    B = cprops(i).Solid_Heat_Capacity_Const_B
                    C = cprops(i).Solid_Heat_Capacity_Const_C
                    D = cprops(i).Solid_Heat_Capacity_Const_D
                    E = cprops(i).Solid_Heat_Capacity_Const_E
                    '<SolidHeatCapacityCp name="Solid heat capacity"  units="J/kmol/K" >
                    result = Me.CalcCSTDepProp(eqno, A, B, C, D, E, T, 0) / 1000 / mw 'kJ/kg.K
                    Cpi = result
                Else
                    Cpi = 0.0#
                End If
                HS += Vx(i) * Cpi * (T - 298)
            Next

            Return HS 'kJ/kg

        End Function

        Public Overridable Function DW_CalcSolidHeatCapacityCp(ByVal T As Double, ByVal Vx As Double(), cprops As List(Of ConstantProperties)) As Double

            Dim n As Integer = UBound(Vx)
            Dim i As Integer
            Dim Cp As Double = 0.0#
            Dim Cpi As Double

            For i = 0 To n
                If cprops(i).OriginalDB = "ChemSep" Or cprops(i).OriginalDB = "User" Then
                    Dim A, B, C, D, E, result As Double
                    Dim eqno As String = cprops(i).SolidHeatCapacityEquation
                    Dim mw As Double = cprops(i).Molar_Weight
                    A = cprops(i).Solid_Heat_Capacity_Const_A
                    B = cprops(i).Solid_Heat_Capacity_Const_B
                    C = cprops(i).Solid_Heat_Capacity_Const_C
                    D = cprops(i).Solid_Heat_Capacity_Const_D
                    E = cprops(i).Solid_Heat_Capacity_Const_E
                    '<SolidHeatCapacityCp name="Solid heat capacity"  units="J/kmol/K" >
                    result = Me.CalcCSTDepProp(eqno, A, B, C, D, E, T, 0) / 1000 / mw 'kJ/kg.K
                    Cpi = result
                Else
                    Cpi = 0.0#
                End If
                Cp += Vx(i) * Cpi
            Next

            Return Cp 'kJ/kg.K

        End Function

        Public Function RET_VMOL(ByVal fase As Fase) As Double()

            Dim val(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim i As Integer = 0

            For Each subst In Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).Componentes.Values
                val(i) = subst.FracaoMolar.GetValueOrDefault
                i += 1
            Next

            Return val

        End Function

        Public Function RET_VMM()

            Dim val(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim i As Integer = 0

            For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                val(i) = subst.ConstantProperties.Molar_Weight
                i += 1
            Next

            Return val

        End Function

        Public Function RET_VMAS(ByVal fase As Fase) As Double()

            Dim val(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim i As Integer = 0
            Dim sum As Double = 0.0#

            For Each subst In Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).Componentes.Values
                val(i) = subst.FracaoMassica.GetValueOrDefault
                sum += val(i)
                i += 1
            Next

            If sum <> 0.0# Then
                Return val
            Else
                Return AUX_CONVERT_MOL_TO_MASS(RET_VMOL(fase))
            End If

        End Function

        Public Function RET_VTC()

            Dim val(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim i As Integer = 0

            For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                val(i) = subst.ConstantProperties.Critical_Temperature
                i += 1
            Next

            Return val

        End Function

        Public Function RET_VTF()

            Dim val(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim i As Integer = 0

            For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                val(i) = subst.ConstantProperties.TemperatureOfFusion
                i += 1
            Next

            Return val

        End Function

        Public Function RET_VTB()

            Dim val(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim i As Integer = 0

            For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                val(i) = subst.ConstantProperties.Normal_Boiling_Point
                i += 1
            Next

            Return val

        End Function

        Public Function RET_VPC()

            Dim val(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim i As Integer = 0


            For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                val(i) = subst.ConstantProperties.Critical_Pressure
                i += 1
            Next

            Return val

        End Function

        Public Function RET_VZC()

            Dim val(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim i As Integer = 0

            For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                val(i) = subst.ConstantProperties.Critical_Compressibility
                i += 1
            Next

            Return val

        End Function

        Public Function RET_VZRa()

            Dim val(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim i As Integer = 0

            For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                val(i) = subst.ConstantProperties.Z_Rackett
                i += 1
            Next
            Return val

        End Function

        Public Function RET_VVC()

            Dim vc, val(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim i As Integer = 0

            For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                vc = subst.ConstantProperties.Critical_Volume
                If subst.ConstantProperties.Critical_Volume = 0.0# Then
                    vc = m_props.Vc(subst.ConstantProperties.Critical_Temperature, subst.ConstantProperties.Critical_Pressure, subst.ConstantProperties.Acentric_Factor)
                End If
                val(i) = vc
                i += 1
            Next

            Return val

        End Function

        Public Function RET_VW()

            Dim val(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim i As Integer = 0

            For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                val(i) = subst.ConstantProperties.Acentric_Factor
                i += 1
            Next

            Return val

        End Function

        Public Function RET_VCP(ByVal T As Double)

            Dim val(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim i As Integer = 0

            For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                val(i) = Me.AUX_CPi(subst.Nome, T)
                i += 1
            Next

            Return val

        End Function

        Public Function RET_VHVAP(ByVal T As Double) As Array

            Dim val(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim i As Integer = 0

            For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                val(i) = Me.AUX_HVAPi(subst.Nome, T)
                i += 1
            Next

            Return val

        End Function


        Public Function RET_HVAPM(ByVal Vxw As Array, ByVal T As Double) As Double

            Dim val As Double = 0.0#
            Dim i As Integer
            Dim n As Integer = Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1

            i = 0
            For Each subst As Substancia In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                val += Vxw(i) * Me.AUX_HVAPi(subst.Nome, T)
                i += 1
            Next

            Return val

        End Function

        Public Function AUX_HVAPi(ByVal sub1 As String, ByVal T As Double)

            Dim A, B, C, D, E, Tr, result As Double
            A = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.HVap_A
            B = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.HVap_B
            C = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.HVap_C
            D = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.HVap_D
            E = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.HVap_E

            Tr = T / Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Critical_Temperature

            If Tr >= 1 Then Return 0.0#

            If Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.OriginalDB = "DWSIM" Or _
                                Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.OriginalDB = "" Then
                If Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.IsHYPO = 1 Or _
                Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.IsPF = 1 Then
                    Dim tr1 As Double
                    tr1 = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Normal_Boiling_Point / Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Critical_Temperature
                    result = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.HVap_A * ((1 - Tr) / (1 - tr1)) ^ 0.375
                    Return result 'kJ/kg
                Else
                    result = A * (1 - Tr) ^ (B + C * Tr + D * Tr ^ 2)
                    Return result / Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Molar_Weight / 1000 'kJ/kg
                End If
            ElseIf Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.OriginalDB = "ChemSep" Then
                Dim eqno As String = Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.VaporizationEnthalpyEquation
                result = Me.CalcCSTDepProp(eqno, A, B, C, D, E, T, T / Tr) / Me.CurrentMaterialStream.Fases(0).Componentes(sub1).ConstantProperties.Molar_Weight / 1000 'kJ/kg
                Return result
            Else
                Return 0.0#
            End If

        End Function

        Public Function RET_Hid(ByVal T1 As Double, ByVal T2 As Double, ByVal fase As Fase) As Double

            Dim val As Double
            'Dim subst As DTL.ClassesBasicasTermodinamica.Substancia

            Dim phaseID As Integer

            If fase = PropertyPackages.Fase.Liquid Then phaseID = 1
            If fase = PropertyPackages.Fase.Liquid1 Then phaseID = 3
            If fase = PropertyPackages.Fase.Liquid2 Then phaseID = 4
            If fase = PropertyPackages.Fase.Liquid3 Then phaseID = 5
            If fase = PropertyPackages.Fase.Vapor Then phaseID = 2
            If fase = PropertyPackages.Fase.Mixture Then phaseID = 0

            'For Each subst In Me.CurrentMaterialStream.Fases(phaseID).Componentes.Values
            '    val += subst.FracaoMassica.GetValueOrDefault * subst.ConstantProperties.Enthalpy_of_Formation_25C
            'Next

            Return Me.AUX_INT_CPDTm(T1, T2, fase) + val

        End Function


        Public Function RET_Hid(ByVal T1 As Double, ByVal T2 As Double, ByVal Vz As Object) As Double

            Return Me.AUX_INT_CPDTm(T1, T2, Me.AUX_CONVERT_MOL_TO_MASS(Vz))

        End Function

        Public Function RET_Hid_L(ByVal T1 As Double, ByVal T2 As Double, ByVal Vz As Object) As Double

            Return Me.AUX_INT_CPDTm_L(T1, T2, Me.AUX_CONVERT_MOL_TO_MASS(Vz))

        End Function

        Public Function RET_Sid_L(ByVal T1 As Double, ByVal T2 As Double, ByVal Vz As Object) As Double

            Return Me.RET_Hid_L(T1, T2, Vz) / T2

        End Function

        Public Function RET_Hid_i(ByVal T1 As Double, ByVal T2 As Double, ByVal id As String) As Double

            Return Me.AUX_INT_CPDTi(T1, T2, id)

        End Function

        Public Function RET_Hid_i_L(ByVal T1 As Double, ByVal T2 As Double, ByVal id As String) As Double

            Return Me.AUX_INT_CPDTi_L(T1, T2, id)

        End Function

        Public Function RET_Sid(ByVal T1 As Double, ByVal T2 As Double, ByVal P2 As Double, ByVal fase As Fase) As Double

            Dim val As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia

            Dim phaseID As Integer

            If fase = PropertyPackages.Fase.Liquid Then phaseID = 1
            If fase = PropertyPackages.Fase.Liquid1 Then phaseID = 3
            If fase = PropertyPackages.Fase.Liquid2 Then phaseID = 4
            If fase = PropertyPackages.Fase.Liquid3 Then phaseID = 5
            If fase = PropertyPackages.Fase.Vapor Then phaseID = 2
            If fase = PropertyPackages.Fase.Mixture Then phaseID = 0
            'subst.FracaoMassica.GetValueOrDefault * subst.ConstantProperties.Standard_Absolute_Entropy 
            For Each subst In Me.CurrentMaterialStream.Fases(phaseID).Componentes.Values
                If subst.FracaoMolar.GetValueOrDefault <> 0 Then val += -8.314 * subst.FracaoMolar.GetValueOrDefault * Math.Log(subst.FracaoMolar.GetValueOrDefault) / subst.ConstantProperties.Molar_Weight
            Next

            Dim tmp = Me.AUX_INT_CPDT_Tm(T1, T2, fase) + val - 8.314 * Math.Log(P2 / 101325) / Me.AUX_MMM(fase)

            Return tmp

        End Function

        Public Function RET_Sid(ByVal T1 As Double, ByVal T2 As Double, ByVal P2 As Double, ByVal Vz As Object) As Double

            Dim val As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim Vw = Me.AUX_CONVERT_MOL_TO_MASS(Vz)

            Dim i As Integer = 0

            'Vw(i) * subst.ConstantProperties.Standard_Absolute_Entropy
            val = 0
            For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                If Vz(i) <> 0 Then val += -8.314 * Vz(i) * Math.Log(Vz(i)) / subst.ConstantProperties.Molar_Weight
                i = i + 1
            Next
            Dim tmp1 = 8.314 * Math.Log(P2 / 101325) / Me.AUX_MMM(Vz)
            Dim tmp2 = Me.AUX_INT_CPDT_Tm(T1, T2, Me.AUX_CONVERT_MOL_TO_MASS(Vz))
            Return tmp2 + val - tmp1

        End Function

        Public Function RET_Sid_i(ByVal T1 As Double, ByVal T2 As Double, ByVal P2 As Double, ByVal id As String) As Double

            Dim val As Double

            Dim tmp1 = 8.314 * Math.Log(P2 / 101325) / Me.CurrentMaterialStream.Fases(0).Componentes(id).ConstantProperties.Molar_Weight
            Dim tmp2 = Me.AUX_INT_CPDT_Ti(T1, T2, id)
            Return tmp2 + val - tmp1

        End Function

        Public Function RET_Gid(ByVal T1 As Double, ByVal T2 As Double, ByVal P2 As Double, ByVal Vz As Object) As Double

            Dim hid = Me.RET_Hid(T1, T2, Vz)
            Dim sid = Me.RET_Sid(T1, T2, P2, Vz)

            Return hid - T2 * sid

        End Function

        Public Function RET_Gid_i(ByVal T1 As Double, ByVal T2 As Double, ByVal P2 As Double, ByVal id As String) As Double

            Dim hid = Me.RET_Hid_i(T1, T2, id)
            Dim sid = Me.RET_Sid_i(T1, T2, P2, id)

            Return hid - T2 * sid

        End Function

        Public Function RET_Gid(ByVal T1 As Double, ByVal T2 As Double, ByVal P2 As Double, ByVal fase As Fase) As Double

            Dim hid = Me.RET_Hid(T1, T2, fase)
            Dim sid = Me.RET_Sid(T1, T2, P2, fase)

            Return hid - T2 * sid

        End Function

        Public Overridable Function RET_VKij() As Double(,)

            If Me.ParametrosDeInteracao Is Nothing Then
                Me.ParametrosDeInteracao = New DataTable
                With Me.ParametrosDeInteracao.Columns
                    For Each subst As DTL.ClassesBasicasTermodinamica.Substancia In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                        .Add(subst.ConstantProperties.ID, GetType(System.Double))
                    Next
                End With

                With Me.ParametrosDeInteracao.Rows
                    For Each subst As DTL.ClassesBasicasTermodinamica.Substancia In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                        .Add()
                    Next
                End With
            End If

            Dim val(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1, Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As Double
            Dim i As Integer = 0
            Dim l As Integer = 0
            For Each r As DataRow In Me.ParametrosDeInteracao.Rows
                i = 0
                Do
                    If l <> i And r.Item(i).GetType().ToString <> "System.DBNull" Then
                        Dim value As Double
                        If Double.TryParse(r.Item(i), value) Then
                            val(l, i) = r.Item(i)
                        Else
                            val(l, i) = 0
                        End If
                    Else
                        val(l, i) = 0
                    End If
                    i = i + 1
                Loop Until i = Me.ParametrosDeInteracao.Rows.Count
                l = l + 1
            Next

            Return val

        End Function

        Public Function RET_VCSACIDS() As String()

            Dim val(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As String
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim i As Integer = 0

            For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                val(i) = subst.ConstantProperties.COSMODBName
                If val(i) = "" Then val(i) = subst.ConstantProperties.CAS_Number
                i += 1
            Next

            Return val

        End Function

        Public Function RET_VIDS() As String()

            Dim val(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As String
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim i As Integer = 0

            For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                val(i) = subst.ConstantProperties.ID
                i += 1
            Next

            Return val

        End Function

        Public Function RET_VCAS() As String()

            Dim val(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As String
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim i As Integer = 0

            For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                val(i) = subst.ConstantProperties.CAS_Number
                i += 1
            Next

            Return val

        End Function

        Public Function RET_VNAMES() As String()

            Dim val(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As String
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim i As Integer = 0

            For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                val(i) = subst.ConstantProperties.Name
                i += 1
            Next

            Return val

        End Function

        Public Function RET_NullVector() As Double()

            Dim val(Me.CurrentMaterialStream.Fases(0).Componentes.Count - 1) As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            Dim i As Integer = 0
            For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                val(i) = 0.0#
                i += 1
            Next
            Return val

        End Function

        Public Function RET_PHASECODE(ByVal phaseID As Integer) As SimulationObjects.PropertyPackages.Fase
            Select Case phaseID
                Case 0
                    Return Fase.Mixture
                Case 1
                    Return Fase.Liquid
                Case 2
                    Return Fase.Vapor
                Case 3
                    Return Fase.Liquid1
                Case 4
                    Return Fase.Liquid2
                Case 5
                    Return Fase.Liquid3
                Case 6
                    Return Fase.Aqueous
                Case 7
                    Return Fase.Solid
            End Select
        End Function

        Public Function RET_PHASEID(ByVal phasecode As SimulationObjects.PropertyPackages.Fase) As Integer
            Select Case phasecode
                Case Fase.Mixture
                    Return 0
                Case Fase.Liquid
                    Return 1
                Case Fase.Vapor
                    Return 2
                Case Fase.Liquid1
                    Return 3
                Case Fase.Liquid2
                    Return 4
                Case Fase.Liquid3
                    Return 5
                Case Fase.Aqueous
                    Return 6
                Case Fase.Solid
                    Return 7
                Case Else
                    Return -1
            End Select
        End Function

        Public Function AUX_MMM(ByVal Vz As Object) As Double

            Dim val As Double
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia

            Dim i As Integer = 0

            For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                val += Vz(i) * subst.ConstantProperties.Molar_Weight
                i += 1
            Next

            Return val

        End Function

        Public Function AUX_NORMALIZE(ByVal Vx As Object) As Double()

            Dim sum As Double = 0
            Dim i, n As Integer

            n = UBound(Vx)

            Dim Vxnew(n) As Double

            For i = 0 To n
                sum += Vx(i)
            Next

            For i = 0 To n
                Vxnew(i) = Vx(i) / sum
            Next

            Return Vxnew

        End Function

        Public Function AUX_ERASE(ByVal Vx As Object) As Double()

            Dim i, n As Integer

            n = UBound(Vx)

            Dim Vx2(n) As Double

            For i = 0 To n
                Vx2(i) = 0
            Next

            Return Vx2

        End Function

        Public Function AUX_SINGLECOMPIDX(ByVal Vx As Object) As Integer

            Dim i, n As Integer

            n = UBound(Vx)

            For i = 0 To n
                If Vx(i) <> 0 Then Return i
            Next

            Return -1

        End Function

        Public Function AUX_SINGLECOMPIDX(ByVal fase As Fase) As Integer

            Dim i As Integer

            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia

            i = 0
            For Each subst In Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).Componentes.Values
                If subst.FracaoMolar.GetValueOrDefault <> 0 Then Return i
            Next

            Return -1

        End Function

        Public Function AUX_IS_SINGLECOMP(ByVal Vx As Object) As Boolean

            Dim i, c, n As Integer

            n = UBound(Vx)

            c = 0
            For i = 0 To n
                If Vx(i) <> 0 Then c += 1
            Next

            If c = 1 Then Return True Else Return False

        End Function

        Public Function AUX_IS_SINGLECOMP(ByVal fase As Fase) As Boolean

            Dim c As Integer

            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia

            c = 0
            For Each subst In Me.CurrentMaterialStream.Fases(Me.RET_PHASEID(fase)).Componentes.Values
                If subst.FracaoMolar.GetValueOrDefault <> 0 Then c += 1
            Next

            If c = 1 Then Return True Else Return False

        End Function

        Public Function AUX_INT_CPDTm(ByVal T1 As Double, ByVal T2 As Double, ByVal Vw As Object)

            Dim val As Double
            Dim i As Integer = 0
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                val += Vw(i) * Me.AUX_INT_CPDTi(T1, T2, subst.Nome)
                i += 1
            Next

            Return val

        End Function

        Public Function AUX_INT_CPDT_Tm(ByVal T1 As Double, ByVal T2 As Double, ByVal Vw As Object)

            Dim val As Double
            Dim i As Integer = 0
            Dim subst As DTL.ClassesBasicasTermodinamica.Substancia
            For Each subst In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                val += Vw(i) * Me.AUX_INT_CPDT_Ti(T1, T2, subst.Nome)
                i += 1
            Next

            Return val

        End Function

        Public Function AUX_CONVERT_MOL_TO_MASS(ByVal Vz As Object) As Double()

            Dim Vwe(UBound(Vz)) As Double
            Dim mol_x_mm As Double = 0
            Dim i As Integer = 0
            Dim sub1 As DTL.ClassesBasicasTermodinamica.Substancia
            For Each sub1 In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                mol_x_mm += Vz(i) * sub1.ConstantProperties.Molar_Weight
                i += 1
            Next

            i = 0
            For Each sub1 In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                If mol_x_mm <> 0 Then
                    Vwe(i) = Vz(i) * sub1.ConstantProperties.Molar_Weight / mol_x_mm
                Else
                    Vwe(i) = 0.0#
                End If
                i += 1
            Next

            Return Vwe

        End Function

        Public Function AUX_CONVERT_MASS_TO_MOL(ByVal Vz As Object) As Double()

            Dim Vw(UBound(Vz)) As Double
            Dim mass_div_mm As Double
            Dim i As Integer = 0
            Dim sub1 As DTL.ClassesBasicasTermodinamica.Substancia
            For Each sub1 In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                mass_div_mm += Vz(i) / sub1.ConstantProperties.Molar_Weight
                i += 1
            Next

            i = 0
            For Each sub1 In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                Vw(i) = Vz(i) / sub1.ConstantProperties.Molar_Weight / mass_div_mm
                i += 1
            Next

            Return Vw

        End Function

        Public Function AUX_CalculateSumSquares(ByVal Vz As Object) As Double

            Dim n, i As Integer
            n = UBound(Vz)
            Dim sum As Double = 0.0#

            For i = 0 To n
                sum += Vz(i) ^ 2
            Next

            Return sum

        End Function

        Public Function AUX_CalculateAbsSumSquares(ByVal Vz As Object) As Double

            Dim n, i As Integer
            n = UBound(Vz)
            Dim sum As Double = 0.0#

            For i = 0 To n
                sum += Abs(Vz(i)) ^ 2
            Next

            Return sum

        End Function

        Public Function AUX_CalculateSum(ByVal Vz As Object) As Double

            Dim n, i As Integer
            n = UBound(Vz)
            Dim sum As Double = 0.0#

            For i = 0 To n
                sum += Vz(i)
            Next

            Return sum

        End Function

        Public Function AUX_CalculateAbsSum(ByVal Vz As Object) As Double

            Dim n, i As Integer
            n = UBound(Vz)
            Dim sum As Double = 0.0#

            For i = 0 To n
                sum += Abs(Vz(i))
            Next

            Return sum

        End Function

        Public Function AUX_CheckTrivial(ByVal KI As Object) As Boolean

            Dim isTrivial As Boolean = True
            Dim n, i As Integer
            n = UBound(KI)

            For i = 0 To n
                If Abs(KI(i) - 1) > 0.0001 Then isTrivial = False
            Next

            Return isTrivial

        End Function

        Function CalcCSTDepProp(ByVal eqno As String, ByVal A As Double, ByVal B As Double, ByVal C As Double, ByVal D As Double, ByVal E As Double, ByVal T As Double, ByVal Tc As Double) As Double

            Dim Tr As Double = T / Tc

            Select Case eqno
                Case "1"
                    Return A
                Case "2"
                    Return A + B * T
                Case "3"
                    Return A + B * T + C * T ^ 2
                Case "4"
                    Return A + B * T + C * T ^ 2 + D * T ^ 3
                Case "5"
                    Return A + B * T + C * T ^ 2 + D * T ^ 3 + E * T ^ 4
                Case "10"
                    Return Exp(A - B / (T + C))
                Case "11"
                    Return Exp(A)
                Case "12"
                    Return Exp(A + B * T)
                Case "13"
                    Return Exp(A + B * T + C * T ^ 2)
                Case "14"
                    Return Exp(A + B * T + C * T ^ 2 + D * T ^ 3)
                Case "15"
                    Return Exp(A + B * T + C * T ^ 2 + D * T ^ 3 + E * T ^ 4)
                Case "16"
                    Return A + Exp(B / T + C + D * T + E * T ^ 2)
                Case "17"
                    Return A + Exp(B + C * T + D * T ^ 2 + E * T ^ 3)
                Case "45"
                    Return A * T + B * T ^ 2 / 2 + C * T ^ 3 / 3 + D * T ^ 4 / 4 + E * T ^ 5 / 5
                Case "75"
                    Return B + 2 * C * T + 3 * D * T ^ 2 + 4 * E * T ^ 3
                Case "100"
                    Return A + B * T + C * T ^ 2 + D * T ^ 3 + E * T ^ 4
                Case "101"
                    Return Exp(A + B / T + C * Log(T) + D * T ^ E)
                Case "102"
                    Return A * T ^ B / (1 + C / T + D / T ^ 2)
                Case "103"
                    Return A + B * Exp(-C / (T ^ D))
                Case "104"
                    Return A + B / T + C / T ^ 2 + D / T ^ 8 + E / T ^ 9
                Case "105"
                    Return A / (B ^ (1 + (1 - T / C) ^ D))
                Case "106"
                    Return A * (1 - Tr) ^ (B + C * Tr + D * Tr ^ 2 + E * Tr ^ 3)
                Case "107"
                    Return A + B * (C / T / Sinh(C / T)) ^ 2 + D * (D / T / Cosh(D / T)) ^ 2
                Case "114"
                    Return A * T + B * T ^ 2 / 2 + C * T ^ 3 / 3 + D * T ^ 4 / 4
                Case "115"
                    Return Exp(A + B / T + C * Log(T) + D * T ^ 2 + E / T ^ 2)
                Case "116"
                    Return A + B * (1 - Tr) ^ 0.35 + C * (1 - Tr) ^ (2 / 3) + D * (1 - Tr) + E * (1 - Tr) ^ (4 / 3)
                Case "117"
                    Return A * T + B * (C / T) / Tanh(C / T) - D * (E / T) / Tanh(E / T)
                Case "207"
                    Return Exp(A - B / (T + C))
                Case "208"
                    Return 10 ^ (A - B / (T + C))
                Case "209"
                    Return 10 ^ (A * (1 / T - 1 / B))
                Case "210"
                    Return 10 ^ (A + B / T + C * T + D * T ^ 2)
                Case "211"
                    Return A * ((B - T) / (B - C)) ^ D
                Case Else
                    Return 0
            End Select


        End Function

#End Region

#Region "   CAPE-OPEN 1.0 Methods and Properties"

        Private _compdesc, _compname As String

        ''' <summary>
        ''' Gets the name of the component.
        ''' </summary>
        ''' <value></value>
        ''' <returns>CapeString</returns>
        ''' <remarks>Implements CapeOpen.ICapeIdentification.ComponentDescription</remarks>
        Public Overridable Property ComponentDescription() As String Implements CapeOpen.ICapeIdentification.ComponentDescription
            Get
                Return _compdesc
            End Get
            Set(ByVal value As String)
                _compdesc = value
            End Set
        End Property

        ''' <summary>
        ''' Gets the description of the component.
        ''' </summary>
        ''' <value></value>
        ''' <returns>CapeString</returns>
        ''' <remarks>Implements CapeOpen.ICapeIdentification.ComponentName</remarks>
        Public Overridable Property ComponentName() As String Implements CapeOpen.ICapeIdentification.ComponentName
            Get
                Return _compname
            End Get
            Set(ByVal value As String)
                _compname = value
            End Set
        End Property

        Public Sub PropList1(ByRef props As Object, ByRef phases As Object, ByRef calcType As Object) Implements CAPEOPEN110.ICapeThermoCalculationRoutine.PropList
            Throw New NotImplementedException
        End Sub

        Public Function PropCheck1(ByVal materialObject As Object, ByVal flashType As String, ByVal props As Object) As Object Implements CAPEOPEN110.ICapeThermoEquilibriumServer.PropCheck
            Throw New NotImplementedException
        End Function

        Public Function ValidityCheck1(ByVal materialObject As Object, ByVal props As Object) As Object Implements CAPEOPEN110.ICapeThermoEquilibriumServer.ValidityCheck
            Throw New NotImplementedException
        End Function

        ''' <summary>
        ''' Method responsible for calculating/delegating phase equilibria.
        ''' </summary>
        ''' <param name="materialObject">The Material Object</param>
        ''' <param name="flashType">Flash calculation type.</param>
        ''' <param name="props">Properties to be calculated at equilibrium. UNDEFINED for no
        ''' properties. If a list, then the property values should be set for each
        ''' phase present at equilibrium. (not including the overall phase).</param>
        ''' <remarks><para>On the Material Object the CalcEquilibrium method must set the amounts (phaseFraction) and compositions
        ''' (fraction) for all phases present at equilibrium, as well as the temperature and pressure for the overall
        ''' mixture, if these are not set as part of the calculation specifications. The CalcEquilibrium method must not
        ''' set on the Material Object any other value - in particular it must not set any values for phases that do not
        ''' exist. See 5.2.1 for more information.</para>
        ''' <para>The available list of flashes is given in section 5.6.1.</para>
        ''' <para>When calling this method, it is advised not to combine a flash calculation with a property calculation.
        ''' Through the returned error one cannot see which has failed, plus the additional arguments available in a
        ''' CalcProp call (such as calculation type) cannot be specified in a CalcEquilibrium call. Advice is to perform a
        ''' CalcEquilibrium, get the phaseIDs and perform a CalcProp for the existing phases.</para></remarks>
        Public Overridable Sub CalcEquilibrium(ByVal materialObject As Object, ByVal flashType As String, ByVal props As Object) Implements CapeOpen.ICapeThermoPropertyPackage.CalcEquilibrium

            Dim mymat As ICapeThermoMaterial = materialObject
            Me.CurrentMaterialStream = mymat
            Select Case flashType.ToLower
                Case "tp"
                    Me.DW_CalcEquilibrium(FlashSpec.T, FlashSpec.P)
                Case "ph"
                    Me.DW_CalcEquilibrium(FlashSpec.P, FlashSpec.H)
                Case "ps"
                    Me.DW_CalcEquilibrium(FlashSpec.P, FlashSpec.S)
                Case "tvf"
                    Me.DW_CalcEquilibrium(FlashSpec.T, FlashSpec.VAP)
                Case "pvf"
                    Me.DW_CalcEquilibrium(FlashSpec.P, FlashSpec.VAP)
                Case "pt"
                    Me.DW_CalcEquilibrium(FlashSpec.T, FlashSpec.P)
                Case "hp"
                    Me.DW_CalcEquilibrium(FlashSpec.P, FlashSpec.H)
                Case "sp"
                    Me.DW_CalcEquilibrium(FlashSpec.P, FlashSpec.S)
                Case "vft"
                    Me.DW_CalcEquilibrium(FlashSpec.T, FlashSpec.VAP)
                Case "vfp"
                    Me.DW_CalcEquilibrium(FlashSpec.P, FlashSpec.VAP)
                Case Else
                    Throw New NotImplementedException()
            End Select

        End Sub

        ''' <summary>
        ''' This method is responsible for doing all calculations on behalf of the Calculation Routine component.
        ''' </summary>
        ''' <param name="materialObject">The Material Object of the calculation.</param>
        ''' <param name="props">The list of properties to be calculated.</param>
        ''' <param name="phases">List of phases for which the properties are to be calculated.</param>
        ''' <param name="calcType">Type of calculation: Mixture Property or Pure Compound Property. For
        ''' partial property, such as fugacity coefficients of compounds in a
        ''' mixture, use “Mixture” CalcType. For pure compound fugacity
        ''' coefficients, use “Pure” CalcType.</param>
        ''' <remarks></remarks>
        Public Overridable Sub CalcProp(ByVal materialObject As Object, ByVal props As Object, ByVal phases As Object, ByVal calcType As String) Implements CapeOpen.ICapeThermoPropertyPackage.CalcProp
            Dim mymat As MaterialStream = materialObject
            Me.CurrentMaterialStream = mymat
            Dim ph As String() = phases
            For Each f As String In ph
                For Each pi As PhaseInfo In Me.PhaseMappings.Values
                    If f = pi.PhaseLabel Then
                        If Not pi.DWPhaseID = Fase.Solid Then
                            For Each p As String In props
                                Me.DW_CalcProp(p, pi.DWPhaseID)
                            Next
                            Exit For
                        Else
                            Me.DW_CalcSolidPhaseProps()
                        End If
                    End If
                Next
            Next
        End Sub

        ''' <summary>
        ''' Returns the values of the Constant properties of the compounds contained in the passed Material Object.
        ''' </summary>
        ''' <param name="materialObject">The Material Object.</param>
        ''' <param name="props">The list of properties.</param>
        ''' <returns>Compound Constant values.</returns>
        ''' <remarks></remarks>
        Public Overridable Function GetComponentConstant(ByVal materialObject As Object, ByVal props As Object) As Object Implements CapeOpen.ICapeThermoPropertyPackage.GetComponentConstant
            Dim vals As New ArrayList
            Dim mymat As MaterialStream = materialObject
            For Each c As Substancia In mymat.Fases(0).Componentes.Values
                For Each p As String In props
                    Select Case p.ToLower
                        Case "molecularweight"
                            vals.Add(c.ConstantProperties.Molar_Weight)
                        Case "criticaltemperature"
                            vals.Add(c.ConstantProperties.Critical_Temperature)
                        Case "criticalpressure"
                            vals.Add(c.ConstantProperties.Critical_Pressure)
                        Case "criticalvolume"
                            vals.Add(c.ConstantProperties.Critical_Volume)
                        Case "criticalcompressibilityfactor"
                            vals.Add(c.ConstantProperties.Critical_Compressibility)
                        Case "acentricfactor"
                            vals.Add(c.ConstantProperties.Acentric_Factor)
                        Case "normalboilingpoint"
                            vals.Add(c.ConstantProperties.Normal_Boiling_Point)
                        Case "idealgasgibbsfreeenergyofformationat25c"
                            vals.Add(c.ConstantProperties.IG_Gibbs_Energy_of_Formation_25C * c.ConstantProperties.Molar_Weight)
                        Case "idealgasenthalpyofformationat25c"
                            vals.Add(c.ConstantProperties.IG_Enthalpy_of_Formation_25C * c.ConstantProperties.Molar_Weight)
                        Case "casregistrynumber"
                            vals.Add(c.ConstantProperties.CAS_Number)
                        Case "chemicalformula"
                            vals.Add(c.ConstantProperties.Formula)
                        Case Else
                            Throw New NotImplementedException
                    End Select
                Next
            Next
            Dim arr2(vals.Count) As Double
            Array.Copy(vals.ToArray, arr2, vals.Count)
            Return arr2
        End Function

        ''' <summary>
        ''' Returns the list of compounds of a given Property Package.
        ''' </summary>
        ''' <param name="compIds">List of compound IDs</param>
        ''' <param name="formulae">List of compound formulae</param>
        ''' <param name="names">List of compound names.</param>
        ''' <param name="boilTemps">List of boiling point temperatures.</param>
        ''' <param name="molWt">List of molecular weights.</param>
        ''' <param name="casNo">List of CAS numbers .</param>
        ''' <remarks>Compound identification could be necessary if the PME has internal knowledge of chemical compounds, or in case of
        ''' use of multiple Property Packages. In order to identify the compounds of a Property Package, the PME will use the
        ''' 'casno’ argument instead of the compIds. The reason is that different PMEs may give different names to the same
        ''' chemical compounds, whereas CAS Numbers are universal. Therefore, it is recommended to provide a value for the
        ''' casno argument wherever available.</remarks>
        Public Overridable Sub GetComponentList(ByRef compIds As Object, ByRef formulae As Object, ByRef names As Object, ByRef boilTemps As Object, ByRef molWt As Object, ByRef casNo As Object) Implements CapeOpen.ICapeThermoPropertyPackage.GetComponentList

            Dim ids, formulas, nms, bts, casnos, molws As New ArrayList

            If My.Application.CAPEOPENMode Then
                For Each c As ConstantProperties In _selectedcomps.Values
                    ids.Add(c.Name)
                    formulas.Add(c.Formula)
                    nms.Add(DTL.App.GetComponentName(c.Name))
                    bts.Add(c.Normal_Boiling_Point)
                    casnos.Add(c.CAS_Number)
                    molws.Add(c.Molar_Weight)
                Next
            Else
                For Each c As Substancia In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                    ids.Add(c.ConstantProperties.Name)
                    formulas.Add(c.ConstantProperties.Formula)
                    nms.Add(DTL.App.GetComponentName(c.ConstantProperties.Name))
                    bts.Add(c.ConstantProperties.Normal_Boiling_Point)
                    casnos.Add(c.ConstantProperties.CAS_Number)
                    molws.Add(c.ConstantProperties.Molar_Weight)
                Next
            End If

            Dim _i(ids.Count - 1) As String
            Dim _f(ids.Count - 1) As String
            Dim _n(ids.Count - 1) As String
            Dim _c(ids.Count - 1) As String
            Dim _b(ids.Count - 1) As Double
            Dim _m(ids.Count - 1) As Double

            Array.Copy(ids.ToArray, _i, ids.Count)
            Array.Copy(formulas.ToArray, _f, ids.Count)
            Array.Copy(nms.ToArray, _n, ids.Count)
            Array.Copy(casnos.ToArray, _c, ids.Count)
            Array.Copy(bts.ToArray, _b, ids.Count)
            Array.Copy(molws.ToArray, _m, ids.Count)

            If ids.Count > 0 Then
                compIds = _i
                formulae = _f
                names = _n
                boilTemps = _b
                casNo = _c
                molWt = _m
            Else
                compIds = Nothing
                formulae = Nothing
                names = Nothing
                casNo = Nothing
                boilTemps = Nothing
                molWt = Nothing
            End If

        End Sub

        ''' <summary>
        ''' Provides the list of the supported phases. When supported for one or more property calculations, the Overall
        ''' phase and multiphase identifiers must be returned by this method.
        ''' </summary>
        ''' <returns>The list of phases supported by the Property Package.</returns>
        ''' <remarks></remarks>
        Public Overridable Function GetPhaseList() As Object Implements CapeOpen.ICapeThermoPropertyPackage.GetPhaseList
            Dim pl As New ArrayList
            For Each pi As PhaseInfo In Me.PhaseMappings.Values
                If pi.PhaseLabel <> "Disabled" Then
                    pl.Add(pi.PhaseLabel)
                End If
            Next
            pl.Add("Overall")
            Dim arr(pl.Count - 1) As String
            Array.Copy(pl.ToArray, arr, pl.Count)
            Return arr
        End Function

        ''' <summary>
        ''' Returns list of properties supported by the Property Package.
        ''' </summary>
        ''' <returns>List of all supported Properties.</returns>
        ''' <remarks>GetPropList should return identifiers for the non-constant properties calculated by CalcProp. Standard
        ''' identifiers are listed in 3.10.1. Other non-standard properties that are supported by the Property Package can
        ''' also be returned. GetPropList must not return identifiers for compound constant properties returned by
        ''' GetComponentConstant.
        ''' The properties temperature, pressure, fraction, flow, phaseFraction, totalFlow cannot be returned by
        ''' GetPropList, since all thermodynamic software components must support them. Although the property
        ''' identifier of derivative properties is formed from the identifier of another property, the GetPropList method
        ''' must return the identifiers of all supported derivative and non-derivative properties. For instance, a Property
        ''' Package could return the following list:
        ''' enthalpy, enthalpy.Dtemperature, entropy, entropy.Dpressure.</remarks>
        Public Overridable Function GetPropList() As Object Implements CapeOpen.ICapeThermoPropertyPackage.GetPropList
            Dim arr As New ArrayList
            With arr
                .Add("vaporPressure")
                .Add("surfaceTension")
                .Add("compressibilityFactor")
                .Add("heatOfVaporization")
                .Add("heatCapacity")
                .Add("heatCapacityCv")
                .Add("idealGasHeatCapacity")
                .Add("idealGasEnthalpy")
                .Add("excessEnthalpy")
                .Add("excessGibbsFreeEnergy")
                .Add("excessEntropy")
                .Add("viscosity")
                .Add("thermalConductivity")
                .Add("fugacity")
                .Add("fugacityCoefficient")
                .Add("activity")
                .Add("activityCoefficient")
                .Add("dewPointPressure")
                .Add("dewPointTemperature")
                .Add("kvalue")
                .Add("logFugacityCoefficient")
                .Add("logkvalue")
                .Add("volume")
                .Add("density")
                .Add("enthalpy")
                .Add("entropy")
                .Add("enthalpyF")
                .Add("entropyF")
                .Add("enthalpyNF")
                .Add("entropyNF")
                .Add("gibbsFreeEnergy")
                .Add("moles")
                .Add("mass")
                .Add("molecularWeight")
                .Add("boilingPointTemperature")
            End With

            Dim arr2(arr.Count) As String
            Array.Copy(arr.ToArray, arr2, arr.Count)
            Return arr2

        End Function

        ''' <summary>
        ''' Returns the values of the Universal Constants.
        ''' </summary>
        ''' <param name="materialObject">The Material Object.</param>
        ''' <param name="props">List of requested Universal Constants;</param>
        ''' <returns>Values of Universal Constants.</returns>
        ''' <remarks></remarks>
        Public Overridable Function GetUniversalConstant(ByVal materialObject As Object, ByVal props As Object) As Object Implements CapeOpen.ICapeThermoPropertyPackage.GetUniversalConstant
            Dim res As New ArrayList
            For Each p As String In props
                Select Case p.ToLower
                    Case "standardaccelerationofgravity"
                        res.Add(9.80665)
                    Case "avogadroconstant"
                        res.Add(6.0221419947E+23)
                    Case "boltzmannconstant"
                        res.Add(1.38065324E-23)
                    Case "molargasconstant"
                        res.Add(8.31447215)
                End Select
            Next

            Dim arr2(res.Count) As Double
            Array.Copy(res.ToArray, arr2, res.Count)
            Return arr2

        End Function

        ''' <summary>
        ''' Check to see if properties can be calculated.
        ''' </summary>
        ''' <param name="materialObject">The Material Object for the calculations.</param>
        ''' <param name="props">List of Properties to check.</param>
        ''' <returns>The array of booleans for each property.</returns>
        ''' <remarks>As it was unclear from the original specification what PropCheck should exactly be checking, and as the
        ''' argument list does not include a phase specification, implementations vary. It is generally expected that
        ''' PropCheck at least verifies that the Property is available for calculation in the property Package. However,
        ''' this can also be verified with PropList. It is advised not to use PropCheck.</remarks>
        Public Overridable Function PropCheck(ByVal materialObject As Object, ByVal props As Object) As Object Implements CapeOpen.ICapeThermoPropertyPackage.PropCheck
            Return True
        End Function

        ''' <summary>
        ''' Checks the validity of the calculation. This method is deprecated.
        ''' </summary>
        ''' <param name="materialObject">The Material Object for the calculations.</param>
        ''' <param name="props">The list of properties to check.</param>
        ''' <returns>The properties for which reliability is checked.</returns>
        ''' <remarks>The ValidityCheck method must not be used, since the ICapeThermoReliability interface is not yet defined.</remarks>
        Public Overridable Function ValidityCheck(ByVal materialObject As Object, ByVal props As Object) As Object Implements CapeOpen.ICapeThermoPropertyPackage.ValidityCheck
            Return True
        End Function

        ''' <summary>
        ''' Method responsible for calculating phase equilibria.
        ''' </summary>
        ''' <param name="materialObject">Material Object of the calculation</param>
        ''' <param name="flashType">Flash calculation type.</param>
        ''' <param name="props">Properties to be calculated at equilibrium. UNDEFINED for no
        ''' properties. If a list, then the property values should be set for each
        ''' phase present at equilibrium. (not including the overall phase).</param>
        ''' <remarks>The CalcEquilibrium method must set on the Material Object the amounts (phaseFraction) and compositions
        ''' (fraction) for all phases present at equilibrium, as well as the temperature and pressure for the overall
        ''' mixture, if not set as part of the calculation specifications. The CalcEquilibrium method must not set on the
        ''' Material Object any other value - in particular it must not set any values for phases that do not exist. See
        ''' 5.2.1 for more information.
        ''' The available list of flashes is given in section 5.6.1.
        ''' It is advised not to combine a flash calculation with a property calculation. By the returned error one cannot
        ''' see which has failed, plus the additional arguments to CalcProp (such as calculation type) cannot be
        ''' specified. Advice is to perform a CalcEquilibrium, get the phaseIDs and perform a CalcProp for those
        ''' phases.</remarks>
        Public Overridable Sub CalcEquilibrium2(ByVal materialObject As Object, ByVal flashType As String, ByVal props As Object) Implements CapeOpen.ICapeThermoEquilibriumServer.CalcEquilibrium
            CalcEquilibrium(materialObject, flashType, props)
        End Sub

        ''' <summary>
        ''' Returns the flash types, properties, phases, and calculation types that are supported by a given Equilibrium Server Routine.
        ''' </summary>
        ''' <param name="flashType">Type of flash calculations supported.</param>
        ''' <param name="props">List of supported properties.</param>
        ''' <param name="phases">List of supported phases.</param>
        ''' <param name="calcType">List of supported calculation types.</param>
        ''' <remarks></remarks>
        Public Overridable Sub PropList(ByRef flashType As Object, ByRef props As Object, ByRef phases As Object, ByRef calcType As Object) Implements CapeOpen.ICapeThermoEquilibriumServer.PropList
            props = GetPropList()
            flashType = New String() {"TP", "PH", "PS", "TVF", "PVF", "PT", "HP", "SP", "VFT", "VFP"}
            phases = New String() {"Vapor", "Liquid1", "Liquid2", "Solid", "Overall"}
            calcType = New String() {"Mixture"}
        End Sub

        Public Overridable Sub CalcProp1(ByVal materialObject As Object, ByVal props As Object, ByVal phases As Object, ByVal calcType As String) Implements CapeOpen.ICapeThermoCalculationRoutine.CalcProp
            CalcProp(materialObject, props, phases, calcType)
        End Sub

        Public Overridable Function PropCheck2(ByVal materialObject As Object, ByVal props As Object) As Object Implements CapeOpen.ICapeThermoCalculationRoutine.PropCheck
            Return True
        End Function

        Public Overridable Function ValidityCheck2(ByVal materialObject As Object, ByVal props As Object) As Object Implements CapeOpen.ICapeThermoCalculationRoutine.ValidityCheck
            Return True
        End Function

#End Region

#Region "   CAPE-OPEN 1.1 Thermo & Physical Properties"

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
        Public Overridable Function GetCompoundConstant(ByVal props As Object, ByVal compIds As Object) As Object Implements ICapeThermoCompounds.GetCompoundConstant
            Dim vals As New ArrayList
            For Each s As String In compIds
                Dim c As Substancia = Me.CurrentMaterialStream.Fases(0).Componentes(s)
                For Each p As String In props
                    Select Case p.ToLower
                        Case "molecularweight"
                            vals.Add(c.ConstantProperties.Molar_Weight)
                        Case "criticaltemperature"
                            vals.Add(c.ConstantProperties.Critical_Temperature)
                        Case "criticalpressure"
                            vals.Add(c.ConstantProperties.Critical_Pressure)
                        Case "criticalvolume"
                            vals.Add(c.ConstantProperties.Critical_Volume)
                        Case "criticalcompressibilityfactor"
                            vals.Add(c.ConstantProperties.Critical_Compressibility)
                        Case "acentricfactor"
                            vals.Add(c.ConstantProperties.Acentric_Factor)
                        Case "normalboilingpoint"
                            vals.Add(c.ConstantProperties.Normal_Boiling_Point)
                        Case "idealgasgibbsfreeenergyofformationat25c"
                            vals.Add(c.ConstantProperties.IG_Gibbs_Energy_of_Formation_25C * c.ConstantProperties.Molar_Weight)
                        Case "idealgasenthalpyofformationat25c"
                            vals.Add(c.ConstantProperties.IG_Enthalpy_of_Formation_25C * c.ConstantProperties.Molar_Weight)
                        Case "casregistrynumber"
                            vals.Add(c.ConstantProperties.CAS_Number)
                        Case "chemicalformula", "structureformula"
                            vals.Add(c.ConstantProperties.Formula)
                        Case Else
                            Throw New NotImplementedException
                    End Select
                Next
            Next
            Dim arr2(vals.Count - 1) As Object
            Array.Copy(vals.ToArray, arr2, vals.Count)
            Return arr2
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
        Public Overridable Sub GetCompoundList(ByRef compIds As Object, ByRef formulae As Object, ByRef names As Object, ByRef boilTemps As Object, ByRef molwts As Object, ByRef casnos As Object) Implements ICapeThermoCompounds.GetCompoundList
            GetComponentList(compIds, formulae, names, boilTemps, molwts, casnos)
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
        Public Overridable Function GetConstPropList() As Object Implements ICapeThermoCompounds.GetConstPropList
            Dim vals As New ArrayList
            With vals
                .Add("molecularweight")
                .Add("criticaltemperature")
                .Add("criticalpressure")
                .Add("criticalvolume")
                .Add("criticalcompressibilityfactor")
                .Add("acentricfactor")
                .Add("normalboilingpoint")
                .Add("idealgasgibbsfreeenergyofformationat25c")
                .Add("idealgasenthalpyofformationat25c")
                .Add("casregistrynumber")
                .Add("chemicalformula")
            End With
            Dim arr2(vals.Count - 1) As String
            Array.Copy(vals.ToArray, arr2, vals.Count)
            Return arr2
        End Function

        ''' <summary>
        ''' Returns the number of Compounds supported.
        ''' </summary>
        ''' <returns>Number of Compounds supported.</returns>
        ''' <remarks>The number of Compounds returned by this method must be equal to the number of
        ''' Compound identifiers that are returned by the GetCompoundList method of this interface. It
        ''' must be zero or a positive number.</remarks>
        Public Overridable Function GetNumCompounds() As Integer Implements ICapeThermoCompounds.GetNumCompounds
            If My.Application.CAPEOPENMode Then
                Return Me._selectedcomps.Count
            Else
                Return Me.CurrentMaterialStream.Fases(0).Componentes.Count
            End If
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
        Public Overridable Sub GetPDependentProperty(ByVal props As Object, ByVal pressure As Double, ByVal compIds As Object, ByRef propVals As Object) Implements ICapeThermoCompounds.GetPDependentProperty
            Dim vals As New ArrayList
            For Each c As String In compIds
                For Each p As String In props
                    Select Case p.ToLower
                        Case "boilingpointtemperature"
                            vals.Add(Me.AUX_TSATi(pressure, c))
                    End Select
                Next
            Next
            Dim arr2(vals.Count - 1) As Double
            Array.Copy(vals.ToArray, arr2, vals.Count)
            propVals = arr2
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
        Public Overridable Function GetPDependentPropList() As Object Implements ICapeThermoCompounds.GetPDependentPropList
            Return New String() {"boilingPointTemperature"}
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
        Public Overridable Sub GetTDependentProperty(ByVal props As Object, ByVal temperature As Double, ByVal compIds As Object, ByRef propVals As Object) Implements ICapeThermoCompounds.GetTDependentProperty
            Dim vals As New ArrayList
            For Each c As String In compIds
                For Each p As String In props
                    Select Case p.ToLower
                        Case "heatofvaporization"
                            vals.Add(Me.AUX_HVAPi(c, temperature) * Me.CurrentMaterialStream.Fases(0).Componentes(c).ConstantProperties.Molar_Weight)
                        Case "idealgasenthalpy"
                            vals.Add(Me.RET_Hid_i(298.15, temperature, c) * Me.CurrentMaterialStream.Fases(0).Componentes(c).ConstantProperties.Molar_Weight)
                        Case "idealgasentropy"
                            vals.Add(Me.RET_Sid_i(298.15, temperature, 101325, c) * Me.CurrentMaterialStream.Fases(0).Componentes(c).ConstantProperties.Molar_Weight)
                        Case "idealgasheatcapacity"
                            vals.Add(Me.AUX_CPi(c, temperature) * Me.CurrentMaterialStream.Fases(0).Componentes(c).ConstantProperties.Molar_Weight)
                        Case "vaporpressure"
                            vals.Add(Me.AUX_PVAPi(c, temperature))
                        Case "viscosityofliquid"
                            vals.Add(Me.AUX_LIQVISCi(c, temperature))
                        Case "heatcapacityofliquid"
                            vals.Add(Me.AUX_LIQ_Cpi(Me.CurrentMaterialStream.Fases(0).Componentes(c).ConstantProperties, temperature) * Me.CurrentMaterialStream.Fases(0).Componentes(c).ConstantProperties.Molar_Weight)
                        Case "heatcapacityofsolid"
                            vals.Add(Me.AUX_SolidHeatCapacity(Me.CurrentMaterialStream.Fases(0).Componentes(c).ConstantProperties, temperature) * Me.CurrentMaterialStream.Fases(0).Componentes(c).ConstantProperties.Molar_Weight)
                        Case "thermalconductivityofliquid"
                            vals.Add(Me.AUX_LIQTHERMCONDi(Me.CurrentMaterialStream.Fases(0).Componentes(c).ConstantProperties, temperature))
                        Case "thermalconductivityofvapor"
                            vals.Add(Me.AUX_VAPTHERMCONDi(Me.CurrentMaterialStream.Fases(0).Componentes(c).ConstantProperties, temperature, Me.AUX_PVAPi(Me.CurrentMaterialStream.Fases(0).Componentes(c).ConstantProperties.Name, temperature)))
                        Case "viscosityofvapor"
                            vals.Add(Me.AUX_VAPVISCi(Me.CurrentMaterialStream.Fases(0).Componentes(c).ConstantProperties, temperature))
                        Case "densityofliquid"
                            vals.Add(Me.AUX_LIQDENSi(Me.CurrentMaterialStream.Fases(0).Componentes(c).ConstantProperties, temperature))
                        Case "densityofsolid"
                            vals.Add(Me.AUX_SOLIDDENSi(Me.CurrentMaterialStream.Fases(0).Componentes(c).ConstantProperties, temperature))
                    End Select
                Next
            Next
            Dim arr2(vals.Count - 1) As Double
            Array.Copy(vals.ToArray, arr2, vals.Count)
            propVals = arr2
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
        Public Overridable Function GetTDependentPropList() As Object Implements ICapeThermoCompounds.GetTDependentPropList
            Dim vals As New ArrayList
            With vals
                .Add("heatOfVaporization")
                .Add("idealGasEnthalpy")
                .Add("idealGasEntropy")
                .Add("idealGasHeatCapacity")
                .Add("vaporPressure")
                .Add("viscosityOfLiquid")
                .Add("heatCapacityOfLiquid")
                .Add("heatCapacityOfSolid")
                .Add("thermalConductivityOfLiquid")
                .Add("thermalConductivityOfVapor")
                .Add("viscosityOfVapor")
                .Add("densityOfLiquid")
                .Add("densityOfSolid")
            End With
            Dim arr2(vals.Count - 1) As String
            Array.Copy(vals.ToArray, arr2, vals.Count)
            Return arr2
        End Function

        Public Overridable Function GetNumPhases() As Integer Implements ICapeThermoPhases.GetNumPhases
            Dim i As Integer = 0
            For Each pi As PhaseInfo In Me.PhaseMappings.Values
                If pi.PhaseLabel <> "Disabled" Then
                    i += 1
                End If
            Next
            Return i
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
        Public Overridable Function GetPhaseInfo(ByVal phaseLabel As String, ByVal phaseAttribute As String) As Object Implements ICapeThermoPhases.GetPhaseInfo
            Dim retval As Object = Nothing
            Select Case phaseLabel
                Case "Vapor"
                    Select Case phaseAttribute
                        Case "StateOfAggregation"
                            retval = "Vapor"
                        Case "keyCompoundId"
                            retval = Nothing
                        Case "ExcludedCompoundId"
                            retval = Nothing
                        Case "DensityDescription"
                            retval = Nothing
                        Case "UserDescription"
                            retval = "Vapor Phase"
                        Case Else
                            retval = Nothing
                    End Select
                Case "Liquid1", "Liquid2", "Liquid3", "Liquid", "Aqueous"
                    Select Case phaseAttribute
                        Case "StateOfAggregation"
                            retval = "Liquid"
                        Case "keyCompoundId"
                            retval = Nothing
                        Case "ExcludedCompoundId"
                            retval = Nothing
                        Case "DensityDescription"
                            retval = Nothing
                        Case "UserDescription"
                            retval = "Liquid Phase"
                        Case Else
                            retval = Nothing
                    End Select
                Case "Solid"
                    Select Case phaseAttribute
                        Case "StateOfAggregation"
                            retval = "Solid"
                        Case "TypeOfSolid"
                            retval = "SolidSolution"
                        Case "keyCompoundId"
                            retval = Nothing
                        Case "ExcludedCompoundId"
                            retval = Nothing
                        Case "DensityDescription"
                            retval = Nothing
                        Case "UserDescription"
                            retval = "Solid Phase"
                        Case Else
                            retval = Nothing
                    End Select
            End Select
            Return retval
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
        Public Overridable Sub GetPhaseList1(ByRef phaseLabels As Object, ByRef stateOfAggregation As Object, ByRef keyCompoundId As Object) Implements ICapeThermoPhases.GetPhaseList
            Dim pl, sa, kci As New ArrayList, tmpstr As Object
            For Each pin As PhaseInfo In Me.PhaseMappings.Values
                If pin.PhaseLabel <> "Disabled" Then
                    pl.Add(pin.PhaseLabel)
                    tmpstr = Me.GetPhaseInfo(pin.PhaseLabel, "StateOfAggregation")
                    sa.Add(tmpstr)
                    tmpstr = Me.GetPhaseInfo(pin.PhaseLabel, "keyCompoundId")
                    kci.Add(tmpstr)
                End If
            Next
            Dim myarr1(pl.Count - 1), myarr2(pl.Count - 1), myarr3(pl.Count - 1) As String
            Array.Copy(pl.ToArray, myarr1, pl.Count)
            Array.Copy(sa.ToArray, myarr2, pl.Count)
            Array.Copy(kci.ToArray, myarr3, pl.Count)
            phaseLabels = myarr1
            stateOfAggregation = myarr2
            keyCompoundId = myarr3
        End Sub

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
        Public Overridable Sub CalcAndGetLnPhi(ByVal phaseLabel As String, ByVal temperature As Double, ByVal pressure As Double, ByVal moleNumbers As Object, ByVal fFlags As Integer, ByRef lnPhi As Object, ByRef lnPhiDT As Object, ByRef lnPhiDP As Object, ByRef lnPhiDn As Object) Implements ICapeThermoPropertyRoutine.CalcAndGetLnPhi
            Select Case fFlags
                Case eCapeCalculationCode.CAPE_LOG_FUGACITY_COEFFICIENTS
                    'normalize mole fractions
                    Dim tmols As Double, Vx(moleNumbers.length) As Double
                    For i As Integer = 0 To moleNumbers.length - 1
                        tmols += moleNumbers(i)
                    Next
                    For i As Integer = 0 To moleNumbers.length - 1
                        moleNumbers(i) /= tmols
                    Next
                    Select Case phaseLabel
                        Case "Vapor"
                            lnPhi = Me.DW_CalcFugCoeff(Vx, temperature, pressure, State.Vapor)
                        Case "Liquid"
                            lnPhi = Me.DW_CalcFugCoeff(Vx, temperature, pressure, State.Liquid)
                        Case "Solid"
                            lnPhi = Me.DW_CalcFugCoeff(Vx, temperature, pressure, State.Solid)
                    End Select
                    For i As Integer = 0 To moleNumbers.length - 1
                        lnPhi(i) = Log(lnPhi(i))
                    Next
                Case Else
                    Throw New Exception
            End Select
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
        Public Overridable Sub CalcSinglePhaseProp(ByVal props As Object, ByVal phaseLabel As String) Implements ICapeThermoPropertyRoutine.CalcSinglePhaseProp

            If Not My.Application.CAPEOPENMode Then

                For Each pi As PhaseInfo In Me.PhaseMappings.Values
                    If phaseLabel = pi.PhaseLabel Then
                        For Each p As String In props
                            Me.DW_CalcProp(p, pi.DWPhaseID)
                        Next
                        'Me.DW_CalcPhaseProps(pi.DWPhaseID)
                        Exit For
                    End If
                Next

            Else

                Dim res As New ArrayList
                Dim comps As New ArrayList

                For Each c As Substancia In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                    comps.Add(c.Nome)
                Next

                Dim f As Integer = -1
                Dim phs As DTL.SimulationObjects.PropertyPackages.Fase
                Select Case phaseLabel.ToLower
                    Case "overall"
                        f = 0
                        phs = PropertyPackages.Fase.Mixture
                    Case Else
                        For Each pi As PhaseInfo In Me.PhaseMappings.Values
                            If phaseLabel = pi.PhaseLabel Then
                                f = pi.DWPhaseIndex
                                phs = pi.DWPhaseID
                                Exit For
                            End If
                        Next
                End Select

                If f = -1 Then
                    Dim ex As New Exception("Invalid Phase ID", New ArgumentException)
                    Dim hcode As Integer = 0
                    ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoMaterial", ex.Source, ex.StackTrace, "CalcSinglePhaseProp", hcode)
                End If

                Dim basis As String = "Mole"

                For Each [property] As String In props

                    Dim mymo As ICapeThermoMaterial = _como
                    Dim T As Double
                    Dim P As Double
                    Dim Vx As Object = Nothing

                    mymo.GetTPFraction(phaseLabel, T, P, Vx)

                    Me.CurrentMaterialStream.Fases(0).SPMProperties.temperature = T
                    Me.CurrentMaterialStream.Fases(0).SPMProperties.pressure = P
                    Me.CurrentMaterialStream.SetPhaseComposition(Vx, phs)

                    If phs = Fase.Solid Then
                        Me.DW_CalcSolidPhaseProps()
                    Else
                        Me.DW_CalcProp([property], phs)
                    End If

                    basis = "Mole"
                    Select Case [property].ToLower
                        Case "compressibilityfactor"
                            res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.compressibilityFactor.GetValueOrDefault)
                            basis = ""
                        Case "heatofvaporization"
                        Case "heatcapacity", "heatcapacitycp"
                            Select Case basis
                                Case "Molar", "molar", "mole", "Mole"
                                    res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.heatCapacityCp.GetValueOrDefault * Me.AUX_MMM(phs))
                                Case "Mass", "mass"
                                    res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.heatCapacityCp.GetValueOrDefault * 1000)
                            End Select
                        Case "heatcapacitycv"
                            Select Case basis
                                Case "Molar", "molar", "mole", "Mole"
                                    res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.heatCapacityCv.GetValueOrDefault * Me.AUX_MMM(phs))
                                Case "Mass", "mass"
                                    res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.heatCapacityCv.GetValueOrDefault * 1000)
                            End Select
                        Case "idealgasheatcapacity"
                            If f = 1 Then
                                res.Add(Me.CurrentMaterialStream.PropertyPackage.AUX_CPm(PropertyPackages.Fase.Liquid, Me.CurrentMaterialStream.Fases(0).SPMProperties.temperature * 1000))
                            ElseIf f = 2 Then
                                res.Add(Me.CurrentMaterialStream.PropertyPackage.AUX_CPm(PropertyPackages.Fase.Vapor, Me.CurrentMaterialStream.Fases(0).SPMProperties.temperature * 1000))
                            Else
                                res.Add(Me.CurrentMaterialStream.PropertyPackage.AUX_CPm(PropertyPackages.Fase.Solid, Me.CurrentMaterialStream.Fases(0).SPMProperties.temperature * 1000))
                            End If
                        Case "idealgasenthalpy"
                            If f = 1 Then
                                res.Add(Me.CurrentMaterialStream.PropertyPackage.RET_Hid(298.15, Me.CurrentMaterialStream.Fases(0).SPMProperties.temperature.GetValueOrDefault * 1000, PropertyPackages.Fase.Liquid))
                            ElseIf f = 2 Then
                                res.Add(Me.CurrentMaterialStream.PropertyPackage.RET_Hid(298.15, Me.CurrentMaterialStream.Fases(0).SPMProperties.temperature.GetValueOrDefault * 1000, PropertyPackages.Fase.Vapor))
                            Else
                                res.Add(Me.CurrentMaterialStream.PropertyPackage.RET_Hid(298.15, Me.CurrentMaterialStream.Fases(0).SPMProperties.temperature.GetValueOrDefault * 1000, PropertyPackages.Fase.Solid))
                            End If
                        Case "excessenthalpy"
                            Select Case basis
                                Case "Molar", "molar", "mole", "Mole"
                                    res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.excessEnthalpy.GetValueOrDefault * Me.AUX_MMM(phs))
                                Case "Mass", "mass"
                                    res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.excessEnthalpy.GetValueOrDefault * 1000)
                            End Select
                        Case "excessentropy"
                            Select Case basis
                                Case "Molar", "molar", "mole", "Mole"
                                    res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.excessEntropy.GetValueOrDefault * Me.AUX_MMM(phs))
                                Case "Mass", "mass"
                                    res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.excessEntropy.GetValueOrDefault * 1000)
                            End Select
                        Case "viscosity"
                            res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.viscosity.GetValueOrDefault)
                            basis = ""
                        Case "thermalconductivity"
                            res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.thermalConductivity.GetValueOrDefault)
                            basis = ""
                        Case "fugacity"
                            For Each c As String In comps
                                res.Add(Me.CurrentMaterialStream.Fases(f).Componentes(c).FracaoMolar.GetValueOrDefault * Me.CurrentMaterialStream.Fases(f).Componentes(c).FugacityCoeff.GetValueOrDefault * Me.CurrentMaterialStream.Fases(0).SPMProperties.pressure.GetValueOrDefault)
                            Next
                            basis = ""
                        Case "activity"
                            For Each c As String In comps
                                res.Add(Me.CurrentMaterialStream.Fases(f).Componentes(c).ActivityCoeff.GetValueOrDefault * Me.CurrentMaterialStream.Fases(f).Componentes(c).FracaoMolar.GetValueOrDefault)
                            Next
                            basis = ""
                        Case "fugacitycoefficient"
                            For Each c As String In comps
                                res.Add(Me.CurrentMaterialStream.Fases(f).Componentes(c).FugacityCoeff.GetValueOrDefault)
                            Next
                            basis = ""
                        Case "activitycoefficient"
                            For Each c As String In comps
                                res.Add(Me.CurrentMaterialStream.Fases(f).Componentes(c).ActivityCoeff.GetValueOrDefault)
                            Next
                            basis = ""
                        Case "logfugacitycoefficient"
                            For Each c As String In comps
                                res.Add(Math.Log(Me.CurrentMaterialStream.Fases(f).Componentes(c).FugacityCoeff.GetValueOrDefault))
                            Next
                            basis = ""
                        Case "volume"
                            res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.molecularWeight.GetValueOrDefault / Me.CurrentMaterialStream.Fases(f).SPMProperties.density.GetValueOrDefault / 1000)
                        Case "density"
                            res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.density.GetValueOrDefault / Me.AUX_MMM(phs) * 1000)
                        Case "enthalpy", "enthalpynf"
                            Select Case basis
                                Case "Molar", "molar", "mole", "Mole"
                                    Dim val = Me.CurrentMaterialStream.Fases(f).SPMProperties.molecularWeight.GetValueOrDefault
                                    If val = 0.0# Then
                                        res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.molar_enthalpy.GetValueOrDefault)
                                    Else
                                        res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.enthalpy.GetValueOrDefault * val)
                                    End If
                                Case "Mass", "mass"
                                    res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.enthalpy.GetValueOrDefault * 1000)
                            End Select
                        Case "entropy", "entropynf"
                            Select Case basis
                                Case "Molar", "molar", "mole", "Mole"
                                    Dim val = Me.CurrentMaterialStream.Fases(f).SPMProperties.molecularWeight.GetValueOrDefault
                                    If val = 0.0# Then
                                        res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.molar_entropy.GetValueOrDefault)
                                    Else
                                        res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.entropy.GetValueOrDefault * val)
                                    End If
                                Case "Mass", "mass"
                                    res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.entropy.GetValueOrDefault * 1000)
                            End Select
                        Case "enthalpyf"
                            Select Case basis
                                Case "Molar", "molar", "mole", "Mole"
                                    Dim val = Me.CurrentMaterialStream.Fases(f).SPMProperties.molecularWeight.GetValueOrDefault
                                    If val = 0.0# Then
                                        res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.molar_enthalpyF.GetValueOrDefault)
                                    Else
                                        res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.enthalpyF.GetValueOrDefault * val)
                                    End If
                                Case "Mass", "mass"
                                    res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.enthalpyF.GetValueOrDefault * 1000)
                            End Select
                        Case "entropyf"
                            Select Case basis
                                Case "Molar", "molar", "mole", "Mole"
                                    Dim val = Me.CurrentMaterialStream.Fases(f).SPMProperties.molecularWeight.GetValueOrDefault
                                    If val = 0.0# Then
                                        res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.molar_entropyF.GetValueOrDefault)
                                    Else
                                        res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.entropyF.GetValueOrDefault * val)
                                    End If
                                Case "Mass", "mass"
                                    res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.entropy.GetValueOrDefault * 1000)
                            End Select
                        Case "moles"
                            res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.molarflow.GetValueOrDefault)
                            basis = ""
                        Case "mass"
                            res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.massflow.GetValueOrDefault)
                            basis = ""
                        Case "molecularweight"
                            res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.molecularWeight.GetValueOrDefault)
                            basis = ""
                        Case "temperature"
                            res.Add(Me.CurrentMaterialStream.Fases(0).SPMProperties.temperature.GetValueOrDefault)
                            basis = ""
                        Case "pressure"
                            res.Add(Me.CurrentMaterialStream.Fases(0).SPMProperties.pressure.GetValueOrDefault)
                            basis = ""
                        Case "flow"
                            Select Case basis
                                Case "Molar", "molar", "mole", "Mole"
                                    For Each c As String In comps
                                        res.Add(Me.CurrentMaterialStream.Fases(f).Componentes(c).MolarFlow.GetValueOrDefault)
                                    Next
                                Case "Mass", "mass"
                                    For Each c As String In comps
                                        res.Add(Me.CurrentMaterialStream.Fases(f).Componentes(c).MassFlow.GetValueOrDefault)
                                    Next
                            End Select
                        Case "fraction", "massfraction", "molarfraction"
                            Select Case basis
                                Case "Molar", "molar", "mole", "Mole"
                                    For Each c As String In comps
                                        res.Add(Me.CurrentMaterialStream.Fases(f).Componentes(c).FracaoMolar.GetValueOrDefault)
                                    Next
                                Case "Mass", "mass"
                                    For Each c As String In comps
                                        res.Add(Me.CurrentMaterialStream.Fases(f).Componentes(c).FracaoMassica.GetValueOrDefault)
                                    Next
                                Case ""
                                    If [property].ToLower.Contains("mole") Then
                                        For Each c As String In comps
                                            res.Add(Me.CurrentMaterialStream.Fases(f).Componentes(c).FracaoMolar.GetValueOrDefault)
                                        Next
                                    ElseIf [property].ToLower.Contains("mass") Then
                                        For Each c As String In comps
                                            res.Add(Me.CurrentMaterialStream.Fases(f).Componentes(c).FracaoMassica.GetValueOrDefault)
                                        Next
                                    End If
                            End Select
                        Case "concentration"
                            For Each c As String In comps
                                res.Add(Me.CurrentMaterialStream.Fases(f).Componentes(c).MassFlow.GetValueOrDefault / Me.CurrentMaterialStream.Fases(f).SPMProperties.volumetric_flow.GetValueOrDefault)
                            Next
                            basis = ""
                        Case "molarity"
                            For Each c As String In comps
                                res.Add(Me.CurrentMaterialStream.Fases(f).Componentes(c).MolarFlow.GetValueOrDefault / Me.CurrentMaterialStream.Fases(f).SPMProperties.volumetric_flow.GetValueOrDefault)
                            Next
                            basis = ""
                        Case "phasefraction"
                            Select Case basis
                                Case "Molar", "molar", "mole", "Mole"
                                    res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.molarfraction.GetValueOrDefault)
                                Case "Mass", "mass"
                                    res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.massfraction.GetValueOrDefault)
                            End Select
                        Case "totalflow"
                            Select Case basis
                                Case "Molar", "molar", "mole", "Mole"
                                    res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.molarflow.GetValueOrDefault)
                                Case "Mass", "mass"
                                    res.Add(Me.CurrentMaterialStream.Fases(f).SPMProperties.massflow.GetValueOrDefault)
                            End Select
                        Case Else
                            Dim ex = New NotImplementedException()
                            Dim hcode As Integer = 0
                            ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoMaterial", ex.Source, ex.StackTrace, "CalcSinglePhaseProp", hcode)
                    End Select

                    Dim i As Integer
                    For i = 0 To res.Count - 1
                        If Double.IsNaN(res(i)) Then res(i) = 0.0#
                    Next

                    Dim arr(res.Count - 1) As Double
                    Array.Copy(res.ToArray, arr, res.Count)

                    mymo.SetSinglePhaseProp([property], phaseLabel, basis, arr)

                Next

            End If

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
        Public Overridable Sub CalcTwoPhaseProp(ByVal props As Object, ByVal phaseLabels As Object) Implements ICapeThermoPropertyRoutine.CalcTwoPhaseProp

            Me.DW_CalcTwoPhaseProps(Fase.Liquid, Fase.Vapor)

            If My.Application.CAPEOPENMode Then

                Dim res As New ArrayList
                Dim comps As New ArrayList
                For Each c As Substancia In Me.CurrentMaterialStream.Fases(0).Componentes.Values
                    comps.Add(c.Nome)
                Next

                Dim basis As String = Nothing

                For Each [property] As String In props
                    Select Case [property].ToLower
                        Case "kvalue"
                            For Each c As String In comps
                                res.Add(Me.CurrentMaterialStream.Fases(0).Componentes(c).Kvalue)
                            Next
                        Case "logkvalue"
                            For Each c As String In comps
                                res.Add(Me.CurrentMaterialStream.Fases(0).Componentes(c).lnKvalue)
                            Next
                        Case "surfacetension"
                            res.Add(Me.CurrentMaterialStream.Fases(0).TPMProperties.surfaceTension.GetValueOrDefault)
                        Case Else
                            Dim ex = New Exception
                            Dim hcode As Integer = 0
                            ThrowCAPEException(ex, "Error", ex.Message, "ICapeThermoMaterial", ex.Source, ex.StackTrace, "CalcTwoPhaseProp", hcode)
                    End Select

                    Dim arr(res.Count - 1) As Double
                    Array.Copy(res.ToArray, arr, res.Count)

                    Dim mymo As ICapeThermoMaterial = _como
                    mymo.SetTwoPhaseProp([property], phaseLabels, basis, arr)

                Next

            End If

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
        Public Overridable Function CheckSinglePhasePropSpec(ByVal [property] As String, ByVal phaseLabel As String) As Boolean Implements ICapeThermoPropertyRoutine.CheckSinglePhasePropSpec
            Select Case [property].ToLower
                Case "compressibilityfactor", "heatofvaporization", "heatcapacity", "heatcapacitycv", _
                    "idealgasheatcapacity", "idealgasenthalpy", "excessenthalpy", "excessentropy", _
                    "viscosity", "thermalconductivity", "fugacity", "fugacitycoefficient", "activity", "activitycoefficient", _
                    "dewpointpressure", "dewpointtemperature", "logfugacitycoefficient", "volume", "density", _
                    "enthalpy", "entropy", "gibbsfreeenergy", "moles", "mass", "molecularweight", "totalflow"
                    Return True
                Case Else
                    Return False
            End Select
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
        Public Overridable Function CheckTwoPhasePropSpec(ByVal [property] As String, ByVal phaseLabels As Object) As Boolean Implements ICapeThermoPropertyRoutine.CheckTwoPhasePropSpec
            Return True
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
        Public Overridable Function GetSinglePhasePropList() As Object Implements ICapeThermoPropertyRoutine.GetSinglePhasePropList
            Dim arr As New ArrayList
            With arr
                .Add("compressibilityFactor")
                .Add("heatCapacityCp")
                .Add("heatCapacityCv")
                .Add("excessEnthalpy")
                .Add("excessEntropy")
                .Add("viscosity")
                .Add("thermalConductivity")
                .Add("fugacity")
                .Add("fugacityCoefficient")
                .Add("activityCoefficient")
                .Add("logFugacityCoefficient")
                .Add("volume")
                .Add("density")
                .Add("enthalpy")
                .Add("entropy")
                .Add("enthalpyF")
                .Add("entropyF")
                .Add("enthalpyNF")
                .Add("entropyNF")
                .Add("molecularWeight")
            End With
            Dim arr2(arr.Count - 1) As String
            Array.Copy(arr.ToArray, arr2, arr.Count)
            Return arr2
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
        Public Overridable Function GetTwoPhasePropList() As Object Implements ICapeThermoPropertyRoutine.GetTwoPhasePropList
            Return New String() {"kvalue", "logKvalue", "surfaceTension"}
        End Function

        ''' <summary>
        ''' Retrieves the value of a Universal Constant.
        ''' </summary>
        ''' <param name="constantId">Identifier of Universal Constant. The list of constants supported should be 
        ''' obtained by using the GetUniversalConstantList method.</param>
        ''' <returns>Value of Universal Constant. This could be a numeric or a string value. For numeric values 
        ''' the units of measurement are specified in section 7.5.1.</returns>
        ''' <remarks>Universal Constants (often called fundamental constants) are quantities like the gas constant,
        ''' or the Avogadro constant.</remarks>
        Public Overridable Function GetUniversalConstant1(ByVal constantId As String) As Object Implements ICapeThermoUniversalConstant.GetUniversalConstant
            Dim res As New ArrayList
            Select Case constantId.ToLower
                Case "standardaccelerationofgravity"
                    res.Add(9.80665)
                Case "avogadroconstant"
                    res.Add(6.0221419947E+23)
                Case "boltzmannconstant"
                    res.Add(1.38065324E-23)
                Case "molargasconstant"
                    res.Add(8.31447215)
            End Select
            Dim arr2(res.Count) As Object
            Array.Copy(res.ToArray, arr2, res.Count)
            Return arr2
        End Function

        ''' <summary>
        ''' Returns the identifiers of the supported Universal Constants.
        ''' </summary>
        ''' <returns>List of identifiers of Universal Constants. The list of standard identifiers is given in section 7.5.1.</returns>
        ''' <remarks>A component may return Universal Constant identifiers that do not belong to the list defined
        ''' in section 7.5.1. However, these proprietary identifiers may not be understood by most of the
        ''' clients of this component.</remarks>
        Public Overridable Function GetUniversalConstantList() As Object Implements ICapeThermoUniversalConstant.GetUniversalConstantList
            Return New String() {"standardAccelerationOfGravity", "avogadroConstant", "boltzmannConstant", "molarGasConstant"}
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
        Public Overridable Sub CalcEquilibrium1(ByVal specification1 As Object, ByVal specification2 As Object, ByVal solutionType As String) Implements ICapeThermoEquilibriumRoutine.CalcEquilibrium

            Dim spec1, spec2 As FlashSpec
            If specification1(0).ToString.ToLower = "temperature" And specification2(0).ToString.ToLower = "pressure" Then
                spec1 = FlashSpec.T
                spec2 = FlashSpec.P
            ElseIf specification1(0).ToString.ToLower = "pressure" And specification2(0).ToString.ToLower = "enthalpy" Then
                spec1 = FlashSpec.P
                spec2 = FlashSpec.H
            ElseIf specification1(0).ToString.ToLower = "pressure" And specification2(0).ToString.ToLower = "entropy" Then
                spec1 = FlashSpec.P
                spec2 = FlashSpec.S
            ElseIf specification1(0).ToString.ToLower = "pressure" And specification2(0).ToString.ToLower = "phasefraction" Then
                spec1 = FlashSpec.P
                spec2 = FlashSpec.VAP
            ElseIf specification1(0).ToString.ToLower = "temperature" And specification2(0).ToString.ToLower = "phasefraction" Then
                spec1 = FlashSpec.T
                spec2 = FlashSpec.VAP
            ElseIf specification2(0).ToString.ToLower = "temperature" And specification1(0).ToString.ToLower = "pressure" Then
                spec1 = FlashSpec.T
                spec2 = FlashSpec.P
            ElseIf specification2(0).ToString.ToLower = "pressure" And specification1(0).ToString.ToLower = "enthalpy" Then
                spec1 = FlashSpec.P
                spec2 = FlashSpec.H
            ElseIf specification2(0).ToString.ToLower = "pressure" And specification1(0).ToString.ToLower = "entropy" Then
                spec1 = FlashSpec.P
                spec2 = FlashSpec.S
            ElseIf specification2(0).ToString.ToLower = "pressure" And specification1(0).ToString.ToLower = "phasefraction" Then
                spec1 = FlashSpec.P
                spec2 = FlashSpec.VAP
            ElseIf specification2(0).ToString.ToLower = "temperature" And specification1(0).ToString.ToLower = "phasefraction" Then
                spec1 = FlashSpec.T
                spec2 = FlashSpec.VAP
            Else
                Throw New NotImplementedException("Flash spec not supported.")
            End If

            Dim T, P, Hm, Sm As Double

            If My.Application.CAPEOPENMode Then

                Dim res As Object = Nothing

                Dim mys As ICapeThermoMaterial = _como

                If spec1 = FlashSpec.T And spec2 = P Then
                    mys.GetOverallProp("temperature", Nothing, res)
                    T = res(0)
                    mys.GetOverallProp("pressure", Nothing, res)
                    P = res(0)
                    Me.CurrentMaterialStream.Fases(0).SPMProperties.temperature = T
                    Me.CurrentMaterialStream.Fases(0).SPMProperties.pressure = P
                ElseIf spec1 = FlashSpec.T And spec2 = FlashSpec.VAP Then
                    mys.GetOverallProp("temperature", Nothing, res)
                    T = res(0)
                    Me.CurrentMaterialStream.Fases(0).SPMProperties.temperature = T
                    mys.GetSinglePhaseProp("phaseFraction", "Vapor", "Mole", res)
                    Me.CurrentMaterialStream.Fases(2).SPMProperties.molarfraction = res(0)
                ElseIf spec1 = FlashSpec.P And spec2 = FlashSpec.H Then
                    mys.GetOverallProp("pressure", Nothing, res)
                    P = res(0)
                    mys.GetOverallProp("enthalpy", "Mole", res)
                    Hm = res(0) / AUX_MMM(Fase.Mixture)
                    Me.CurrentMaterialStream.Fases(0).SPMProperties.enthalpy = Hm
                    Me.CurrentMaterialStream.Fases(0).SPMProperties.pressure = P
                ElseIf spec1 = FlashSpec.P And spec2 = FlashSpec.S Then
                    mys.GetOverallProp("pressure", Nothing, res)
                    P = res(0)
                    mys.GetOverallProp("entropy", "Mole", res)
                    Sm = res(0) / AUX_MMM(Fase.Mixture)
                    Me.CurrentMaterialStream.Fases(0).SPMProperties.entropy = Sm
                    Me.CurrentMaterialStream.Fases(0).SPMProperties.pressure = P
                ElseIf spec1 = FlashSpec.P And spec2 = FlashSpec.VAP Then
                    mys.GetOverallProp("pressure", Nothing, res)
                    P = res(0)
                    Me.CurrentMaterialStream.Fases(0).SPMProperties.pressure = P
                    mys.GetSinglePhaseProp("phaseFraction", "Vapor", "Mole", res)
                    Me.CurrentMaterialStream.Fases(2).SPMProperties.molarfraction = res(0)
                End If

            End If

            Try
                Me.DW_CalcEquilibrium(spec1, spec2)
            Catch ex As Exception
                ThrowCAPEException(ex, ex.GetType.ToString, ex.ToString, "ICapeThermoEquilibriumRoutine", ex.ToString, "CalcEquilibrium", "", 0)
            End Try

            If My.Application.CAPEOPENMode Then

                Dim ms As MaterialStream = Me.CurrentMaterialStream
                Dim mo As ICapeThermoMaterial = _como

                Dim vok As Boolean = False
                Dim l1ok As Boolean = False
                Dim l2ok As Boolean = False
                Dim sok As Boolean = False

                If ms.Fases(2).SPMProperties.molarfraction.HasValue Then vok = True
                If ms.Fases(3).SPMProperties.molarfraction.HasValue Then l1ok = True
                If ms.Fases(4).SPMProperties.molarfraction.HasValue Then l2ok = True
                If ms.Fases(7).SPMProperties.molarfraction.HasValue Then sok = True

                Dim phases As String() = Nothing
                Dim statuses As eCapePhaseStatus() = Nothing

                If vok And l1ok And l2ok Then
                    phases = New String() {"Vapor", "Liquid", "Liquid2"}
                    statuses = New eCapePhaseStatus() {eCapePhaseStatus.CAPE_ATEQUILIBRIUM, eCapePhaseStatus.CAPE_ATEQUILIBRIUM, eCapePhaseStatus.CAPE_ATEQUILIBRIUM}
                ElseIf vok And l1ok And Not l2ok Then
                    phases = New String() {"Vapor", "Liquid"}
                    statuses = New eCapePhaseStatus() {eCapePhaseStatus.CAPE_ATEQUILIBRIUM, eCapePhaseStatus.CAPE_ATEQUILIBRIUM}
                ElseIf vok And l1ok And Not l2ok And sok Then
                    phases = New String() {"Vapor", "Liquid", "Solid"}
                    statuses = New eCapePhaseStatus() {eCapePhaseStatus.CAPE_ATEQUILIBRIUM, eCapePhaseStatus.CAPE_ATEQUILIBRIUM, eCapePhaseStatus.CAPE_ATEQUILIBRIUM}
                ElseIf vok And Not l1ok And l2ok Then
                    phases = New String() {"Vapor", "Liquid2"}
                    statuses = New eCapePhaseStatus() {eCapePhaseStatus.CAPE_ATEQUILIBRIUM, eCapePhaseStatus.CAPE_ATEQUILIBRIUM}
                ElseIf Not vok And l1ok And l2ok Then
                    phases = New String() {"Liquid", "Liquid2"}
                    statuses = New eCapePhaseStatus() {eCapePhaseStatus.CAPE_ATEQUILIBRIUM, eCapePhaseStatus.CAPE_ATEQUILIBRIUM}
                ElseIf vok And Not l1ok And Not l2ok Then
                    phases = New String() {"Vapor"}
                    statuses = New eCapePhaseStatus() {eCapePhaseStatus.CAPE_ATEQUILIBRIUM}
                ElseIf Not vok And l1ok And Not l2ok Then
                    phases = New String() {"Liquid"}
                    statuses = New eCapePhaseStatus() {eCapePhaseStatus.CAPE_ATEQUILIBRIUM}
                ElseIf vok And Not l1ok And sok Then
                    phases = New String() {"Vapor", "Solid"}
                    statuses = New eCapePhaseStatus() {eCapePhaseStatus.CAPE_ATEQUILIBRIUM, eCapePhaseStatus.CAPE_ATEQUILIBRIUM}
                ElseIf Not vok And l1ok And sok Then
                    phases = New String() {"Liquid", "Solid"}
                    statuses = New eCapePhaseStatus() {eCapePhaseStatus.CAPE_ATEQUILIBRIUM, eCapePhaseStatus.CAPE_ATEQUILIBRIUM}
                ElseIf Not vok And Not l1ok And l2ok Then
                    phases = New String() {"Liquid2"}
                    statuses = New eCapePhaseStatus() {eCapePhaseStatus.CAPE_ATEQUILIBRIUM}
                ElseIf Not vok And Not l1ok And sok Then
                    phases = New String() {"Solid"}
                    statuses = New eCapePhaseStatus() {eCapePhaseStatus.CAPE_ATEQUILIBRIUM}
                End If

                mo.SetPresentPhases(phases, statuses)

                Dim nc As Integer = ms.Fases(0).Componentes.Count
                T = ms.Fases(0).SPMProperties.temperature.GetValueOrDefault
                P = ms.Fases(0).SPMProperties.pressure.GetValueOrDefault

                Dim vz(nc - 1), pf As Double, i As Integer

                mo.SetOverallProp("temperature", Nothing, New Double() {T})
                mo.SetOverallProp("pressure", Nothing, New Double() {P})

                If vok Then

                    i = 0
                    For Each s As Substancia In ms.Fases(2).Componentes.Values
                        vz(i) = s.FracaoMolar.GetValueOrDefault
                        i += 1
                    Next
                    pf = ms.Fases(2).SPMProperties.molarfraction.GetValueOrDefault

                    mo.SetSinglePhaseProp("temperature", "Vapor", Nothing, New Double() {T})
                    mo.SetSinglePhaseProp("pressure", "Vapor", Nothing, New Double() {P})
                    mo.SetSinglePhaseProp("phasefraction", "Vapor", "Mole", New Double() {pf})
                    mo.SetSinglePhaseProp("fraction", "Vapor", "Mole", vz)

                End If

                If l1ok Then

                    i = 0
                    For Each s As Substancia In ms.Fases(3).Componentes.Values
                        vz(i) = s.FracaoMolar.GetValueOrDefault
                        i += 1
                    Next
                    pf = ms.Fases(3).SPMProperties.molarfraction.GetValueOrDefault

                    mo.SetSinglePhaseProp("temperature", "Liquid", Nothing, New Double() {T})
                    mo.SetSinglePhaseProp("pressure", "Liquid", Nothing, New Double() {P})
                    mo.SetSinglePhaseProp("phasefraction", "Liquid", "Mole", New Double() {pf})
                    mo.SetSinglePhaseProp("fraction", "Liquid", "Mole", vz)

                End If

                If l2ok Then

                    i = 0
                    For Each s As Substancia In ms.Fases(4).Componentes.Values
                        vz(i) = s.FracaoMolar.GetValueOrDefault
                        i += 1
                    Next
                    pf = ms.Fases(4).SPMProperties.molarfraction.GetValueOrDefault

                    mo.SetSinglePhaseProp("temperature", "Liquid2", Nothing, New Double() {T})
                    mo.SetSinglePhaseProp("pressure", "Liquid2", Nothing, New Double() {P})
                    mo.SetSinglePhaseProp("phasefraction", "Liquid2", "Mole", New Double() {pf})
                    mo.SetSinglePhaseProp("fraction", "Liquid2", "Mole", vz)

                End If

                If sok Then

                    i = 0
                    For Each s As Substancia In ms.Fases(7).Componentes.Values
                        vz(i) = s.FracaoMolar.GetValueOrDefault
                        i += 1
                    Next
                    pf = ms.Fases(7).SPMProperties.molarfraction.GetValueOrDefault

                    mo.SetSinglePhaseProp("temperature", "Solid", Nothing, New Double() {T})
                    mo.SetSinglePhaseProp("pressure", "Solid", Nothing, New Double() {P})
                    mo.SetSinglePhaseProp("phasefraction", "Solid", "Mole", New Double() {pf})
                    mo.SetSinglePhaseProp("fraction", "Solid", "Mole", vz)

                End If

            End If

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
        Public Overridable Function CheckEquilibriumSpec(ByVal specification1 As Object, ByVal specification2 As Object, ByVal solutionType As String) As Boolean Implements ICapeThermoEquilibriumRoutine.CheckEquilibriumSpec
            If specification1(0).ToString.ToLower = "temperature" And specification2(0).ToString.ToLower = "pressure" Then
                Return True
            ElseIf specification1(0).ToString.ToLower = "pressure" And specification2(0).ToString.ToLower = "enthalpy" Then
                Return True
            ElseIf specification1(0).ToString.ToLower = "pressure" And specification2(0).ToString.ToLower = "entropy" Then
                Return True
            ElseIf specification1(0).ToString.ToLower = "pressure" And specification2(0).ToString.ToLower = "phasefraction" Then
                Return True
            ElseIf specification1(0).ToString.ToLower = "temperature" And specification2(0).ToString.ToLower = "phasefraction" Then
                Return True
            ElseIf specification2(0).ToString.ToLower = "temperature" And specification1(0).ToString.ToLower = "pressure" Then
                Return True
            ElseIf specification2(0).ToString.ToLower = "pressure" And specification1(0).ToString.ToLower = "enthalpy" Then
                Return True
            ElseIf specification2(0).ToString.ToLower = "pressure" And specification1(0).ToString.ToLower = "entropy" Then
                Return True
            ElseIf specification2(0).ToString.ToLower = "pressure" And specification1(0).ToString.ToLower = "phasefraction" Then
                Return True
            ElseIf specification2(0).ToString.ToLower = "temperature" And specification1(0).ToString.ToLower = "phasefraction" Then
                Return True
            Else
                Return False
            End If
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
        Public Overridable Sub SetMaterial(ByVal material As Object) Implements ICapeThermoMaterialContext.SetMaterial
            Me.CurrentMaterialStream = Nothing
            If _como IsNot Nothing Then
                If System.Runtime.InteropServices.Marshal.IsComObject(_como) Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_como)
                End If
            End If
            If My.Application.CAPEOPENMode Then
                _como = material
                If TryCast(material, MaterialStream) Is Nothing Then
                    Me.CurrentMaterialStream = COMaterialtoDWMaterial(material)
                Else
                    Me.CurrentMaterialStream = material
                End If
            Else
                Me.CurrentMaterialStream = material
            End If
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
        Public Overridable Sub UnsetMaterial() Implements ICapeThermoMaterialContext.UnsetMaterial
            Me.CurrentMaterialStream = Nothing
            If My.Application.CAPEOPENMode Then
                If _como IsNot Nothing Then
                    If System.Runtime.InteropServices.Marshal.IsComObject(_como) Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_como)
                    End If
                End If
            End If
        End Sub

        ''' <summary>
        ''' Converts a COM Material Object into a DWSIM Material Stream.
        ''' </summary>
        ''' <param name="material">The Material Object to convert from</param>
        ''' <returns>A DWSIM Material Stream</returns>
        ''' <remarks>This function is called by SetMaterial when DWSIM Property Packages are working in outside environments (CAPE-OPEN COSEs) like COCO/COFE.</remarks>
        Friend Function COMaterialtoDWMaterial(ByVal material As Object) As MaterialStream

            Dim ms As New MaterialStream(CType(material, ICapeIdentification).ComponentName, "")
            For Each phase As DTL.ClassesBasicasTermodinamica.Fase In ms.Fases.Values
                For Each tmpcomp As ConstantProperties In _selectedcomps.Values
                    phase.Componentes.Add(tmpcomp.Name, New DTL.ClassesBasicasTermodinamica.Substancia(tmpcomp.Name, ""))
                    phase.Componentes(tmpcomp.Name).ConstantProperties = tmpcomp
                Next
            Next

            'transfer values

            Dim mys As ICapeThermoMaterial = material

            Dim Tv, Tl1, Tl2, Pv, Pl1, Pl2, xv, xl1, xl2 As Double
            Dim Vz As Object = Nothing
            Dim Vy As Object = Nothing
            Dim Vx1 As Object = Nothing
            Dim Vx2 As Object = Nothing
            Dim Vwy As Object = Nothing
            Dim Vwx1 As Object = Nothing
            Dim Vwx2 As Object = Nothing
            Dim labels As Object = Nothing
            Dim statuses As Object = Nothing
            Dim res As Object = Nothing

            Try
                mys.GetOverallProp("fraction", "Mole", res)
                Vz = res
            Catch ex As Exception
                Vz = RET_NullVector()
            End Try

            mys.GetPresentPhases(labels, statuses)

            Dim data(0) As Double

            Dim i As Integer = 0
            Dim n As Integer = UBound(labels)

            For i = 0 To n
                If statuses(i) = CapeOpen.eCapePhaseStatus.CAPE_ATEQUILIBRIUM Then
                    Select Case labels(i)
                        Case "Vapor"
                            mys.GetTPFraction(labels(i), Tv, Pv, Vy)
                        Case "Liquid"
                            mys.GetTPFraction(labels(i), Tl1, Pl1, Vx1)
                        Case "Liquid2"
                            mys.GetTPFraction(labels(i), Tl2, Pl2, Vx2)
                    End Select
                    Select Case labels(i)
                        Case "Vapor"
                            mys.GetSinglePhaseProp("phasefraction", labels(i), "Mole", res)
                            xv = res(0)
                        Case "Liquid"
                            mys.GetSinglePhaseProp("phasefraction", labels(i), "Mole", res)
                            xl1 = res(0)
                        Case "Liquid2"
                            mys.GetSinglePhaseProp("phasefraction", labels(i), "Mole", res)
                            xl2 = res(0)
                    End Select
                Else
                    Select Case labels(i)
                        Case "Vapor"
                            xv = 0.0#
                        Case "Liquid"
                            xl1 = 0.0#
                        Case "Liquid2"
                            xl2 = 0.0#
                    End Select
                End If
            Next

            Me.CurrentMaterialStream = ms

            'copy fractions

            With ms
                i = 0
                For Each s As Substancia In .Fases(0).Componentes.Values
                    s.FracaoMolar = Vz(i)
                    i += 1
                Next
                If Vy IsNot Nothing Then
                    i = 0
                    For Each s As Substancia In .Fases(2).Componentes.Values
                        s.FracaoMolar = Vy(i)
                        i += 1
                    Next
                    Vwy = Me.AUX_CONVERT_MOL_TO_MASS(Vy)
                    i = 0
                    For Each s As Substancia In .Fases(2).Componentes.Values
                        s.FracaoMassica = Vwy(i)
                        i += 1
                    Next
                Else
                    i = 0
                    For Each s As Substancia In .Fases(2).Componentes.Values
                        s.FracaoMolar = 0.0#
                        i += 1
                    Next
                    i = 0
                    For Each s As Substancia In .Fases(2).Componentes.Values
                        s.FracaoMassica = 0.0#
                        i += 1
                    Next
                End If
                .Fases(2).SPMProperties.molarfraction = xv
                If Vx1 IsNot Nothing Then
                    i = 0
                    For Each s As Substancia In .Fases(3).Componentes.Values
                        s.FracaoMolar = Vx1(i)
                        i += 1
                    Next
                    Vwx1 = Me.AUX_CONVERT_MOL_TO_MASS(Vx1)
                    i = 0
                    For Each s As Substancia In .Fases(3).Componentes.Values
                        s.FracaoMassica = Vwx1(i)
                        i += 1
                    Next
                Else
                    i = 0
                    For Each s As Substancia In .Fases(3).Componentes.Values
                        s.FracaoMolar = 0.0#
                        i += 1
                    Next
                    i = 0
                    For Each s As Substancia In .Fases(3).Componentes.Values
                        s.FracaoMassica = 0.0#
                        i += 1
                    Next
                End If
                .Fases(3).SPMProperties.molarfraction = xl1
                If Vx2 IsNot Nothing Then
                    i = 0
                    For Each s As Substancia In .Fases(4).Componentes.Values
                        s.FracaoMolar = Vx2(i)
                        i += 1
                    Next
                    Vwx2 = Me.AUX_CONVERT_MOL_TO_MASS(Vx2)
                    i = 0
                    For Each s As Substancia In .Fases(4).Componentes.Values
                        s.FracaoMassica = Vwx2(i)
                        i += 1
                    Next
                Else
                    i = 0
                    For Each s As Substancia In .Fases(4).Componentes.Values
                        s.FracaoMolar = 0.0#
                        i += 1
                    Next
                    i = 0
                    For Each s As Substancia In .Fases(4).Componentes.Values
                        s.FracaoMassica = 0.0#
                        i += 1
                    Next
                End If
                .Fases(4).SPMProperties.molarfraction = xl2
            End With

            Return ms

        End Function

#End Region

#Region "   CAPE-OPEN ICapeUtilities Implementation"

        Friend _availablecomps As Dictionary(Of String, ConstantProperties)
        Friend _selectedcomps As Dictionary(Of String, ConstantProperties)

        Public ReadOnly Property AvailableCompounds As Dictionary(Of String, ConstantProperties)
            Get
                Return _availablecomps
            End Get
        End Property

        Public ReadOnly Property SelectedCompounds As Dictionary(Of String, ConstantProperties)
            Get
                Return _selectedcomps
            End Get
        End Property

        <System.NonSerialized()> Friend _pme As Object

        ''' <summary>
        ''' The PMC displays its user interface and allows the Flowsheet User to interact with it. If no user interface is
        ''' available it returns an error.</summary>
        ''' <remarks></remarks>
        Public Overridable Sub Edit() Implements CapeOpen.ICapeUtilities.Edit

            'do nothing

        End Sub

        ''' <summary>
        ''' Initially, this method was only present in the ICapeUnit interface. Since ICapeUtilities.Initialize is now
        ''' available for any kind of PMC, ICapeUnit. Initialize is deprecated.
        ''' The PME will order the PMC to get initialized through this method. Any initialisation that could fail must be
        ''' placed here. Initialize is guaranteed to be the first method called by the client (except low level methods such
        ''' as class constructors or initialization persistence methods). Initialize has to be called once when the PMC is
        ''' instantiated in a particular flowsheet.
        ''' When the initialization fails, before signalling an error, the PMC must free all the resources that were
        ''' allocated before the failure occurred. When the PME receives this error, it may not use the PMC anymore.
        ''' The method terminate of the current interface must not either be called. Hence, the PME may only release
        ''' the PMC through the middleware native mechanisms.
        ''' </summary>
        ''' <remarks></remarks>
        Public Overridable Sub Initialize() Implements CapeOpen.ICapeUtilities.Initialize

            Me.m_ip = New DataTable
            Me.m_props = New DTL.SimulationObjects.PropertyPackages.Auxiliary.PROPS
            ConfigParameters()

            'load Henry Coefficients
       
            Using stream As IO.Stream = New IO.MemoryStream(My.Resources.Henry)
                Using reader As TextReader = New IO.StreamReader(stream)
                    reader.ReadLine()
                    reader.ReadLine()
                    Dim line As String = ""
                    While line = reader.ReadLine() <> Nothing
                        Dim HP As New HenryParam
                        HP.Component = line.Split(";")(1)
                        HP.CAS = line.Split(";")(2)
                        HP.KHcp = Val(line.Split(";")(3))
                        HP.C = Val(line.Split(";")(4))
                        m_Henry.Add(HP.CAS, HP)
                    End While
                End Using
            End Using

        End Sub

        ''' <summary>
        ''' Returns an ICapeCollection interface.
        ''' </summary>
        ''' <value></value>
        ''' <returns>CapeInterface (ICapeCollection)</returns>
        ''' <remarks>This interface will contain a collection of ICapeParameter interfaces.
        ''' This method allows any client to access all the CO Parameters exposed by a PMC. Initially, this method was
        ''' only present in the ICapeUnit interface. Since ICapeUtilities.GetParameters is now available for any kind of
        ''' PMC, ICapeUnit.GetParameters is deprecated. Consult the “Open Interface Specification: Parameter
        ''' Common Interface” document for more information about parameter. Consult the “Open Interface
        ''' Specification: Collection Common Interface” document for more information about collection.
        ''' If the PMC does not support exposing its parameters, it should raise the ECapeNoImpl error, instead of
        ''' returning a NULL reference or an empty Collection. But if the PMC supports parameters but has for this call
        ''' no parameters, it should return a valid ICapeCollection reference exposing zero parameters.</remarks>
        Public Overridable ReadOnly Property parameters1() As Object Implements CapeOpen.ICapeUtilities.parameters
            Get
                Throw New NotImplementedException
            End Get
        End Property

        ''' <summary>
        ''' Allows the PME to convey the PMC a reference to the former’s simulation context. 
        ''' </summary>
        ''' <value>The reference to the PME’s simulation context class. For the PMC to
        ''' use this class, this reference will have to be converted to each of the
        ''' defined CO Simulation Context interfaces.</value>
        ''' <remarks>The simulation context
        ''' will be PME objects which will expose a given set of CO interfaces. Each of these interfaces will allow the
        ''' PMC to call back the PME in order to benefit from its exposed services (such as creation of material
        ''' templates, diagnostics or measurement unit conversion). If the PMC does not support accessing the
        ''' simulation context, it is recommended to raise the ECapeNoImpl error.
        ''' Initially, this method was only present in the ICapeUnit interface. Since ICapeUtilities.SetSimulationContext
        ''' is now available for any kind of PMC, ICapeUnit. SetSimulationContext is deprecated.</remarks>
        Public Overridable WriteOnly Property simulationContext() As Object Implements CapeOpen.ICapeUtilities.simulationContext
            Set(ByVal value As Object)
                _pme = value
            End Set
        End Property

        ''' <summary>
        ''' Initially, this method was only present in the ICapeUnit interface. Since ICapeUtilities.Terminate is now
        ''' available for any kind of PMC, ICapeUnit.Terminate is deprecated.
        ''' The PME will order the PMC to get destroyed through this method. Any uninitialization that could fail must
        ''' be placed here. ‘Terminate’ is guaranteed to be the last method called by the client (except low level methods
        ''' such as class destructors). ‘Terminate’ may be called at any time, but may be only called once.
        ''' When this method returns an error, the PME should report the user. However, after that the PME is not
        ''' allowed to use the PMC anymore.
        ''' The Unit specification stated that “Terminate may check if the data has been saved and return an error if
        ''' not.” It is suggested not to follow this recommendation, since it’s the PME responsibility to save the state of
        ''' the PMC before terminating it. In the case that a user wants to close a simulation case without saving it, it’s
        ''' better to leave the PME to handle the situation instead of each PMC providing a different implementation.
        ''' </summary>
        ''' <remarks></remarks>
        Public Overridable Sub Terminate() Implements CapeOpen.ICapeUtilities.Terminate
            Me.CurrentMaterialStream = Nothing
            If _como IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(_como)
            Me.Dispose()
        End Sub

#End Region

#Region "   CAPE-OPEN Error Interfaces"

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

        Public ReadOnly Property scope() As String Implements CapeOpen.ECapeUser.scope
            Get
                Return _scope
            End Get
        End Property

#End Region

#Region "   IDisposable Support "
        Private disposedValue As Boolean = False        ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: free other state (managed objects).
                End If

                ' TODO: free your own state (unmanaged objects).
                If _como IsNot Nothing Then
                    If System.Runtime.InteropServices.Marshal.IsComObject(_como) Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_como)
                    Else
                        _como = Nothing
                    End If
                End If
                If _pme IsNot Nothing Then
                    If System.Runtime.InteropServices.Marshal.IsComObject(_pme) Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(_pme)
                    Else
                        _pme = Nothing
                    End If
                End If


                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

#Region " IDisposable Support "
        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Overridable Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

#End Region

    End Class

    ''' <summary>
    ''' Class to store Phase Info and mapping for CAPE-OPEN Property Packages
    ''' </summary>
    ''' <remarks>Used only in the context of CAPE-OPEN Objects.</remarks>
    <System.Serializable()> Public Class PhaseInfo

        Public PhaseLabel As String = ""
        Public DWPhaseIndex As Integer
        Public DWPhaseID As Fase = Fase.Aqueous

        Sub New(ByVal pl As String, ByVal pi As Integer, ByVal pid As Fase)
            PhaseLabel = pl
            DWPhaseIndex = pi
            DWPhaseID = pid
        End Sub

    End Class

    ''' <summary>
    ''' COM IStream Class Implementation
    ''' </summary>
    ''' <remarks></remarks>
    <System.Serializable()> Public Class ComStreamWrapper
        Inherits System.IO.Stream
        Private mSource As IStream
        Private mInt64 As IntPtr

        Public Sub New(ByVal source As IStream)
            mSource = source
            mInt64 = iop.Marshal.AllocCoTaskMem(8)
        End Sub

        Protected Overrides Sub Finalize()
            Try
                iop.Marshal.FreeCoTaskMem(mInt64)
            Finally
                MyBase.Finalize()
            End Try
        End Sub

        Public Overrides ReadOnly Property CanRead() As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property CanSeek() As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property CanWrite() As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides Sub Flush()
            mSource.Commit(0)
        End Sub

        Public Overrides ReadOnly Property Length() As Long
            Get
                Dim stat As ComTypes.STATSTG
                stat = Nothing
                mSource.Stat(stat, 1)
                Return stat.cbSize
            End Get
        End Property

        Public Overrides Property Position() As Long
            Get
                Throw New NotImplementedException()
            End Get
            Set(ByVal value As Long)
                Throw New NotImplementedException()
            End Set
        End Property

        Public Overrides Function Read(ByVal buffer As Byte(), ByVal offset As Integer, ByVal count As Integer) As Integer
            If offset <> 0 Then
                Throw New NotImplementedException()
            End If
            mSource.Read(buffer, count, mInt64)
            Return iop.Marshal.ReadInt32(mInt64)
        End Function

        Public Overrides Function Seek(ByVal offset As Long, ByVal origin As System.IO.SeekOrigin) As Long
            mSource.Seek(offset, CInt(origin), mInt64)
            Return iop.Marshal.ReadInt64(mInt64)
        End Function

        Public Overrides Sub SetLength(ByVal value As Long)
            mSource.SetSize(value)
        End Sub

        Public Overrides Sub Write(ByVal buffer As Byte(), ByVal offset As Integer, ByVal count As Integer)
            If offset <> 0 Then
                Throw New NotImplementedException()
            End If
            mSource.Write(buffer, count, IntPtr.Zero)
        End Sub

    End Class

    <System.Serializable()> Public Class HenryParam
        Public Component As String
        Public CAS As String
        Public KHcp As Double
        Public C As Double
    End Class

End Namespace
