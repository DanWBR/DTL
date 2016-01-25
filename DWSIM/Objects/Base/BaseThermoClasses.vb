'    Basic Thermodynamic Classes for DWSIM
'    Copyright 2008 Daniel Wagner O. de Medeiros
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

Imports System.Xml.Serialization
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Runtime.Serialization
Imports System.IO

Namespace DTL.BaseThermoClasses

    <Serializable()> Public Class Substance

        Protected m_componentdescription As String = ""
        Protected m_componentname As String = ""
        Protected m_molarfraction As Double? = 0
        Protected m_massfraction As Double? = 0
        Protected m_molarflow As Double? = 0
        Protected m_massflow As Double? = 0
        Protected m_fugacitycoeff As Double? = 0
        Protected m_activitycoeff As Double? = 0
        Protected m_partialvolume As Double? = 0
        Protected m_partialpressure As Double? = 0
        Protected m_volumetricflow As Double? = 0
        Protected m_volumetricfraction As Double? = 0
        Protected m_isPF As Boolean = False
        Protected m_lnKval As Double = 0
        Protected m_Kval As Double = 0

        Public Property lnKvalue() As Double
            Get
                Return m_lnKval
            End Get
            Set(ByVal value As Double)
                m_lnKval = value
            End Set
        End Property

        Public Property Kvalue() As Double
            Get
                Return m_Kval
            End Get
            Set(ByVal value As Double)
                m_Kval = value
            End Set
        End Property

        Public Property PetroleumFraction() As Boolean
            Get
                Return m_isPF
            End Get
            Set(ByVal value As Boolean)
                m_isPF = value
            End Set
        End Property

        Public Property MolarFraction() As Double?
            Get
                Return m_molarfraction
            End Get
            Set(ByVal value As Double?)
                m_molarfraction = value
            End Set
        End Property

        Public Property MassFraction() As Double?
            Get
                Return m_massfraction
            End Get
            Set(ByVal value As Double?)
                m_massfraction = value
            End Set
        End Property

        Public Property MolarFlow() As Double?
            Get
                Return Me.m_molarflow
            End Get
            Set(ByVal value As Double?)
                Me.m_molarflow = value
            End Set
        End Property

        Public Property MassFlow() As Double?
            Get
                Return Me.m_massflow
            End Get
            Set(ByVal value As Double?)
                Me.m_massflow = value
            End Set
        End Property

        Public Property FugacityCoeff() As Double?
            Get
                Return Me.m_fugacitycoeff
            End Get
            Set(ByVal value As Double?)
                Me.m_fugacitycoeff = value
            End Set
        End Property

        Public Property ActivityCoeff() As Double?
            Get
                Return Me.m_activitycoeff
            End Get
            Set(ByVal value As Double?)
                Me.m_activitycoeff = value
            End Set
        End Property

        Public Property PartialVolume() As Double?
            Get
                Return Me.m_partialvolume
            End Get
            Set(ByVal value As Double?)
                Me.m_partialvolume = value
            End Set
        End Property

        Public Property PartialPressure() As Double?
            Get
                Return Me.m_partialpressure
            End Get
            Set(ByVal value As Double?)
                Me.m_partialpressure = value
            End Set
        End Property

        Public Property VolumetricFlow() As Double?
            Get
                Return Me.m_volumetricflow
            End Get
            Set(ByVal value As Double?)
                Me.m_volumetricflow = value
            End Set
        End Property

        Public Property VolumetricFraction() As Double?
            Get
                Return Me.m_volumetricfraction
            End Get
            Set(ByVal value As Double?)
                Me.m_volumetricfraction = value
            End Set
        End Property

        Public Property Description() As String
            Get
                Return m_componentdescription
            End Get
            Set(ByVal value As String)
                m_componentdescription = value
            End Set
        End Property

        Public Property Name() As String
            Get
                Return m_componentname
            End Get
            Set(ByVal value As String)
                m_componentname = value
            End Set
        End Property

        Public TDProperties As New TemperatureDependentProperties
        Public PDProperties As New PressureDependentProperties
        Public ConstantProperties As New ConstantProperties

        Public Sub New(ByVal name As String, ByVal description As String)

            Me.m_componentname = name
            Me.m_componentdescription = description

        End Sub

    End Class

    <Serializable()> Public Class Phase

        Protected m_ComponentDescription As String
        Protected m_ComponentName As String

        Public Property Description() As String
            Get
                Return m_ComponentDescription
            End Get
            Set(ByVal value As String)
                m_ComponentDescription = value
            End Set
        End Property

        Public Property Name() As String
            Get
                Return m_ComponentName
            End Get
            Set(ByVal value As String)
                m_ComponentName = value
            End Set
        End Property

        Public Components As Dictionary(Of String, Substance)

        Public SPMProperties As New SinglePhaseMixtureProperties
        Public TPMProperties As New TwoPhaseMixtureProperties

        Public Sub New(ByVal name As String, ByVal description As String)

            Me.m_ComponentName = name
            Me.m_ComponentDescription = description
            Me.Components = New Dictionary(Of String, Substance)

        End Sub

        Public Sub New(ByVal name As String, ByVal description As String, ByVal Substances As Dictionary(Of String, Substance))

            Me.m_ComponentName = name
            Me.m_ComponentDescription = description
            Me.Components = Substances

        End Sub

    End Class

    Public Enum PhaseName
        Liquid
        Vapor
        Mixture
    End Enum

    Public Enum ReactionType
        Equilibrium
        Kinetic
        Heterogeneous_Catalytic
        Conversion
    End Enum

    Public Enum ReactionBasis
        Activity
        Fugacity
        MolarConc
        MassConc
        MolarFrac
        MassFrac
        PartialPress
    End Enum

#Region "Subclasses"

    <Serializable()> Public Class TemperatureDependentProperties

        Protected tdp_idealGasHeatCapacity As Double? = Nothing
        Protected tdp_surfaceTension As Double? = Nothing
        Protected tdp_thermalConductivityOfLiquid As Double? = Nothing
        Protected tdp_thermalConductivityOfVapor As Double? = Nothing
        Protected tdp_vaporPressure As Double? = Nothing
        Protected tdp_viscosityOfLiquid As Double? = Nothing
        Protected tdp_viscosityOfVapor As Double? = Nothing

        Public Property idealGasHeatCapacity() As Double?
            Get
                Return tdp_idealGasHeatCapacity
            End Get
            Set(ByVal value As Double?)
                tdp_idealGasHeatCapacity = value
            End Set
        End Property

        Public Property thermalConductivityOfLiquid() As Double?
            Get
                Return tdp_thermalConductivityOfLiquid
            End Get
            Set(ByVal value As Double?)
                tdp_thermalConductivityOfLiquid = value
            End Set
        End Property

        Public Property thermalConductivityOfVapor() As Double?
            Get
                Return tdp_thermalConductivityOfVapor
            End Get
            Set(ByVal value As Double?)
                tdp_thermalConductivityOfVapor = value
            End Set
        End Property

        Public Property vaporPressure() As Double?
            Get
                Return tdp_vaporPressure
            End Get
            Set(ByVal value As Double?)
                tdp_vaporPressure = value
            End Set
        End Property

        Public Property viscosityOfLiquid() As Double?
            Get
                Return tdp_viscosityOfLiquid
            End Get
            Set(ByVal value As Double?)
                tdp_viscosityOfLiquid = value
            End Set
        End Property

        Public Property viscosityOfVapor() As Double?
            Get
                Return tdp_viscosityOfVapor
            End Get
            Set(ByVal value As Double?)
                tdp_viscosityOfVapor = value
            End Set
        End Property

        Public Property surfaceTension() As Double?
            Get
                Return tdp_surfaceTension
            End Get
            Set(ByVal value As Double?)
                tdp_surfaceTension = value
            End Set
        End Property

    End Class

    <Serializable()> Public Class PressureDependentProperties

        Protected pdp_boilingPointTemperature As Double? = Nothing
        Public Property boilingPointTemperature() As Double?
            Get
                Return pdp_boilingPointTemperature
            End Get
            Set(ByVal value As Double?)
                pdp_boilingPointTemperature = value
            End Set
        End Property

        Protected pdp_meltingTemperature As Double? = Nothing
        Public Property meltingTemperature() As Double?
            Get
                Return pdp_meltingTemperature
            End Get
            Set(ByVal value As Double?)
                pdp_meltingTemperature = value
            End Set
        End Property

    End Class

    <Serializable()> Public Class SinglePhaseMixtureProperties

        Public Property osmoticCoefficient As Double?
        Public Property freezingPointDepression As Double?
        Public Property freezingPoint As Double?
        Public Property ionicStrength As Double?
        Public Property pH As Double?
        Protected _dewtemperature As Double? = Nothing
        Public Property dewTemperature() As Double?
            Get
                Return _dewtemperature
            End Get
            Set(ByVal value As Double?)
                _dewtemperature = value
            End Set
        End Property

        Protected _dewpressure As Double? = Nothing
        Public Property dewPressure() As Double?
            Get
                Return _dewpressure
            End Get
            Set(ByVal value As Double?)
                _dewpressure = value
            End Set
        End Property

        Protected _bubbletemperature As Double? = Nothing
        Public Property bubbleTemperature() As Double?
            Get
                Return _bubbletemperature
            End Get
            Set(ByVal value As Double?)
                _bubbletemperature = value
            End Set
        End Property

        Protected _bubblepressure As Double? = Nothing
        Public Property bubblePressure() As Double?
            Get
                Return _bubblepressure
            End Get
            Set(ByVal value As Double?)
                _bubblepressure = value
            End Set
        End Property

        Protected spmp_activity As Double? = Nothing
        Public Property activity() As Double?
            Get
                Return spmp_activity
            End Get
            Set(ByVal value As Double?)
                spmp_activity = value
            End Set
        End Property

        Protected spmp_activityCoefficient As Double? = Nothing
        Public Property activityCoefficient() As Double?
            Get
                Return spmp_activityCoefficient
            End Get
            Set(ByVal value As Double?)
                spmp_activityCoefficient = value
            End Set
        End Property

        Protected spmp_compressibility As Double? = Nothing
        Public Property compressibility() As Double?
            Get
                Return spmp_compressibility
            End Get
            Set(ByVal value As Double?)
                spmp_compressibility = value
            End Set
        End Property

        Protected spmp_compressibilityFactor As Double? = Nothing
        Public Property compressibilityFactor() As Double?
            Get
                Return spmp_compressibilityFactor
            End Get
            Set(ByVal value As Double?)
                spmp_compressibilityFactor = value
            End Set
        End Property

        Protected spmp_density As Double? = Nothing
        Public Property density() As Double?
            Get
                Return spmp_density
            End Get
            Set(ByVal value As Double?)
                spmp_density = value
            End Set
        End Property

        Protected spmp_enthalpy As Double? = Nothing
        Public Property enthalpy() As Double?
            Get
                Return spmp_enthalpy
            End Get
            Set(ByVal value As Double?)
                spmp_enthalpy = value
            End Set
        End Property

        Protected spmp_entropy As Double? = Nothing
        Public Property entropy() As Double?
            Get
                Return spmp_entropy
            End Get
            Set(ByVal value As Double?)
                spmp_entropy = value
            End Set
        End Property

        Protected spmp_enthalpyF As Double? = Nothing
        Public Property enthalpyF() As Double?
            Get
                Return spmp_enthalpyF
            End Get
            Set(ByVal value As Double?)
                spmp_enthalpyF = value
            End Set
        End Property

        Protected spmp_entropyF As Double? = Nothing
        Public Property entropyF() As Double?
            Get
                Return spmp_entropyF
            End Get
            Set(ByVal value As Double?)
                spmp_entropyF = value
            End Set
        End Property

        Protected spmp_excessEnthalpy As Double? = Nothing
        Public Property excessEnthalpy() As Double?
            Get
                Return spmp_excessEnthalpy
            End Get
            Set(ByVal value As Double?)
                spmp_excessEnthalpy = value
            End Set
        End Property

        Protected spmp_excessEntropy As Double? = Nothing
        Public Property excessEntropy() As Double?
            Get
                Return spmp_excessEntropy
            End Get
            Set(ByVal value As Double?)
                spmp_excessEntropy = value
            End Set
        End Property

        Protected spmp_molarflow As Double? = Nothing
        Public Property molarflow() As Double?
            Get
                Return spmp_molarflow
            End Get
            Set(ByVal value As Double?)
                spmp_molarflow = value
            End Set
        End Property

        Protected spmp_massflow As Double? = Nothing
        Public Property massflow() As Double?
            Get
                Return spmp_massflow
            End Get
            Set(ByVal value As Double?)
                spmp_massflow = value
            End Set
        End Property

        Protected spmp_molarfraction As Double? = Nothing
        Public Property molarfraction() As Double?
            Get
                Return spmp_molarfraction
            End Get
            Set(ByVal value As Double?)
                spmp_molarfraction = value
            End Set
        End Property

        Protected spmp_massfraction As Double? = Nothing
        Public Property massfraction() As Double?
            Get
                Return spmp_massfraction
            End Get
            Set(ByVal value As Double?)
                spmp_massfraction = value
            End Set
        End Property

        Protected spmp_fugacity As Double? = Nothing
        Public Property fugacity() As Double?
            Get
                Return spmp_fugacity
            End Get
            Set(ByVal value As Double?)
                spmp_fugacity = value
            End Set
        End Property

        Protected spmp_fugacityCoefficient As Double? = Nothing
        Public Property fugacityCoefficient() As Double?
            Get
                Return spmp_fugacityCoefficient
            End Get
            Set(ByVal value As Double?)
                spmp_fugacityCoefficient = value
            End Set
        End Property

        Protected spmp_heatCapacityCp As Double? = Nothing
        Public Property heatCapacityCp() As Double?
            Get
                Return spmp_heatCapacityCp
            End Get
            Set(ByVal value As Double?)
                spmp_heatCapacityCp = value
            End Set
        End Property

        Protected spmp_heatCapacityCv As Double? = Nothing
        Public Property heatCapacityCv() As Double?
            Get
                Return spmp_heatCapacityCv
            End Get
            Set(ByVal value As Double?)
                spmp_heatCapacityCv = value
            End Set
        End Property

        Protected spmp_jouleThomsonCoefficient As Double? = Nothing
        Public Property jouleThomsonCoefficient() As Double?
            Get
                Return spmp_jouleThomsonCoefficient
            End Get
            Set(ByVal value As Double?)
                spmp_jouleThomsonCoefficient = value
            End Set
        End Property

        Protected spmp_logFugacityCoefficient As Double? = Nothing
        Public Property logFugacityCoefficient() As Double?
            Get
                Return spmp_logFugacityCoefficient
            End Get
            Set(ByVal value As Double?)
                spmp_logFugacityCoefficient = value
            End Set
        End Property

        Protected spmp_molecularWeight As Double? = Nothing
        Public Property molecularWeight() As Double?
            Get
                Return spmp_molecularWeight
            End Get
            Set(ByVal value As Double?)
                spmp_molecularWeight = value
            End Set
        End Property

        Protected spmp_pressure As Double? = Nothing
        Public Property pressure() As Double?
            Get
                Return spmp_pressure
            End Get
            Set(ByVal value As Double?)
                spmp_pressure = value
            End Set
        End Property

        Protected spmp_temperature As Double? = Nothing
        Public Property temperature() As Double?
            Get
                Return spmp_temperature
            End Get
            Set(ByVal value As Double?)
                spmp_temperature = value
            End Set
        End Property

        Protected spmp_speedOfSound As Double? = Nothing
        Public Property speedOfSound() As Double?
            Get
                Return spmp_speedOfSound
            End Get
            Set(ByVal value As Double?)
                spmp_speedOfSound = value
            End Set
        End Property

        Protected spmp_thermalConductivity As Double? = Nothing
        Public Property thermalConductivity() As Double?
            Get
                Return spmp_thermalConductivity
            End Get
            Set(ByVal value As Double?)
                spmp_thermalConductivity = value
            End Set
        End Property

        Protected spmp_dynamic_viscosity As Double? = Nothing
        Public Property viscosity() As Double?
            Get
                Return spmp_dynamic_viscosity
            End Get
            Set(ByVal value As Double?)
                spmp_dynamic_viscosity = value
            End Set
        End Property

        Protected spmp_kinematic_viscosity As Double? = Nothing
        Public Property kinematic_viscosity() As Double?
            Get
                Return spmp_kinematic_viscosity
            End Get
            Set(ByVal value As Double?)
                spmp_kinematic_viscosity = value
            End Set
        End Property

        Protected spmp_volumetric_flow As Double? = Nothing
        Public Property volumetric_flow() As Double?
            Get
                Return spmp_volumetric_flow
            End Get
            Set(ByVal value As Double?)
                spmp_volumetric_flow = value
            End Set
        End Property

        Protected spmp_molarenthalpy As Double? = Nothing
        Public Property molar_enthalpy() As Double?
            Get
                Return spmp_molarenthalpy
            End Get
            Set(ByVal value As Double?)
                spmp_molarenthalpy = value
            End Set
        End Property

        Protected spmp_molarentropy As Double? = Nothing
        Public Property molar_entropy() As Double?
            Get
                Return spmp_molarentropy
            End Get
            Set(ByVal value As Double?)
                spmp_molarentropy = value
            End Set
        End Property

        Protected spmp_molarenthalpyF As Double? = Nothing
        Public Property molar_enthalpyF() As Double?
            Get
                Return spmp_molarenthalpyF
            End Get
            Set(ByVal value As Double?)
                spmp_molarenthalpyF = value
            End Set
        End Property

        Protected spmp_molarentropyF As Double? = Nothing
        Public Property molar_entropyF() As Double?
            Get
                Return spmp_molarentropyF
            End Get
            Set(ByVal value As Double?)
                spmp_molarentropyF = value
            End Set
        End Property

        Public Sub New()

        End Sub

    End Class

    <Serializable()> Public Class TwoPhaseMixtureProperties

        Protected tpmp_kvalue As Double? = Nothing
        Public Property kvalue() As Double?
            Get
                Return tpmp_kvalue
            End Get
            Set(ByVal value As Double?)
                tpmp_kvalue = value
            End Set
        End Property

        Protected tpmp_logKvalue As Double? = Nothing
        Public Property logKvalue() As Double?
            Get
                Return tpmp_logKvalue
            End Get
            Set(ByVal value As Double?)
                tpmp_logKvalue = value
            End Set
        End Property

        Protected tpmp_surfaceTension As Double? = Nothing
        Public Property surfaceTension() As Double?
            Get
                Return tpmp_surfaceTension
            End Get
            Set(ByVal value As Double?)
                tpmp_surfaceTension = value
            End Set
        End Property

    End Class

    <Serializable()> Public Class InteractionParameter

        Implements ICloneable
        Public Comp1 As String = ""
        Public Comp2 As String = ""
        Public Model As String = ""
        Public DataType As String = ""
        Public Description As String = ""
        Public RegressionFile As String = ""
        Public Parameters As Dictionary(Of String, Object)

        Public Sub New()
            Parameters = New Dictionary(Of String, Object)
        End Sub

        Public Function Clone() As Object Implements System.ICloneable.Clone
            Return ObjectCopy(Me)
        End Function

        Function ObjectCopy(ByVal obj As InteractionParameter) As InteractionParameter

            Dim objMemStream As New MemoryStream(50000)
            Dim objBinaryFormatter As New BinaryFormatter(Nothing, New StreamingContext(StreamingContextStates.Clone))

            objBinaryFormatter.Serialize(objMemStream, obj)

            objMemStream.Seek(0, SeekOrigin.Begin)

            ObjectCopy = objBinaryFormatter.Deserialize(objMemStream)

            objMemStream.Close()

        End Function

    End Class

    <Serializable()> Public Class ConstantProperties

        Implements ICloneable

        Public Name As String = ""
        Public CAS_Number As String = ""
        Public Formula As String = ""
        Public SMILES As String = ""
        Public InChI As String = ""
        Public ChemicalStructure As String = ""
        Public Molar_Weight As Double
        Public Critical_Temperature As Double
        Public Critical_Pressure As Double
        Public Critical_Volume As Double
        Public Critical_Compressibility As Double
        Public Acentric_Factor As Double
        Public Z_Rackett As Double
        Public PR_Volume_Translation_Coefficient As Double
        Public SRK_Volume_Translation_Coefficient As Double
        Public Chao_Seader_Acentricity As Double
        Public Chao_Seader_Solubility_Parameter As Double
        Public Chao_Seader_Liquid_Molar_Volume As Double
        Public IG_Entropy_of_Formation_25C As Double
        Public IG_Enthalpy_of_Formation_25C As Double
        Public IG_Gibbs_Energy_of_Formation_25C As Double
        Public Dipole_Moment As Double
        Public Vapor_Pressure_Constant_A As Double
        Public Vapor_Pressure_Constant_B As Double
        Public Vapor_Pressure_Constant_C As Double
        Public Vapor_Pressure_Constant_D As Double
        Public Vapor_Pressure_Constant_E As Double
        Public Vapor_Pressure_TMIN As Double
        Public Vapor_Pressure_TMAX As Double
        Public Ideal_Gas_Heat_Capacity_Const_A As Double
        Public Ideal_Gas_Heat_Capacity_Const_B As Double
        Public Ideal_Gas_Heat_Capacity_Const_C As Double
        Public Ideal_Gas_Heat_Capacity_Const_D As Double
        Public Ideal_Gas_Heat_Capacity_Const_E As Double
        Public Liquid_Viscosity_Const_A As Double
        Public Liquid_Viscosity_Const_B As Double
        Public Liquid_Viscosity_Const_C As Double
        Public Liquid_Viscosity_Const_D As Double
        Public Liquid_Viscosity_Const_E As Double
        Public Liquid_Density_Const_A As Double
        Public Liquid_Density_Const_B As Double
        Public Liquid_Density_Const_C As Double
        Public Liquid_Density_Const_D As Double
        Public Liquid_Density_Const_E As Double
        Public Liquid_Density_Tmin As Double
        Public Liquid_Density_Tmax As Double
        Public Liquid_Heat_Capacity_Const_A As Double
        Public Liquid_Heat_Capacity_Const_B As Double
        Public Liquid_Heat_Capacity_Const_C As Double
        Public Liquid_Heat_Capacity_Const_D As Double
        Public Liquid_Heat_Capacity_Const_E As Double
        Public Liquid_Heat_Capacity_Tmin As Double
        Public Liquid_Heat_Capacity_Tmax As Double
        Public Liquid_Thermal_Conductivity_Const_A As Double
        Public Liquid_Thermal_Conductivity_Const_B As Double
        Public Liquid_Thermal_Conductivity_Const_C As Double
        Public Liquid_Thermal_Conductivity_Const_D As Double
        Public Liquid_Thermal_Conductivity_Const_E As Double
        Public Liquid_Thermal_Conductivity_Tmin As Double
        Public Liquid_Thermal_Conductivity_Tmax As Double
        Public Vapor_Thermal_Conductivity_Const_A As Double
        Public Vapor_Thermal_Conductivity_Const_B As Double
        Public Vapor_Thermal_Conductivity_Const_C As Double
        Public Vapor_Thermal_Conductivity_Const_D As Double
        Public Vapor_Thermal_Conductivity_Const_E As Double
        Public Vapor_Thermal_Conductivity_Tmin As Double
        Public Vapor_Thermal_Conductivity_Tmax As Double
        Public Vapor_Viscosity_Const_A As Double
        Public Vapor_Viscosity_Const_B As Double
        Public Vapor_Viscosity_Const_C As Double
        Public Vapor_Viscosity_Const_D As Double
        Public Vapor_Viscosity_Const_E As Double
        Public Vapor_Viscosity_Tmin As Double
        Public Vapor_Viscosity_Tmax As Double
        Public Solid_Density_Const_A As Double
        Public Solid_Density_Const_B As Double
        Public Solid_Density_Const_C As Double
        Public Solid_Density_Const_D As Double
        Public Solid_Density_Const_E As Double
        Public Solid_Density_Tmin As Double
        Public Solid_Density_Tmax As Double
        Public Surface_Tension_Const_A As Double
        Public Surface_Tension_Const_B As Double
        Public Surface_Tension_Const_C As Double
        Public Surface_Tension_Const_D As Double
        Public Surface_Tension_Const_E As Double
        Public Surface_Tension_Tmin As Double
        Public Surface_Tension_Tmax As Double
        Public Solid_Heat_Capacity_Const_A As Double
        Public Solid_Heat_Capacity_Const_B As Double
        Public Solid_Heat_Capacity_Const_C As Double
        Public Solid_Heat_Capacity_Const_D As Double
        Public Solid_Heat_Capacity_Const_E As Double
        Public Solid_Heat_Capacity_Tmin As Double
        Public Solid_Heat_Capacity_Tmax As Double
        Public Normal_Boiling_Point As Double
        Public ID As Integer
        Public IsPF As Integer = 0
        Public IsHYPO As Integer = 0
        Public HVap_A As Double
        Public HVap_B As Double
        Public HVap_C As Double
        Public HVap_D As Double
        Public HVap_E As Double
        Public HVap_TMIN As Double
        Public HVap_TMAX As Double
        Public UNIQUAC_R As Double
        Public UNIQUAC_Q As Double

        Public UNIFACGroups As UNIFACGroupCollection
        Public MODFACGroups As UNIFACGroupCollection
        Public NISTMODFACGroups As UNIFACGroupCollection
        Public Elements As New ElementCollection

        Public VaporPressureEquation As String = ""
        Public IdealgasCpEquation As String = ""
        Public LiquidViscosityEquation As String = ""
        Public VaporViscosityEquation As String = ""
        Public VaporizationEnthalpyEquation As String = ""
        Public LiquidDensityEquation As String = ""
        Public LiquidHeatCapacityEquation As String = ""
        Public LiquidThermalConductivityEquation As String = ""
        Public VaporThermalConductivityEquation As String = ""
        Public SolidDensityEquation As String = ""
        Public SolidHeatCapacityEquation As String = ""
        Public SurfaceTensionEquation As String = ""

        Public PC_SAFT_sigma As Double = 0.0#
        Public PC_SAFT_epsilon_k As Double = 0.0#
        Public PC_SAFT_m As Double = 0.0#

        Public PF_MM As Double? = Nothing
        Public NBP As Double? = Nothing
        Public PF_vA As Double? = Nothing
        Public PF_vB As Double? = Nothing
        Public PF_Watson_K As Double? = Nothing
        Public PF_SG As Double? = Nothing
        Public PF_v1 As Double? = Nothing
        Public PF_Tv1 As Double? = Nothing
        Public PF_v2 As Double? = Nothing
        Public PF_Tv2 As Double? = Nothing

        'Databases: 'DWSIM', 'CheResources', 'ChemSep' OR 'User'
        'User databases are XML-serialized versions of this base class, and they may include hypos and pseudos.
        Public OriginalDB As String = "DWSIM"
        Public CurrentDB As String = "DWSIM"

        'COSMO-SAC's database equivalent name
        Public COSMODBName = ""

        'Electrolyte-related properties
        Public IsIon As Boolean = False
        Public IsSalt As Boolean = False
        Public IsHydratedSalt As Boolean = False
        Public HydrationNumber As Double = 0.0#
        Public Charge As Integer = 0
        Public PositiveIon As String = ""
        Public NegativeIon As String = ""
        Public PositiveIonStoichCoeff As Integer = 0
        Public NegativeIonStoichCoeff As Integer = 0
        Public StoichSum As Integer = 0
        Public Electrolyte_DelGF As Double = 0.0#
        Public Electrolyte_DelHF As Double = 0.0#
        Public Electrolyte_Cp0 As Double = 0.0#
        Public TemperatureOfFusion As Double = 0.0#
        Public EnthalpyOfFusionAtTf As Double = 0.0#
        Public SolidTs As Double = 0.0#
        Public SolidDensityAtTs As Double = 0.0#
        Public StandardStateMolarVolume As Double = 0.0#
        Public MolarVolume_v2i As Double = 0.0#
        Public MolarVolume_v3i As Double = 0.0#
        Public MolarVolume_k1i As Double = 0.0#
        Public MolarVolume_k2i As Double = 0.0#
        Public MolarVolume_k3i As Double = 0.0#
        Public Ion_CpAq_a As Double = 0.0#
        Public Ion_CpAq_b As Double = 0.0#
        Public Ion_CpAq_c As Double = 0.0#

        'the following properties are no longer used but kept for compatibility reasons
        <XmlIgnore()> Public UNIFAC_Ri As Double
        <XmlIgnore()> Public UNIFAC_Qi As Double
        <XmlIgnore()> Public UNIFAC_01_CH4 As Integer
        <XmlIgnore()> Public UNIFAC_02_CH3 As Integer
        <XmlIgnore()> Public UNIFAC_03_CH2 As Integer
        <XmlIgnore()> Public UNIFAC_04_CH As Integer
        <XmlIgnore()> Public UNIFAC_05_H2O As Integer
        <XmlIgnore()> Public UNIFAC_06_H2S As Integer
        <XmlIgnore()> Public UNIFAC_07_CO2 As Integer
        <XmlIgnore()> Public UNIFAC_08_N2 As Integer
        <XmlIgnore()> Public UNIFAC_09_O2 As Integer
        <XmlIgnore()> Public UNIFAC_10_OH As Integer
        <XmlIgnore()> Public UNIFAC_11_ACH As Integer
        <XmlIgnore()> Public UNIFAC_12_ACCH2 As Integer
        <XmlIgnore()> Public UNIFAC_13_ACCH3 As Integer
        <XmlIgnore()> Public UNIFAC_14_ACOH As Integer
        <XmlIgnore()> Public UNIFAC_15_CH3CO As Integer
        <XmlIgnore()> Public UNIFAC_16_CH2CO As Integer
        <XmlIgnore()> Public Element_C As Integer = 0
        <XmlIgnore()> Public Element_H As Integer = 0
        <XmlIgnore()> Public Element_O As Integer = 0
        <XmlIgnore()> Public Element_N As Integer = 0
        <XmlIgnore()> Public Element_S As Integer = 0

        Public Sub New()
            UNIFACGroups = New UNIFACGroupCollection
            MODFACGroups = New UNIFACGroupCollection
            NISTMODFACGroups = New UNIFACGroupCollection
        End Sub

        Public Function Clone() As Object Implements System.ICloneable.Clone

            Return ObjectCopy(Me)

        End Function

        Function ObjectCopy(ByVal obj As ConstantProperties) As ConstantProperties

            Dim objMemStream As New MemoryStream(50000)
            Dim objBinaryFormatter As New BinaryFormatter(Nothing, New StreamingContext(StreamingContextStates.Clone))

            objBinaryFormatter.Serialize(objMemStream, obj)

            objMemStream.Seek(0, SeekOrigin.Begin)

            ObjectCopy = objBinaryFormatter.Deserialize(objMemStream)

            objMemStream.Close()

        End Function

    End Class

    <Serializable()> Public Class ConstantPropertiesCollection
        Public Collection() As ConstantProperties
    End Class

    <Serializable()> Public Class UNIFACGroupCollection

        Public Collection As System.Collections.SortedList

        Sub New()
            Collection = New System.Collections.SortedList
        End Sub

    End Class

    <Serializable()> Public Class ElementCollection

        Public Collection As New System.Collections.SortedList

        Sub New()
            Collection = New System.Collections.SortedList
        End Sub
    End Class

#End Region

End Namespace
