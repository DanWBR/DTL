'    Unit System Classes for DWSIM
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

Namespace DTL.UnitsOfMeasure

    <Serializable()> <Xml.Serialization.XmlRoot(ElementName:="Component")> Public Class Units

        Public Name As String

        Public area, distance, time, volume, molar_volume, diameter, thickness, molar_conc, mass_conc, _
                heat_transf_coeff, force, accel, spec_vol, reac_rate, velocity, foulingfactor As String

        Public molar_enthalpy, molar_entropy As String

        Public tdp_idealGasHeatCapacity As String
        Public tdp_surfaceTension As String
        Public tdp_thermalConductivityOfLiquid As String
        Public tdp_thermalConductivityOfVapor As String
        Public tdp_vaporPressure As String
        Public tdp_viscosityOfLiquid As String
        Public tdp_viscosityOfVapor As String

        Public pdp_boilingPointTemperature As String
        Public pdp_meltingTemperature As String

        Public spmp_activity As String
        Public spmp_activityCoefficient As String
        Public spmp_compressibility As String
        Public spmp_compressibilityFactor As String
        Public spmp_density As String
        Public spmp_enthalpy As String
        Public spmp_entropy As String
        Public spmp_excessEnthalpy As String
        Public spmp_excessEntropy As String
        Public spmp_molarflow As String
        Public spmp_massflow As String
        Public spmp_molarfraction As String
        Public spmp_massfraction As String
        Public spmp_fugacity As String
        Public spmp_fugacityCoefficient As String
        Public spmp_heatCapacityCp As String
        Public spmp_heatCapacityCv As String
        Public spmp_jouleThomsonCoefficient As String
        Public spmp_logFugacityCoefficient As String
        Public spmp_molecularWeight As String
        Public spmp_pressure As String
        Public spmp_temperature As String
        Public spmp_speedOfSound As String
        Public spmp_thermalConductivity As String
        Public spmp_viscosity As String
        Public spmp_cinematic_viscosity As String
        Public spmp_volumetricFlow As String
        Public spmp_heatflow As String
        Public spmp_head As String
        Public spmp_deltaT As String
        Public spmp_deltaP As String

        Public tpmp_kvalue As String
        Public tpmp_logKvalue As String
        Public tpmp_surfaceTension As String

        Public Sub New()

        End Sub

    End Class

    <Serializable()> Public Class UnitsSI

        Inherits Units

        Public Sub New()

            With Me

                .Name = App.GetLocalString("SystemSI")

                .accel = "m2/s"
                .area = "m2"
                .diameter = "mm"
                .distance = "m"
                .force = "N"
                .heat_transf_coeff = "W/[m2.K]"
                .mass_conc = "kg/m3"
                .molar_conc = "mol/m3"
                .molar_volume = "m3/kmol"
                .reac_rate = "mol/[m3.s]"
                .spec_vol = "m3/kg"
                .time = "s"
                .volume = "m3"
                .thickness = "mm"
                .molar_enthalpy = "kJ/kmol"
                .molar_entropy = "kJ/[kmol.K]"
                .velocity = "m/s"
                .foulingfactor = "K.m2/W"

                .pdp_boilingPointTemperature = "K"
                .pdp_meltingTemperature = "K"
                .spmp_activity = "Pa"
                .spmp_activityCoefficient = "-"
                .spmp_compressibility = "Pa-1"
                .spmp_compressibilityFactor = "-"
                .spmp_density = "kg/m3"
                .spmp_enthalpy = "kJ/kg"
                .spmp_entropy = "kJ/[kg.K]"
                .spmp_excessEnthalpy = "kJ/kg"
                .spmp_excessEntropy = "kJ/[kg.K]"
                .spmp_fugacity = "Pa"
                .spmp_fugacityCoefficient = "-"
                .spmp_heatCapacityCp = "kJ/[kg.K]"
                .spmp_heatCapacityCv = "kJ/[kg.K]"
                .spmp_jouleThomsonCoefficient = "K/Pa"
                .spmp_logFugacityCoefficient = "-"
                .spmp_massflow = "kg/s"
                .spmp_massfraction = "-"
                .spmp_molarflow = "mol/s"
                .spmp_molarfraction = "-"
                .spmp_molecularWeight = "kg/kmol"
                .spmp_pressure = "Pa"
                .spmp_speedOfSound = "m/s"
                .spmp_temperature = "K"
                .spmp_thermalConductivity = "W/[m.K]"
                .spmp_viscosity = "Pa.s"
                .spmp_volumetricFlow = "m3/s"
                .spmp_cinematic_viscosity = "m2/s"
                .tdp_idealGasHeatCapacity = "kJ/[kg.K]"
                .tdp_surfaceTension = "N/m"
                .tdp_thermalConductivityOfLiquid = "W/[m.K]"
                .tdp_thermalConductivityOfVapor = "W/[m.K]"
                .tdp_vaporPressure = "Pa"
                .tdp_viscosityOfLiquid = "Pa.s"
                .tdp_viscosityOfVapor = "Pa.s"
                .tpmp_kvalue = "-"
                .tpmp_logKvalue = "-"
                .tpmp_surfaceTension = "N/m"
                .spmp_heatflow = "kW"
                .spmp_head = "m"
                .spmp_deltaP = "Pa"
                .spmp_deltaT = "K"

            End With

        End Sub

    End Class

    <Serializable()> Public Class UnitsSI_Deriv1

        Inherits Units

        Public Sub New()

            With Me

                .Name = App.GetLocalString("Custom1BR")

                .accel = "m2/s"
                .area = "m2"
                .diameter = "mm"
                .distance = "m"
                .force = "N"
                .heat_transf_coeff = "W/[m2.K]"
                .mass_conc = "kg/m3"
                .molar_conc = "kmol/m3"
                .molar_volume = "m3/kmol"
                .reac_rate = "kmol.[m3.s]"
                .spec_vol = "m3/kg"
                .time = "s"
                .volume = "m3"
                .thickness = "mm"
                .molar_enthalpy = "kJ/kmol"
                .molar_entropy = "kJ/[kmol.K]"
                .velocity = "m/s"
                .foulingfactor = "K.m2/W"

                .pdp_boilingPointTemperature = "C"
                .pdp_meltingTemperature = "C"
                .spmp_activity = "Pa"
                .spmp_activityCoefficient = "-"
                .spmp_compressibility = "Pa-1"
                .spmp_compressibilityFactor = "-"
                .spmp_density = "kg/m3"
                .spmp_enthalpy = "kJ/kg"
                .spmp_entropy = "kJ/[kg.K]"
                .spmp_excessEnthalpy = "kJ/kg"
                .spmp_excessEntropy = "kJ/[kg.K]"
                .spmp_fugacity = "Pa"
                .spmp_fugacityCoefficient = "-"
                .spmp_heatCapacityCp = "kJ/[kg.K]"
                .spmp_heatCapacityCv = "kJ/[kg.K]"
                .spmp_jouleThomsonCoefficient = "K/Pa"
                .spmp_logFugacityCoefficient = "-"
                .spmp_massflow = "kg/h"
                .spmp_massfraction = "-"
                .spmp_molarflow = "m3/d @ BR"
                .spmp_molarfraction = "-"
                .spmp_molecularWeight = "kg/kmol"
                .spmp_pressure = "kgf/cm2_a"
                .spmp_speedOfSound = "m/s"
                .spmp_temperature = "C"
                .spmp_thermalConductivity = "W/[m.K]"
                .spmp_viscosity = "cP"
                .spmp_volumetricFlow = "m3/h"
                .spmp_cinematic_viscosity = "cSt"
                .tdp_idealGasHeatCapacity = "kJ/[kg.K]"
                .tdp_surfaceTension = "N/m"
                .tdp_thermalConductivityOfLiquid = "W/[m.K]"
                .tdp_thermalConductivityOfVapor = "W/[m.K]"
                .tdp_vaporPressure = "kgf/cm2_a"
                .tdp_viscosityOfLiquid = "cP"
                .tdp_viscosityOfVapor = "cP"
                .tpmp_kvalue = "-"
                .tpmp_logKvalue = "-"
                .tpmp_surfaceTension = "N/m"
                .spmp_heatflow = "kW"
                .spmp_head = "m"
                .spmp_deltaP = "kgf/cm2"
                .spmp_deltaT = "C."

            End With

        End Sub

    End Class

    <Serializable()> Public Class UnitsSI_Deriv2

        Inherits Units

        Public Sub New()

            With Me

                .Name = App.GetLocalString("Custom2SC")

                .accel = "m2/s"
                .area = "m2"
                .diameter = "mm"
                .distance = "m"
                .force = "N"
                .heat_transf_coeff = "W/[m2.K]"
                .mass_conc = "kg/m3"
                .molar_conc = "kmol/m3"
                .molar_volume = "m3/kmol"
                .reac_rate = "kmol.[m3.s]"
                .spec_vol = "m3/kg"
                .time = "s"
                .volume = "m3"
                .thickness = "mm"
                .molar_enthalpy = "kJ/kmol"
                .molar_entropy = "kJ/[kmol.K]"
                .velocity = "m/s"
                .foulingfactor = "K.m2/W"

                .pdp_boilingPointTemperature = "C"
                .pdp_meltingTemperature = "C"
                .spmp_activity = "Pa"
                .spmp_activityCoefficient = "-"
                .spmp_compressibility = "Pa-1"
                .spmp_compressibilityFactor = "-"
                .spmp_density = "kg/m3"
                .spmp_enthalpy = "kJ/kg"
                .spmp_entropy = "kJ/[kg.K]"
                .spmp_excessEnthalpy = "kJ/kg"
                .spmp_excessEntropy = "kJ/[kg.K]"
                .spmp_fugacity = "Pa"
                .spmp_fugacityCoefficient = "-"
                .spmp_heatCapacityCp = "kJ/[kg.K]"
                .spmp_heatCapacityCv = "kJ/[kg.K]"
                .spmp_jouleThomsonCoefficient = "K/Pa"
                .spmp_logFugacityCoefficient = "-"
                .spmp_massflow = "kg/h"
                .spmp_massfraction = "-"
                .spmp_molarflow = "m3/d @ SC"
                .spmp_molarfraction = "-"
                .spmp_molecularWeight = "kg/kmol"
                .spmp_pressure = "kgf/cm2_a"
                .spmp_speedOfSound = "m/s"
                .spmp_temperature = "C"
                .spmp_thermalConductivity = "W/[m.K]"
                .spmp_viscosity = "cP"
                .spmp_volumetricFlow = "m3/h"
                .spmp_cinematic_viscosity = "cSt"
                .tdp_idealGasHeatCapacity = "kJ/[kg.K]"
                .tdp_surfaceTension = "N/m"
                .tdp_thermalConductivityOfLiquid = "W/[m.K]"
                .tdp_thermalConductivityOfVapor = "W/[m.K]"
                .tdp_vaporPressure = "kgf/cm2_a"
                .tdp_viscosityOfLiquid = "cP"
                .tdp_viscosityOfVapor = "cP"
                .tpmp_kvalue = "-"
                .tpmp_logKvalue = "-"
                .tpmp_surfaceTension = "N/m"
                .spmp_heatflow = "kW"
                .spmp_head = "m"
                .spmp_deltaP = "kgf/cm2"
                .spmp_deltaT = "C."

            End With

        End Sub

    End Class

    <Serializable()> Public Class UnitsSI_Deriv3

        Inherits Units

        Public Sub New()

            With Me

                .Name = App.GetLocalString("Custom3CNTP")

                .accel = "m/s2"
                .area = "m2"
                .diameter = "mm"
                .distance = "m"
                .force = "N"
                .heat_transf_coeff = "W/[m2.K]"
                .mass_conc = "kg/m3"
                .molar_conc = "kmol/m3"
                .molar_volume = "m3kmol"
                .reac_rate = "kmol.[m3.s]"
                .spec_vol = "m3/kg"
                .time = "s"
                .volume = "m3"
                .thickness = "mm"
                .molar_enthalpy = "kJ/kmol"
                .molar_entropy = "kJ/[kmol.K]"
                .velocity = "m/s"
                .foulingfactor = "K.m2/W"

                .pdp_boilingPointTemperature = "C"
                .pdp_meltingTemperature = "C"
                .spmp_activity = "Pa"
                .spmp_activityCoefficient = "-"
                .spmp_compressibility = "Pa-1"
                .spmp_compressibilityFactor = "-"
                .spmp_density = "kg/m3"
                .spmp_enthalpy = "kJ/kg"
                .spmp_entropy = "kJ/[kg.K]"
                .spmp_excessEnthalpy = "kJ/kg"
                .spmp_excessEntropy = "kJ/[kg.K]"
                .spmp_fugacity = "Pa"
                .spmp_fugacityCoefficient = "-"
                .spmp_heatCapacityCp = "kJ/[kg.K]"
                .spmp_heatCapacityCv = "kJ/[kg.K]"
                .spmp_jouleThomsonCoefficient = "K/Pa"
                .spmp_logFugacityCoefficient = "-"
                .spmp_massflow = "kg/h"
                .spmp_massfraction = "-"
                .spmp_molarflow = "m3/d @ CNTP"
                .spmp_molarfraction = "-"
                .spmp_molecularWeight = "kg/kmol"
                .spmp_pressure = "kgf/cm2_a"
                .spmp_speedOfSound = "m/s"
                .spmp_temperature = "C"
                .spmp_thermalConductivity = "W/[m.K]"
                .spmp_viscosity = "cP"
                .spmp_volumetricFlow = "m3/h"
                .spmp_cinematic_viscosity = "cSt"
                .tdp_idealGasHeatCapacity = "kJ/[kg.K]"
                .tdp_surfaceTension = "N/m"
                .tdp_thermalConductivityOfLiquid = "W/[m.K]"
                .tdp_thermalConductivityOfVapor = "W/[m.K]"
                .tdp_vaporPressure = "kgf/cm2_a"
                .tdp_viscosityOfLiquid = "cP"
                .tdp_viscosityOfVapor = "cP"
                .tpmp_kvalue = "-"
                .tpmp_logKvalue = "-"
                .tpmp_surfaceTension = "N/m"
                .spmp_heatflow = "kW"
                .spmp_head = "m"
                .spmp_deltaP = "kgf/cm2"
                .spmp_deltaT = "C."

            End With

        End Sub

    End Class

    <Serializable()> Public Class UnitsSI_Deriv4

        Inherits Units

        Public Sub New()

            With Me

                .Name = App.GetLocalString("Custom4")

                .accel = "m/s2"
                .area = "m2"
                .diameter = "mm"
                .distance = "m"
                .force = "N"
                .heat_transf_coeff = "W/[m2.K]"
                .mass_conc = "kg/m3"
                .molar_conc = "kmol/m3"
                .molar_volume = "m3kmol"
                .reac_rate = "kmol.[m3.s]"
                .spec_vol = "m3/kg"
                .time = "s"
                .volume = "m3"
                .thickness = "mm"
                .molar_enthalpy = "kJ/kmol"
                .molar_entropy = "kJ/[kmol.K]"
                .velocity = "m/s"
                .foulingfactor = "K.m2/W"

                .pdp_boilingPointTemperature = "C"
                .pdp_meltingTemperature = "C"
                .spmp_activity = "Pa"
                .spmp_activityCoefficient = "-"
                .spmp_compressibility = "Pa-1"
                .spmp_compressibilityFactor = "-"
                .spmp_density = "kg/m3"
                .spmp_enthalpy = "kJ/kg"
                .spmp_entropy = "kJ/[kg.K]"
                .spmp_excessEnthalpy = "kJ/kg"
                .spmp_excessEntropy = "kJ/[kg.K]"
                .spmp_fugacity = "Pa"
                .spmp_fugacityCoefficient = "-"
                .spmp_heatCapacityCp = "kJ/[kg.K]"
                .spmp_heatCapacityCv = "kJ/[kg.K]"
                .spmp_jouleThomsonCoefficient = "K/Pa"
                .spmp_logFugacityCoefficient = "-"
                .spmp_massflow = "kg/d"
                .spmp_massfraction = "-"
                .spmp_molarflow = "m3/d @ SC"
                .spmp_molarfraction = "-"
                .spmp_molecularWeight = "kg/kmol"
                .spmp_pressure = "kPa"
                .spmp_speedOfSound = "m/s"
                .spmp_temperature = "C"
                .spmp_thermalConductivity = "W/[m.K]"
                .spmp_viscosity = "cP"
                .spmp_volumetricFlow = "m3/d"
                .spmp_cinematic_viscosity = "cSt"
                .tdp_idealGasHeatCapacity = "kJ/[kg.K]"
                .tdp_surfaceTension = "N/m"
                .tdp_thermalConductivityOfLiquid = "W/[m.K]"
                .tdp_thermalConductivityOfVapor = "W/[m.K]"
                .tdp_vaporPressure = "kPa"
                .tdp_viscosityOfLiquid = "cP"
                .tdp_viscosityOfVapor = "cP"
                .tpmp_kvalue = "-"
                .tpmp_logKvalue = "-"
                .tpmp_surfaceTension = "N/m"
                .spmp_heatflow = "kJ/d"
                .spmp_head = "m"
                .spmp_deltaP = "kgf/cm2"
                .spmp_deltaT = "C."

            End With

        End Sub

    End Class

    <Serializable()> Public Class UnitsENGLISH

        Inherits Units

        Public Sub New()

            With Me

                .Name = App.GetLocalString("SystemEnglish")

                .accel = "ft/s2"
                .area = "ft2"
                .diameter = "in"
                .distance = "ft"
                .force = "lbf"
                .heat_transf_coeff = "BTU/[ft2.h.R]"
                .mass_conc = "lbm/ft3"
                .molar_conc = "lbmol/ft3"
                .molar_volume = "ft3/lbmol"
                .reac_rate = "lbmol.[ft3.h]"
                .spec_vol = "ft3/lbm"
                .time = "h"
                .volume = "ft3"
                .thickness = "in"
                .molar_enthalpy = "BTU/lbmol"
                .molar_entropy = "BTU/[lbmol.R]"
                .velocity = "ft/s"
                .foulingfactor = "ft2.h.F/BTU"

                .pdp_boilingPointTemperature = "R"
                .pdp_meltingTemperature = "R"
                .spmp_activity = "lbf/ft2"
                .spmp_activityCoefficient = "-"
                .spmp_compressibility = "ft2/lbf"
                .spmp_compressibilityFactor = "-"
                .spmp_density = "lbm/ft3"
                .spmp_enthalpy = "BTU/lbm"
                .spmp_entropy = "BTU/[lbm.R]"
                .spmp_excessEnthalpy = "BTU/lbm"
                .spmp_excessEntropy = "BTU/[lbm.R]"
                .spmp_fugacity = "lbf/ft2"
                .spmp_fugacityCoefficient = "-"
                .spmp_heatCapacityCp = "BTU/[lbm.R]"
                .spmp_heatCapacityCv = "BTU/[lbm.R]"
                .spmp_jouleThomsonCoefficient = "R/[lbf/ft2]"
                .spmp_logFugacityCoefficient = "-"
                .spmp_massflow = "lbm/h"
                .spmp_massfraction = "-"
                .spmp_molarflow = "lbmol/h"
                .spmp_molarfraction = "-"
                .spmp_molecularWeight = "lbm/lbmol"
                .spmp_pressure = "lbf/ft2"
                .spmp_speedOfSound = "ft/s"
                .spmp_temperature = "R"
                .spmp_thermalConductivity = "BTU/[ft.h.R]"
                .spmp_viscosity = "lbm/[ft.h]"
                .tdp_idealGasHeatCapacity = "BTU/[lbm.R]"
                .tdp_surfaceTension = "lbf/in"
                .tdp_thermalConductivityOfLiquid = "BTU/[ft.h.R]"
                .tdp_thermalConductivityOfVapor = "BTU/[ft.h.R]"
                .tdp_vaporPressure = "lbf/ft2"
                .tdp_viscosityOfLiquid = "lbm/[ft.h]"
                .tdp_viscosityOfVapor = "lbm/[ft.h]"
                .tpmp_kvalue = "-"
                .tpmp_logKvalue = "-"
                .tpmp_surfaceTension = "lbf/in"
                .spmp_volumetricFlow = "ft3/s"
                .spmp_cinematic_viscosity = "ft2/s"
                .spmp_heatflow = "BTU/h"
                .spmp_head = "ft"
                .spmp_deltaP = "lbf/ft2"
                .spmp_deltaT = "R."

            End With

        End Sub

    End Class

    <Serializable()> Public Class UnitsCGS

        Inherits Units

        Public Sub New()

            With Me

                .Name = App.GetLocalString("SystemCGS")

                .accel = "cm/s2"
                .area = "cm2"
                .diameter = "mm"
                .distance = "cm"
                .force = "dyn"
                .heat_transf_coeff = "cal/[cm2.s.C]"
                .mass_conc = "g/cm3"
                .molar_conc = "mol/cm3"
                .molar_volume = "cm3/mol"
                .reac_rate = "mol.[cm3.s]"
                .spec_vol = "cm3/g"
                .time = "s"
                .volume = "cm3"
                .thickness = "cm"
                .molar_enthalpy = "cal/mol"
                .molar_entropy = "cal/[mol.C]"
                .velocity = "cm/s"
                .foulingfactor = "C.cm2.s/cal"

                .pdp_boilingPointTemperature = "C"
                .pdp_meltingTemperature = "C"
                .spmp_activity = "atm"
                .spmp_activityCoefficient = "-"
                .spmp_compressibility = "atm-1"
                .spmp_compressibilityFactor = "-"
                .spmp_density = "g/cm3"
                .spmp_enthalpy = "cal/g"
                .spmp_entropy = "cal/[g.C]"
                .spmp_excessEnthalpy = "cal/g"
                .spmp_excessEntropy = "cal/[g.C]"
                .spmp_fugacity = "atm"
                .spmp_fugacityCoefficient = "-"
                .spmp_heatCapacityCp = "cal/[g.C]"
                .spmp_heatCapacityCv = "cal/[g.C]"
                .spmp_jouleThomsonCoefficient = "C/atm"
                .spmp_logFugacityCoefficient = "-"
                .spmp_massflow = "g/s"
                .spmp_massfraction = "-"
                .spmp_molarflow = "mol/s"
                .spmp_molarfraction = "-"
                .spmp_molecularWeight = "g/mol"
                .spmp_pressure = "atm"
                .spmp_speedOfSound = "cm/s"
                .spmp_temperature = "C"
                .spmp_thermalConductivity = "cal/[cm.s.C]"
                .spmp_viscosity = "cP"
                .tdp_idealGasHeatCapacity = "cal/[g.C]"
                .tdp_surfaceTension = "dyn/cm"
                .tdp_thermalConductivityOfLiquid = "cal/[cm.s.C]"
                .tdp_thermalConductivityOfVapor = "cal/[cm.s.C]"
                .tdp_vaporPressure = "atm"
                .tdp_viscosityOfLiquid = "cP"
                .tdp_viscosityOfVapor = "cP"
                .tpmp_kvalue = "-"
                .tpmp_logKvalue = "-"
                .tpmp_surfaceTension = "dyn/cm2"
                .spmp_volumetricFlow = "cm3/s"
                .spmp_cinematic_viscosity = "cSt"
                .spmp_heatflow = "kcal/h"
                .spmp_head = "m"
                .spmp_deltaT = "C."
                .spmp_deltaP = "atm"

            End With

        End Sub

    End Class

    <Serializable()> Public Class Converter

        Public Function ConvertToSI(ByVal unit As String, ByVal value As Double)

            Select Case unit

                Case "m/s"
                    Return value
                Case "cm/s"
                    Return value / 100
                Case "mm/s"
                    Return value / 1000
                Case "km/h"
                    Return value / 3.6
                Case "ft/h"
                    Return value / 11811
                Case "ft/min"
                    Return value / 196.85
                Case "ft/s"
                    Return value / 3.28084
                Case "in/s"
                    Return value / 39.3701

                Case "kPa"
                    Return value / 0.001
                Case "ftH2O"
                    Return value / 0.000334
                Case "inH2O"
                    Return value / 0.00401463
                Case "inHg"
                    Return value / 0.000295301
                Case "lbf/ft2"
                    Return value / 0.0208854
                Case "mbar"
                    Return value / 0.01
                Case "mH2O"
                    Return value / 0.000101972
                Case "mmH2O"
                    Return value / 0.101972
                Case "mmHg"
                    Return value / 0.00750064
                Case "MPa"
                    Return value / 0.000001
                Case "psi"
                    Return value / 0.000145038
                Case "bar"
                    Return value * 100000
                Case "kPag"
                    Return (value + 101.325) * 1000
                Case "barg"
                    Return (value + 1) * 100000
                Case "kgf/cm2g"
                    Return (value + 1.033) * 101325 / 1.033
                Case "psig"
                    Return (value + 14.696) * 6894.8

                Case "kg/d"
                    Return value / 60 / 60 / 24
                Case "kg/min"
                    Return value / 60
                Case "lb/h"
                    Return value / 7936.64
                Case "lb/min"
                    Return value / 132.277
                Case "lb/s"
                    Return 2.20462

                Case "mol/h"
                    Return value / 3600
                Case "mol/d"
                    Return value / 3600 / 24
                Case "kmol/s"
                    Return value * 1000
                Case "kmol/h"
                    Return value * 1000 / 3600
                Case "kmol/d"
                    Return value * 1000 / 3600 / 24

                Case "bbl/h"
                    Return value / 22643.3
                Case "bbl/d"
                    Return value / 543440
                Case "ft3/min"
                    Return value / 35.3147 / 60
                Case "ft3/s"
                    Return value / 35.3147
                Case "gal[UK]/h"
                    Return value / 791889
                Case "gal[UK]/s"
                    Return value / 219.969
                Case "gal[US]/h"
                    Return value / 951019
                Case "gal[US]/min"
                    Return value / 15850.3
                Case "L/h"
                    Return value / 3600000.0
                Case "L/min"
                    Return value / 60000
                Case "L/s"
                    Return value / 1000
                Case "m3/d"
                    Return value / 3600 / 24

                Case "BTU/h"
                    Return value / 3412.14
                Case "BTU/s"
                    Return value / 0.947817
                Case "cal/s"
                    Return value / 238.846
                Case "HP"
                    Return value / 1.35962
                Case "kcal/h"
                    Return value / 859.845
                Case "kJ/h"
                    Return value / 3600
                Case "kJ/d"
                    Return value / 3600 / 24
                Case "MW"
                    Return value / 0.001
                Case "W"
                    Return value / 1000

                Case "BTU/lb"
                    Return value / 0.429923
                Case "cal/g"
                    Return value / 0.238846
                Case "kcal/kg"
                    Return value / 0.238846

                Case "kJ/kmol"
                    Return value
                Case "cal/mol"
                    Return value * 0.0041868 * 1000
                Case "BTU/lbmol"
                    Return value * 1.05506 * 1000
                Case "kJ/[kmol.K]"
                    Return value * 1
                Case "cal/[mol.°C]"
                    Return value * 0.0041868 * 1000
                Case "cal/[mol.C]"
                    Return value * 0.0041868 * 1000
                Case "BTU/[lbmol.R]"
                    Return value * 1.05506 * 1000

                Case "K.m2/W"
                    Return value
                Case "C.cm2.s/cal"
                    Return value * 0.000023885
                Case "ft2.h.F/BTU"
                    Return value * 0.17611


                Case "m2"
                    Return value
                Case "cm2"
                    Return value / 10000.0
                Case "ft2"
                    Return value / 10.7639
                Case "h"
                    Return value / 3600
                Case "s"
                    Return value
                Case "min."
                    Return value / 60
                Case "ft3"
                    Return value / 35.3147
                Case "m3"
                    Return value
                Case "cm3"
                    Return value / 1000000.0
                Case "L"
                    Return value / 1000.0
                Case "cm3/mol"
                    Return value / 1000.0
                Case "m3/kmol"
                    Return value
                Case "m3/mol"
                    Return value * 1000
                Case "ft3/lbmol"
                    Return value / 35.3147 / 1000
                Case "mm"
                    Return value / 1000
                Case "in.", "in"
                    Return value / 39.3701
                Case "dyn"
                    Return value / 100000
                Case "N"
                    Return value
                Case "lbf"
                    Return value / 0.224809
                Case "mol/L"
                    Return value
                Case "kmol/m3"
                    Return value
                Case "mol/cm3"
                    Return value * 1000000.0 / 1000
                Case "mol/mL"
                    Return value * 1000000.0 / 1000
                Case "lbmol/ft3"
                    Return value * 35.3147 * 1000
                Case "g/L"
                    Return value
                Case "kg/m3"
                    Return value
                Case "g/cm3"
                    Return value * 1000000.0 / 1000
                Case "g/mL"
                    Return value * 1000000.0 / 1000
                Case "m2/s"
                    Return value
                Case "ft/s2"
                    Return value / 3.28084
                Case "cm2/s"
                    Return value / 10000.0
                Case "W/[m2.K]"
                    Return value
                Case "BTU/[ft2.h.R]"
                    Return value / 0.17611
                Case "cal/[cm.s.°C]"
                    Return value / 0.0000238846
                Case "cal/[cm.s.C]"
                    Return value / 0.0000238846
                Case "m3/kg"
                    Return value
                Case "ft3/lbm"
                    Return value * 0.062428
                Case "cm3/g"
                    Return value / 1000
                Case "kmol/[m3.s]"
                    Return value * 1000
                Case "kmol/[m3.min.]"
                    Return value / 60 * 1000
                Case "kmol/[m3.h]"
                    Return value / 3600 * 1000
                Case "mol/[m3.s]"
                    Return value
                Case "mol/[m3.min.]"
                    Return value / 60
                Case "mol/[m3.h]"
                    Return value / 3600
                Case "mol/[L.s]"
                    Return value / 1000
                Case "mol/[L.min.]"
                    Return value / 60 / 1000
                Case "mol/[L.h]"
                    Return value / 3600 / 1000
                Case "mol/[cm3.s]"
                    Return value * 1000000.0
                Case "mol/[cm3.min.]"
                    Return value / 60 * 1000000.0
                Case "mol/[cm3.h]"
                    Return value / 3600 * 1000000.0
                Case "lbmol.[ft3.h]"
                    Return value / 3600 * 35.3147 * 1000
                Case "°C"
                    Return value + 273.15
                Case "C"
                    Return value + 273.15
                Case "°C."
                    Return value
                Case "C."
                    Return value
                Case "atm"
                    Return value * 101325
                Case "g/s"
                    Return value / 1000
                Case "mol/s"
                    Return value
                Case "kmol/s"
                    Return value * 1000
                Case "cal/g"
                    Return value / 0.238846
                Case "g/cm3"
                    Return value * 1000
                Case "dyn/cm"
                    Return value * 0.001
                Case "dyn/cm2"
                    Return value * 0.001
                Case "cal/[cm.s.°C]"
                    Return value / 0.00238846
                Case "cal/[cm.s.C]"
                    Return value / 0.00238846
                Case "cm3/s"
                    Return value * 0.000001
                Case "cal/[g.°C]"
                    Return value / 0.238846
                Case "cal/[g.C]"
                    Return value / 0.238846
                Case "cSt"
                    Return value / 1000000.0
                Case "mm2/s"
                    Return value / 1000000.0
                Case "Pa.s"
                    Return value
                Case "cP"
                    Return value / 1000
                Case "kcal/h"
                    Return value / 859.845
                Case "m"
                    Return value
                Case "R"
                    Return value / 1.8
                Case "R."
                    Return value / 1.8
                Case "lbf/ft2"
                    Return value / 0.0208854
                Case "lbm/h"
                    Return value / 7936.64
                Case "lbmol/h"
                    Return value * 453.59237 / 3600
                Case "BTU/lbm"
                    Return value / 0.429923
                Case "lbm/ft3"
                    Return value / 0.062428
                Case "lbf/in"
                    Return value / 0.00571015
                Case "BTU/[ft.h.R]"
                    Return value / 0.577789
                Case "ft3/s"
                    Return value / 35.3147
                Case "BTU/[lbm.R]"
                    Return value / 0.238846
                Case "ft2/s"
                    Return value / 10.7639
                Case "lbm/[ft.s]"
                    Return value / 0.671969
                Case "BTU/h"
                    Return value / 3412.14
                Case "ft"
                    Return value / 3.28084
                    'Customs
                Case "kgf/cm2_a"
                    Return value * 101325 / 1.033
                Case "kgf/cm2"
                    Return value * 101325 / 1.033
                Case "kgf/cm2_g"
                    Return (value + 1.033) * 101325 / 1.033
                Case "kg/h"
                    Return value / 3600
                Case "kg/d"
                    Return value / 3600 / 24
                Case "m3/h"
                    Return value / 3600
                Case "m3/d"
                    Return value / 3600 / 24
                Case "m3/d @ BR"
                    Return value / (24.055 * 3600 * 24 / 1000)
                Case "m3/d @ CNTP"
                    Return value / (22.71 * 3600 * 24 / 1000)
                Case "m3/d @ NC"
                    Return value / (22.71 * 3600 * 24 / 1000)
                Case "m3/d @ SC"
                    Return value / (23.69 * 3600 * 24 / 1000)
                Case "°F"
                    Return (value - 32) * 5 / 9 + 273.15
                Case "°F."
                    Return value / 1.8
                Case "F"
                    Return (value - 32) * 5 / 9 + 273.15
                Case "F."
                    Return value / 1.8
                Case "cm"
                    Return value / 100
                Case "cal/[mol.°C]"
                    Return value * 238.846 / 1000
                Case "cal/[mol.C]"
                    Return value * 238.846 / 1000
                Case "BTU/[lbmol.R]"
                    Return value * 1.8 * 0.947817
                Case Else
                    Return value
            End Select

        End Function

        Public Function ConvertFromSI(ByVal unit As String, ByVal value As Double)

            Select Case unit

                Case "m/s"
                    Return value
                Case "cm/s"
                    Return value * 100
                Case "mm/s"
                    Return value * 1000
                Case "km/h"
                    Return value * 3.6
                Case "ft/h"
                    Return value * 11811
                Case "ft/min"
                    Return value * 196.85
                Case "ft/s"
                    Return value * 3.28084
                Case "in/s"
                    Return value * 39.3701

                Case "kPa"
                    Return value * 0.001
                Case "ftH2O"
                    Return value * 0.000334
                Case "inH2O"
                    Return value * 0.00401463
                Case "inHg"
                    Return value * 0.000295301
                Case "lbf/ft2"
                    Return value * 0.0208854
                Case "mbar"
                    Return value * 0.01
                Case "mH2O"
                    Return value * 0.000101972
                Case "mmH2O"
                    Return value * 0.101972
                Case "mmHg"
                    Return value * 0.00750064
                Case "MPa"
                    Return value * 0.000001
                Case "psi"
                    Return value * 0.000145038
                Case "bar"
                    Return value / 100000
                Case "kPag"
                    Return (value - 101325) / 1000
                Case "barg"
                    Return (value - 101325) / 100000
                Case "kgf/cm2g"
                    Return (value - 101325) / 101325 * 1.033
                Case "psig"
                    Return (value - 101325) * 0.000145038

                Case "kg/d"
                    Return value * 60 * 60 * 24
                Case "kg/min"
                    Return value * 60
                Case "lb/h"
                    Return value * 7936.64
                Case "lb/min"
                    Return value * 132.277
                Case "lb/s"
                    Return 2.20462

                Case "bbl/h"
                    Return value * 22643.3
                Case "bbl/d"
                    Return value * 543440
                Case "ft3/min"
                    Return value * 35.3147 * 60
                Case "ft3/s"
                    Return value * 35.3147
                Case "gal[UK]/h"
                    Return value * 791889
                Case "gal[UK]/s"
                    Return value * 219.969
                Case "gal[US]/h"
                    Return value * 951019
                Case "gal[US]/min"
                    Return value * 15850.3
                Case "L/h"
                    Return value * 3600000.0
                Case "L/min"
                    Return value * 60000
                Case "L/s"
                    Return value * 1000
                Case "m3/d"
                    Return value * 3600 * 24

                Case "mol/h"
                    Return value * 3600
                Case "mol/d"
                    Return value * 3600 * 24
                Case "kmol/s"
                    Return value / 1000
                Case "kmol/h"
                    Return value / 1000 * 3600
                Case "kmol/d"
                    Return value / 1000 * 3600 * 24

                Case "BTU/h"
                    Return value * 3412.14
                Case "BTU/s"
                    Return value * 0.947817
                Case "cal/s"
                    Return value * 238.846
                Case "HP"
                    Return value * 1.35962
                Case "kcal/h"
                    Return value * 859.845
                Case "kJ/h"
                    Return value * 3600
                Case "kJ/d"
                    Return value * 3600 * 24
                Case "MW"
                    Return value * 0.001
                Case "W"
                    Return value * 1000

                Case "BTU/lb"
                    Return value * 0.429923
                Case "cal/g"
                    Return value * 0.238846
                Case "kcal/kg"
                    Return value * 0.238846

                Case "kJ/kmol"
                    Return value
                Case "cal/mol"
                    Return value / 0.0041868 / 1000
                Case "BTU/lbmol"
                    Return value / 1.05506 / 1000
                Case "kJ/[kmol.K]"
                    Return value
                Case "cal/[mol.°C]"
                    Return value / 0.0041868 / 1000
                Case "cal/[mol.C]"
                    Return value / 0.0041868 / 1000
                Case "BTU/[lbmol.R]"
                    Return value / 1.05506 / 1000

                Case "K.m2/W"
                    Return value
                Case "C.cm2.s/cal"
                    Return value / 0.000023885
                Case "ft2.h.F/BTU"
                    Return value / 0.17611

                Case "m2"
                    Return value
                Case "cm2"
                    Return value * 10000.0
                Case "ft2"
                    Return value * 10.7639
                Case "h"
                    Return value * 3600
                Case "s"
                    Return value
                Case "min."
                    Return value * 60
                Case "ft3"
                    Return value * 35.3147
                Case "m3"
                    Return value
                Case "cm3"
                    Return value * 1000000.0
                Case "L"
                    Return value * 1000.0
                Case "cm3/mol"
                    Return value * 1000.0
                Case "m3/kmol"
                    Return value
                Case "m3/mol"
                    Return value / 1000
                Case "ft3/lbmol"
                    Return value * 35.3147 * 1000
                Case "mm"
                    Return value * 1000
                Case "in.", "in"
                    Return value * 39.3701
                Case "dyn"
                    Return value * 100000
                Case "N"
                    Return value
                Case "lbf"
                    Return value * 0.224809
                Case "mol/L"
                    Return value
                Case "kmol/m3"
                    Return value
                Case "mol/cm3"
                    Return value / 1000000.0 * 1000
                Case "mol/mL"
                    Return value / 1000000.0 * 1000
                Case "lbmol/ft3"
                    Return value * 35.3147 * 1000
                Case "g/L"
                    Return value
                Case "kg/m3"
                    Return value
                Case "g/cm3"
                    Return value / 1000000.0 * 1000
                Case "g/mL"
                    Return value / 1000000.0 * 1000
                Case "lbm/ft3"
                    Return value * 0.062428
                Case "m2/s"
                    Return value
                Case "ft/s2"
                    Return value * 3.28084
                Case "cm2/s"
                    Return value * 10000.0
                Case "W/[m2.K]"
                    Return value
                Case "BTU/[ft2.h.R]"
                    Return value * 0.17611
                Case "cal/[cm.s.°C]"
                    Return value * 0.0000238846
                Case "cal/[cm.s.C]"
                    Return value * 0.0000238846
                Case "m3/kg"
                    Return value
                Case "ft3/lbm"
                    Return value / 0.062428
                Case "cm3/g"
                    Return value * 1000
                Case "kmol/[m3.s]"
                    Return value / 1000
                Case "kmol/[m3.min.]"
                    Return value * 60 / 1000
                Case "kmol/[m3.h]"
                    Return value * 3600 / 1000
                Case "mol/[m3.s]"
                    Return value
                Case "mol/[m3.min.]"
                    Return value * 60
                Case "mol/[m3.h]"
                    Return value * 3600
                Case "mol/[L.s]"
                    Return value * 1000
                Case "mol/[L.min.]"
                    Return value * 60 * 1000
                Case "mol/[L.h]"
                    Return value * 3600 * 1000
                Case "mol/[cm3.s]"
                    Return value / 1000000.0
                Case "mol/[cm3.min.]"
                    Return value * 60 / 1000000.0
                Case "mol/[cm3.h]"
                    Return value * 3600 / 1000000.0
                Case "lbmol.[ft3.h]"
                    Return value * 3600 / 35.3147 / 1000
                Case "°C"
                    Return value - 273.15
                Case "C"
                    Return value - 273.15
                Case "°C."
                    Return value
                Case "C."
                    Return value
                Case "atm"
                    Return value / 101325
                Case "g/s"
                    Return value * 1000
                Case "mol/s"
                    Return value
                Case "cal/g"
                    Return value * 0.238846
                Case "g/cm3"
                    Return value / 1000
                Case "dyn/cm"
                    Return value / 0.001
                Case "dyn/cm2"
                    Return value / 0.001
                Case "cal/[cm.s.°C]"
                    Return value * 0.00238846
                Case "cal/[cm.s.C]"
                    Return value * 0.00238846
                Case "cm3/s"
                    Return value / 0.000001
                Case "cal/[g.°C]"
                    Return value * 0.238846
                Case "cal/[g.C]"
                    Return value * 0.238846
                Case "cSt"
                    Return value * 1000000.0
                Case "mm2/s"
                    Return value * 1000000.0
                Case "Pa.s"
                    Return value
                Case "cP"
                    Return value * 1000
                Case "kcal/h"
                    Return value * 859.845
                Case "R"
                    Return value * 1.8
                Case "R."
                    Return value * 1.8
                Case "lbf/ft2"
                    Return value * 0.0208854
                Case "lbm/h"
                    Return value * 7936.64
                Case "lbmol/h"
                    Return value / 453.59237 * 3600
                Case "BTU/lbm"
                    Return value * 0.429923
                Case "lbf/in"
                    Return value * 0.00571015
                Case "BTU/[ft.h.R]"
                    Return value * 0.577789
                Case "ft3/s"
                    Return value * 35.3147
                Case "BTU/[lbm.R]"
                    Return value * 0.238846
                Case "ft2/s"
                    Return value * 10.7639
                Case "lbm/[ft.s]"
                    Return value * 0.671969
                Case "BTU/h"
                    Return value * 3412.14
                Case "ft"
                    Return value * 3.28084
                    'Customs
                Case "kgf/cm2_a"
                    Return value / 101325 * 1.033
                Case "kgf/cm2"
                    Return value / 101325 * 1.033
                Case "kgf/cm2_g"
                    Return value * 1.033 / 101325 - 1.033
                Case "kg/h"
                    Return value * 3600
                Case "kg/d"
                    Return value * 3600 * 24
                Case "m3/h"
                    Return value * 3600
                Case "m3/d"
                    Return value * 3600 * 24
                Case "m3/d @ BR"
                    Return value * (24.055 * 3600 * 24 / 1000)
                Case "m3/d @ CNTP"
                    Return value * (22.71 * 3600 * 24 / 1000)
                Case "m3/d @ NC"
                    Return value * (22.71 * 3600 * 24 / 1000)
                Case "m3/d @ SC"
                    Return value * (23.69 * 3600 * 24 / 1000)
                Case "°F"
                    Return (value - 273.15) * 9 / 5 + 32
                Case "°F."
                    Return value * 1.8
                Case "F"
                    Return (value - 273.15) * 9 / 5 + 32
                Case "F."
                    Return value * 1.8
                Case "cm"
                    Return value * 100
                Case "cal/[mol.°C]"
                    Return value / 238.846 * 1000
                Case "cal/[mol.C]"
                    Return value / 238.846 * 1000
                Case "BTU/[lbmol.R]"
                    Return value / 1.8 / 0.947817
                Case Else
                    Return value
            End Select

        End Function

        Public Sub New()

        End Sub

    End Class

End Namespace
