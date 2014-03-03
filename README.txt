DWSIM Standalone Thermodynamics Library
Version 2.2.2
Copyright 2014 Daniel Medeiros

The DWSIM Standalone Thermodynamics Library is a .NET/Mono managed dynamic link library (DLL) that exposes DWSIM’s thermodynamics engine to external applications using a simple programming interface, with no dependency on external components. 

DWSIM Standalone Thermodynamics Library is free for commercial and non-commercial use. Read the license.txt file for more details.

Usage

To use the library, add a reference to DWSIM.Thermodynamics.dll to your project. All calculation functions will be available in the DTL.Thermodynamics namespace, Calculator class. 

Some of the available functions in this library include:
- Single Compound Property Calculator
- Single Phase Mixture Property Calculator
- PT, PH, PS, PVF and TVF Flash Calculators, using an algorithm of your choice (two or three phases)
- Other auxiliary functions

Methods

Property and Equilibrium calculation functions require parameters that must be one or more values returned by GetPropPackList, GetCompoundList, GetPropList, GetCompoundConstPropList, GetCompoundTDepPropList, GetCompoundPDepPropList and GetPhaseList. They are self-explanatory, and will return values in an array of strings.

For instance, the PTFlash function requires the name of the Property Package to use, the compound names and mole fractions, temperature in K, pressure in Pa and you may optionally provide new interaction parameters that will override the ones used internally by the library. The calculation results will be returned as a (n+2) x (3) string matrix, where n is the number of compounds. First row will contain the phase names, the second will contain the phase mole fractions and the other lines will contain the compound mole fractions in the corresponding phases.

For PH, PS, TVF and PVF flash calculation functions, an additional line is returned that will contain the temperature in K or pressure in Pa in the last matrix column.