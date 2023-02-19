using NUnit.Framework;
using DTL.DTL.SimulationObjects.PropertyPackages;
using DTL.DTL.SimulationObjects.Streams;
using System;

namespace TestProject1
{
    public class Tests
    {
        DTL.Thermodynamics.Calculator? dtlc;
        [SetUp]
        public void Setup()
        {
            dtlc = new DTL.Thermodynamics.Calculator();
            dtlc.Initialize();
            dtlc.DisableParallelProcessing();
            dtlc.DisableSIMDExtensions();
        }

        [Test]
        public void Test1()
        {


            string[] comps = new string[] { "Water" };//{ "Ethane", "Methane", "Propane" };
            double[] fracs = new double[] { 1.0 };//{ 0.5, 0.5, 1e-6 };

            //double[] critpt = dtlc.CalcTrueCriticalPoint("PR", comps, fracs);
            PropertyPackage prpp = dtlc.GetPropPackInstance("Peng-Robinson (PR)");//"IAPWS-IF97 Steam Tables"

            dtlc.SetupPropertyPackage(prpp, comps, fracs);

            //set stream conditions for mixture phase ("0")

            prpp.CurrentMaterialStream.Phases["0"].SPMProperties.pressure = 1.0 * 1e5;
            //prpp.CurrentMaterialStream.Phases["2"].SPMProperties.molarfraction = 1;
            prpp.CurrentMaterialStream.Phases["0"].SPMProperties.temperature = 400;

            //prpp.CurrentMaterialStream.Phases["0"].SPMProperties.enthalpy = h;
            prpp.CurrentMaterialStream.SpecType = MaterialStream.Flashspec.Temperature_and_Pressure;//MaterialStream.Flashspec.Pressure_and_VaporFraction;
            //prpp.CurrentMaterialStream.SpecType = MaterialStream.Flashspec.Pressure_and_Enthalpy;

            //calculate the stream - equilibrium and properties
            prpp.CurrentMaterialStream.Calculate(true, true);

            Double t = prpp.CurrentMaterialStream.Phases["0"].SPMProperties.temperature.GetValueOrDefault();
            Console.WriteLine(t);
            Assert.Pass();
        }

        [Test]
        public void Test2()
        {


            string[] comps = new string[] { "Water" };//{ "Ethane", "Methane", "Propane" };
            double[] fracs = new double[] { 1.0 };//{ 0.5, 0.5, 1e-6 };

            //double[] critpt = dtlc.CalcTrueCriticalPoint("PR", comps, fracs);
            PropertyPackage prpp = dtlc.GetPropPackInstance("IAPWS-IF97 Steam Tables");//"IAPWS-IF97 Steam Tables"

            dtlc.SetupPropertyPackage(prpp, comps, fracs);

            //set stream conditions for mixture phase ("0")

            prpp.CurrentMaterialStream.Phases["0"].SPMProperties.pressure = 1.0 * 1e5;
            //prpp.CurrentMaterialStream.Phases["2"].SPMProperties.molarfraction = 1;
            prpp.CurrentMaterialStream.Phases["0"].SPMProperties.temperature = 400;

            //prpp.CurrentMaterialStream.Phases["0"].SPMProperties.enthalpy = h;
            prpp.CurrentMaterialStream.SpecType = MaterialStream.Flashspec.Temperature_and_Pressure;//MaterialStream.Flashspec.Pressure_and_VaporFraction;
            //prpp.CurrentMaterialStream.SpecType = MaterialStream.Flashspec.Pressure_and_Enthalpy;

            //calculate the stream - equilibrium and properties
            prpp.CurrentMaterialStream.Calculate(true, true);

            Double t = prpp.CurrentMaterialStream.Phases["0"].SPMProperties.temperature.GetValueOrDefault();
            Console.WriteLine(t);
            Assert.Pass();
        }
    }
}