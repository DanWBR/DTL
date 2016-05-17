using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using DTL.DTL.SimulationObjects.PropertyPackages;
using DTL.DTL.SimulationObjects.Streams;
using DTL;

namespace DTLTest3
{
    class Program
    {
        static void Main(string[] args)
        {

            Test1();

        }

        static void Test1()
        {

            DTL.Thermodynamics.Calculator dtlc = new DTL.Thermodynamics.Calculator();
            dtlc.Initialize();
            string[] comps = new string[] { "Water", "Methane" };
            double[] fracs = new double[] { 0.01, 0.99 };
            double P = 101325;
            double h = 0;
            Parallel.For(0, 5000, (int i) =>
            {
                
                PropertyPackage prpp = dtlc.GetPropPackInstance("Peng-Robinson (PR)");
                
                dtlc.SetupPropertyPackage(prpp, comps, fracs);

                //set stream conditions for mixture phase ("0")

                prpp.CurrentMaterialStream.Phases["0"].SPMProperties.pressure = P + i;
                prpp.CurrentMaterialStream.Phases["0"].SPMProperties.enthalpy = h;
                prpp.CurrentMaterialStream.SpecType = MaterialStream.Flashspec.Pressure_and_Enthalpy;

                //calculate the stream - equilibrium and properties
                prpp.CurrentMaterialStream.Calculate(true, true);

                Double t = prpp.CurrentMaterialStream.Phases["0"].SPMProperties.temperature.GetValueOrDefault();
                
                //get vapor phase ("2") density
                double rho = prpp.CurrentMaterialStream.Phases["2"].SPMProperties.density.GetValueOrDefault();

                //phase IDs: 0 - mixture, 1 - overall liquid, 2 = vapor, 3 = liquid 1, 4 = liquid 2, 5 = liquid 3, 6 = aqueous, 7 = solid

                Console.WriteLine(t.ToString() + " " + rho.ToString());

            });

        
        }


        static void Test2()
        {

            DTL.Thermodynamics.Calculator dtlc = new DTL.Thermodynamics.Calculator();
            dtlc.Initialize();

            // enables parallel threads for low level calculations
            dtlc.EnableParallelProcessing();

            //enables SIMD acceleration for vector operations
            dtlc.EnableSIMDExtensions();

            PropertyPackage prpp = dtlc.GetPropPackInstance("Peng-Robinson (PR)");

            double P = 101325;
            double h = 0;

            string[] comps = new string[] { "Water", "Methane" };
            double[] fracs = new double[] { 0.01, 0.99 };

            // flash algorithm type
            // 0 or 2 = NL VLE, 1 = IO VLE, 3 = IO VLLE, 4 = Gibbs VLE, 5 = Gibbs VLLE, 6 = NL VLLE, 7 = NL SLE, 8 = NL Immisc., 9 = Simple LLE               

            int flashalg = 0;

            Stopwatch sw = new Stopwatch();

            // calls a flash calculation 100x through the default interface

            sw.Start();
            for (int i = 0; i < 100; i++)
            {
                object[,] result = dtlc.PHFlash(prpp, flashalg, P + i, h, comps, fracs, null, null, null, null, 300d);
            }
            sw.Stop();

            Console.WriteLine(sw.ElapsedMilliseconds.ToString());
            Console.ReadKey();

            // calls a flash calculation 100x through the object-oriented structured function

            //creates a material stream and associates it with the property package.
            prpp.SetMaterial(dtlc.CreateMaterialStream(comps, fracs));

            prpp.FlashAlgorithm = flashalg;

            sw.Restart();
            for (int i = 0; i < 100; i++)
            {
                var result = prpp.FlashBase.CalculateEquilibrium(FlashSpec.P, FlashSpec.H, P + i, h, prpp, fracs, null, 300d);
            }
            sw.Stop();

            Console.WriteLine(sw.ElapsedMilliseconds.ToString());
            Console.ReadKey();

            // calls a flash calculation 100x directly through the flash algorithm instance

            prpp.FlashAlgorithm = flashalg;

            sw.Restart();
            for (int i = 0; i < 100; i++)
            {
                var result = prpp.FlashBase.Flash_PH(fracs, P + i, h, 300d, prpp, false, null);
            }
            sw.Stop();

            Console.WriteLine(sw.ElapsedMilliseconds.ToString());
            Console.ReadKey();

            // calls a flash calculation 100x using Parallel.For

            prpp.FlashAlgorithm = flashalg;

            // this will make sure that a new flash algorithm instance is created to avoid thread locking
            prpp.ForceNewFlashAlgorithmInstance = true;
          
            sw.Restart();
            Parallel.For(0, 100, (int i) =>
            {
                var result = prpp.FlashBase.Flash_PH(fracs, P + i, h, 300d, prpp);
            });
            sw.Stop();

            Console.WriteLine(sw.ElapsedMilliseconds.ToString());
            Console.ReadKey();

        }
    }
}
