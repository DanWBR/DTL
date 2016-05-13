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
            
            Stopwatch sw = new Stopwatch();

            // calls a flash calculation 100x through the default interface

            sw.Start();
            for (int i = 0; i < 100; i++)
            {
                object[,] result = dtlc.PHFlash(prpp, 0, P + i, h, comps, fracs, null, null, null, null, 300d);
            }
            sw.Stop();

            Console.WriteLine(sw.ElapsedMilliseconds.ToString());
            Console.ReadKey();

            // calls a flash calculation 100x through the object-oriented structured function

            //creates a material stream and associates it with the property package.
            prpp.SetMaterial(dtlc.CreateMaterialStream(comps, fracs));

            sw.Restart();
            for (int i = 0; i < 100; i++)
            {
                var result = prpp.FlashBase.CalculateEquilibrium(FlashSpec.P, FlashSpec.H, P + i, h, prpp, fracs, null, 300d);
            }
            sw.Stop();

            Console.WriteLine(sw.ElapsedMilliseconds.ToString());
            Console.ReadKey();

            // calls a flash calculation 100x directly through the flash algorithm instance

            sw.Restart();
            for (int i = 0; i < 100; i++)
            {
                var result = prpp.FlashBase.Flash_PH(fracs, P + i, h, 300d, prpp, true, null);
            }
            sw.Stop();

            Console.WriteLine(sw.ElapsedMilliseconds.ToString());
            Console.ReadKey();

            // calsl a flash calculation 100x using Parallel.For

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
