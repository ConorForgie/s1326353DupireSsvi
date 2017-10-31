using System;
using System.Collections.Generic;
using System.Linq;

using MathNet.Numerics;
using System.Diagnostics;

namespace Dupire
{
    class DupireCmdLine
    {
        public static void Main (string[] args)
		{

            //ConstrainedMinimizationTest.RunTest();

            //SsviTests.RunVolatilityTest();

            // MonteCarloTests.RunEurOptionPricing();
            double[] ssviParams = new double[] { Math.Sqrt(0.1), 1, 0.7, 0.2, 0.3 };
            double r = 0.025;
            double S0 = 100;
            double T = 1;
            int N = 100000;
            int M = 128;
            Stopwatch timer = Stopwatch.StartNew();
            double[,] paths = MonteCarloPaths.GenerateMcPaths(N, M, T, r, S0, ssviParams);
            timer.Stop();
            TimeSpan serial_length = timer.Elapsed;
            Console.WriteLine("Serial elapsed time = " + serial_length.ToString("mm\\:ss\\.ff"));

            Stopwatch timer1 = Stopwatch.StartNew();
            double[][] paths1 = MonteCarloPaths.ParallelGenerateMcPaths(N, M, T, r, S0, ssviParams);
            timer1.Stop();
            TimeSpan serial_length1 = timer1.Elapsed;
            Console.WriteLine("Parallel elapsed time = " + serial_length1.ToString("mm\\:ss\\.ff"));


            //using (System.IO.StreamWriter file = new System.IO.StreamWriter("Paths.csv"))
            //{
            //    for (int i = 0; i < paths.GetLength(0); ++i)
            //    {
            //        for (int j = 0; j < paths.GetLength(1); ++j)
            //            file.Write(paths[i, j].ToString() + ",");
            //        file.Write(Environment.NewLine);
            //    }
            //}
            Console.WriteLine ("All done.");
            Console.ReadKey();
		}
	}

    class MonteCarloTests
    {
        public static void RunPathBuildTest()
        {
            double[] ssviParams = new double[] { Math.Sqrt(0.1), 1, 0.7, 0.2, 0.3 };           
            double r = 0.025;
            double S0 = 100;
            double T = 1;
            int N = 6;
            int M = 4;

            double[,] paths = MonteCarloPaths.GenerateMcPaths(N, M, T, r, S0, ssviParams);
            int rLength = paths.GetLength(0);
            int cLength = paths.GetLength(1);

            for (int i = 0; i < rLength; ++i)
            {
                for(int j=0; j < cLength; ++j)
                {
                    Console.Write("{0}, ",paths[i,j]);
                }
                Console.WriteLine();
            }
        }

        public static void RunEurOptionPricing()
        {
            double[] ssviParams = new double[] { Math.Sqrt(0.1), 1, 0.7, 0.2, 0.3 };
            double r = 0.025;
            double S0 = 100;
            double T = 1;
            int N = 10000;
            int M = 128;
            double K = 102;
            double k = Math.Log(K * Math.Exp(-r * T) / S0);

            MonteCarloPricingLocalVol pricer = new MonteCarloPricingLocalVol(r, N, M, ssviParams);
            double McPriceCall = pricer.CalculateEurPutOptionPrice(S0, K, T);
            Ssvi SsviSurface = new Ssvi(ssviParams[0], ssviParams[1], ssviParams[2], ssviParams[3], ssviParams[4]);
            double sigmaBs = Math.Sqrt(SsviSurface.OmegaSsvi(T, k) / T);
            double BsPriceCall = BlackScholesFormula.CalculatePutOptionPrice(sigmaBs, 100, k, r, T);
            Console.WriteLine("Call price: MC: £{0}, BS £{1}", McPriceCall, BsPriceCall);
                
        }
    }

    class SsviTests
    {
        public static void RunVolatilityTest()
        {
            double alpha = Math.Sqrt(0.1);
            double beta = 1;
            double gamma = 0.7;
            double eta = 0.2;
            double rho = 0.3;
            double r = 0.025;
            double S = 100;
            double K = 102.5313;
            double T = 1;

            double k = Math.Log(K * Math.Exp(-r * T) / S);
            Ssvi SsviTest = new Ssvi(alpha, beta, gamma, eta, rho);
            Console.WriteLine("SSVI Volatility = {0}", SsviTest.VolatilityDup(T, k));
        }
    }

    class ConstrainedMinimizationTest
    {
        public static void function1_fvec(double[] x, double[] fi, object obj)
        {
            //
            // this callback calculates
            // f0(x0,x1) = 100*(x0+3)^4,
            // f1(x0,x1) = (x1-3)^4
            //
            fi[0] = 10 * System.Math.Pow(x[0] + 3, 4);
            fi[1] = System.Math.Pow(x[1] - 3, 4);
        }
        public static int RunTest()
        {
            //
            // This example demonstrates minimization of F(x0,x1) = f0^2+f1^2, where 
            //
            //     f0(x0,x1) = 10*(x0+3)^2
            //     f1(x0,x1) = (x1-3)^2
            //
            // with boundary constraints
            //
            //     -1 <= x0 <= +1
            //     -1 <= x1 <= +1
            //
            // using "V" mode of the Levenberg-Marquardt optimizer.
            //
            // Optimization algorithm uses:
            // * function vector f[] = {f1,f2}
            //
            // No other information (Jacobian, gradient, etc.) is needed.
            //
            double[] x = new double[] { 0, 0 };
            double[] bndl = new double[] { -1, -1 };
            double[] bndu = new double[] { +1, +1 };
            double epsg = 0.0000000001;
            double epsf = 0;
            double epsx = 0;
            int maxits = 0;
            alglib.minlmstate state;
            alglib.minlmreport rep;

            alglib.minlmcreatev(2, x, 0.0001, out state);
            alglib.minlmsetbc(state, bndl, bndu);
            alglib.minlmsetcond(state, epsg, epsf, epsx, maxits);
            alglib.minlmoptimize(state, function1_fvec, null, null);
            alglib.minlmresults(state, out x, out rep);

            System.Console.WriteLine("{0}", rep.terminationtype); // EXPECTED: 4
            System.Console.WriteLine("{0}", alglib.ap.format(x, 2)); // EXPECTED: [-1,+1]
            
            return 0;
        }

    }
}
