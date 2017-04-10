using System;
using MathNet.Numerics.Distributions;

namespace Dupire
{
    public class MonteCarloPaths
    {        

        /// <summary>
        /// Generates a 2d array of N stock paths over M time steps.     
        /// </summary>
        /// <returns></returns>
        public static double[,] GenerateMcPaths(int N, int M, double T, double r, double S0, double[] ssviParams)
        {
            Ssvi ssvi = new Ssvi(ssviParams);

            double[,] pathOfS = new double[N, M];
            double[] Z = new double[N];
            double tau = T / M;

            for (int m = 1; m < M; ++m) // foreach time step
            {
                Normal.Samples(Z, 0, 1);    //Fill Z with normal 0,1 samples
                double sqrtTau = Math.Sqrt(tau);

                for (int n = 0; n < N; ++n) // foreach sample
                {
                    pathOfS[n, 0] = S0;
                    double s_n_m = pathOfS[n, m-1];
                    if (s_n_m < 0)
                        s_n_m = 1e-10;
                    double dw = sqrtTau * Z[n];
                    double k = Math.Log(s_n_m * Math.Exp(-r * T) / S0);
                    pathOfS[n, m] = s_n_m + r * s_n_m * tau + ssvi.VolatilityDup(tau * m, k) * s_n_m * dw;                    
                }
            }
            return pathOfS;
        }        
    }
}
