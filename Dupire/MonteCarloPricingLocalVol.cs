using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dupire
{
    public class MonteCarloPricingLocalVol 
    {
        private double r; 
        private int N; // Number of samples   
        private int M; // Number of time steps   
        private double[] ssviParams;

        public MonteCarloPricingLocalVol(double riskFreeRate, int numSamples, int numTimeSteps, double[] ssviParams)
        {
            r = riskFreeRate;
            N = numSamples;
            M = numTimeSteps;
            this.ssviParams = ssviParams;
        }

        public double CalculateEurCallOptionPrice(double S0, double K, double T)
        {
            Func<double, double> callPayoff = (x) => Math.Max(x - K, 0);
            return CalculateEurOptionPrice(S0, T, callPayoff);
        }

        public double CalculateEurPutOptionPrice(double S0, double K, double T)
        {
            Func<double, double> putPayoff = (x) => Math.Max(K - x, 0);
            return CalculateEurOptionPrice(S0, T, putPayoff);
        }

        private double CalculateEurOptionPrice(double S0, double T, Func<double, double> payoffFn)
        {
            double[,] stockPaths = MonteCarloPaths.GenerateMcPaths(N, M, T, r, S0, ssviParams);
            double Price = 0;
            for (int n = 0; n < N; ++n)
            {
                Price += payoffFn(stockPaths[n, M - 1]);
            }

            return Math.Exp(-r * T)*(Price / N);
        }

        public double CalculateAsianCallOptionPrice(double S, double K, double[] T, double maturity )
        {
            Func<double, double> callPayoff = (x) => Math.Max(x - K, 0);
            return CalculateAsianOptionPrice(S, T, maturity, callPayoff);
        }

        public double CalculateAsianPutOptionPrice(double S, double K, double[] T, double maturity)
        {
            Func<double, double> putPayoff = (x) => Math.Max(K - x, 0);
            return CalculateAsianOptionPrice(S, T, maturity, putPayoff);
        }

        private double CalculateAsianOptionPrice(double S0, double[] T, double maturity, Func<double, double> payoffFn)
        {
            int nMonitorTimes = T.Length;
            double[,] stockPaths = MonteCarloPaths.GenerateMcPaths(N, M, maturity, r, S0, ssviParams); // size N x nMonitorTimes


            // the is the average for Asian option price
            double averageAlongPath;

            // this is the MC average
            double averagePayoff = 0;
            for(int n=0; n< N; ++n)
            {
                averageAlongPath = 0;
                for (int i = 0; i < nMonitorTimes; ++i)
                {
                    int index = (int)Math.Round((T[i] / maturity) * (M - 1));
                    averageAlongPath += stockPaths[n, index]; 
                }
                averageAlongPath *= 1.0 / nMonitorTimes;
                averagePayoff += payoffFn(averageAlongPath);
            }
            averagePayoff *= 1.0 / N;
            return Math.Exp(-r * maturity) * averagePayoff;
        }

        public double CalculateLookbackOptionPrice(double S0, double maturity)
        {
            double[,] stockPaths = MonteCarloPaths.GenerateMcPaths(N, M, maturity, r, S0, ssviParams);

            double payoff = 0;

            for(int n = 0; n<N; ++n)
            {
                double minValue = double.MaxValue;
                for(int m =0; m<M; ++m)
                {
                    if (stockPaths[n, m] < minValue)
                        minValue = stockPaths[n, m];
                }
                payoff += stockPaths[n, M - 1] - minValue;                
            }            
            return Math.Exp(-r*maturity)*payoff/N;
        }
     

        public double CalculateCallBarrierOptionPrice(double S0, double K, double maturity, string DownOrUp, string InOrOut, double barrier)
        {
            Func<double,double> callPayoff = (x) => Math.Max(x - K, 0);
            return CalculateBarrierOptionPrice(S0, maturity, DownOrUp, InOrOut, barrier,callPayoff);
        }

        public double CalculatePutBarrierOptionPrice(double S0, double K, double maturity, string DownOrUp, string InOrOut, double barrier)
        {
            Func<double, double> putPayoff = (x) => Math.Max(K -x, 0);
            return CalculateBarrierOptionPrice(S0, maturity, DownOrUp, InOrOut, barrier, putPayoff);
        }

        private double CalculateBarrierOptionPrice(double S0, double maturity, string DownOrUp, string InOrOut, double barrier, Func<double,double> payoffFn)
        {
            double[,] stockPaths = MonteCarloPaths.GenerateMcPaths(N, M, maturity, r, S0, ssviParams);
            double payoff = 0;

            for (int n = 0; n < N; ++n)
            {
                bool barrierAchieved = false;
                for (int m = 0; m < M; ++m)
                {
                    if (DownOrUp == "D")
                    {
                        if (stockPaths[n, m] <= barrier)
                            barrierAchieved = true;
                    }
                    else if (DownOrUp == "U")
                    {
                        if (stockPaths[n, m] >= barrier)
                            barrierAchieved = true;
                    }
                    else
                        throw new ArgumentException("Barrier type must be either D or U");
                }
                if (barrierAchieved)
                {
                    if (InOrOut == "I")
                        payoff += payoffFn(stockPaths[n, M - 1]);
                    else if (InOrOut != "O")
                        throw new ArgumentException("InOrOut must be I or O.");
                }
                else
                {
                    if(InOrOut == "O")
                        payoff += payoffFn(stockPaths[n, M - 1]);
                    else if (InOrOut != "I")
                        throw new ArgumentException("InOrOut must be I or O.");
                }

            }
            return Math.Exp(-r * maturity) * payoff / N;

        }
    }
}
