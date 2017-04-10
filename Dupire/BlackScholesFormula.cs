using System;

namespace Dupire
{
    // Using log-moneyness
    public static class BlackScholesFormula
    {
        public static double CalculateCallOptionPrice(double sigma, double S, double k, double r, double T)
        {
            System.Diagnostics.Debug.Assert(sigma > 0 && T > 0 && S > 0, "Need sigma > 0, T > 0 and S > 0.");
            double d1 = (-r * T - k + (r + (sigma * sigma) / 2) * (T)) / (sigma * Math.Sqrt(T));
            double d2 = d1 - sigma * Math.Sqrt(T);
            return S * Gaussian.N(d1) - S * Math.Exp(k) * Gaussian.N(d2);
        }

        public static double GetCallFromPutPrice(double S, double k, double r, double T, double putPrice)
        {
            return putPrice - Math.Exp(k) * S + S;
        }

        public static double GetPutFromCallPrice(double S, double k, double r, double T, double callPrice)
        {
            return callPrice - S + Math.Exp(k) * S;
        }

        public static double CalculatePutOptionPrice(double sigma, double S, double k, double r, double T)
        {
            return GetPutFromCallPrice(S, k, r, T, CalculateCallOptionPrice(sigma, S, k, r, T));
        }


    }

}


