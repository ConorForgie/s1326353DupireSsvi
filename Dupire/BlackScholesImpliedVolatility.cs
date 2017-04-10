using Duprie;
using System;


namespace Dupire
{
    // In terms of Log-Moneyness
    class BlackScholesImpliedVolatility
    {
        public static double CalculateImpliedVolCall(double C, double S, double k, double r, double T, double x0 = 0.5, double maxErr = 1e-6, int N = 10000)
        {
            Func<double, double> F = (x) =>
            {
                return C - BlackScholesFormula.CalculateCallOptionPrice(x, S, k, r, T);
            };
            NewtonSolver s = new NewtonSolver(maxErr, N);
            return s.Solve(F, null, x0);

        }

        public static double CalculateImpliedVolPut(double P, double S, double k, double r, double T, double x0 = 0.5, double maxErr = 1e-6, int N = 10000)
        {

            double callPrice = BlackScholesFormula.GetCallFromPutPrice(S, k, r, T, P);
            if (callPrice < 0)
                throw new System.ArgumentException("Input arguments violate put/call parity.");

            return CalculateImpliedVolCall(callPrice, S, k, r, T, x0, maxErr, N);
        }
    }
}
