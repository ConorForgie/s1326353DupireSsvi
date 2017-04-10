using System;

namespace Dupire
{

    class Gaussian
    {
        //Gaussian function for approximation for cumulative normal distribution
        static public double psi(double x)
        {
            double A = 1.0 / Math.Sqrt(2.0 * Math.PI);
            return A * Math.Exp(-x * x * 0.5);
        }

        // Approximation for cumulative normal distribution used in B-S Formula
        static public double N(double x)
        {
            System.Diagnostics.Debug.Assert(!Double.IsNaN(x) && !Double.IsInfinity(x),
                "Gaussian.N: x has to be a valid number.");
            double a1 = 0.4361836;
            double a2 = -0.1201676;
            double a3 = 0.9372980;
            double k = 1.0 / (1.0 + (0.33267 * x));
            if (x >= 0.0)
            {
                return 1.0 - psi(x) * (a1 * k + (a2 * k * k) + (a3 * k * k * k));
            }
            else
            {
                return 1.0 - N(-x);
            }
        }
    }
}