using System;

namespace Duprie
{
    public class NewtonSolver
    {
        const double h = 1e-6;
        private double maxError;
        private int maxIter;

        public NewtonSolver(double maxError, int maxIter) { this.maxError = maxError; this.maxIter = maxIter; }

        //Returns an approximate solution s to f(x)=0 using Newtons method starting at x0 and such that |f(x)| < maxError
        //where f is a function R->R and fPrime is f' (if null then approximated using central difference).
        //Throws exception if derivative is too small or number of iterations exceeds maxIter.
        public double Solve(Func<double, double> f, Func<double, double> fPrime, double x0)
        {
            if (fPrime == null)
            {
                fPrime = (x) => (f(x + h) - f(x)) / h;
            }

            int i = 0;
            double xn = x0;

            while (i < maxIter)
            {
                if (Math.Abs(f(xn)) < maxError)
                    return xn;

                if (Math.Abs(fPrime(xn)) < 1e-16)
                    throw new SystemException("NewtonsMethod failed - derivative too small");

                xn = xn - f(xn) / fPrime(xn);

                ++i;
            }

            throw new SystemException("NewtonsMethod did not converge");
        }
    }
}

