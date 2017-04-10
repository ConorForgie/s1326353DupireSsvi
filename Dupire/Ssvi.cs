using System;

namespace Dupire
{
    public class Ssvi
    {
        public const int numModelParams = 5;

        const double h = 1e-6;
        //SSVI Paramters
        protected double alpha, beta, rho, gamma, eta;

        public Ssvi(double alpha, double beta, double gamma, double eta, double rho)
        {
            this.alpha = alpha;
            this.beta = beta;            
            this.gamma = gamma;
            this.eta = eta;
            this.rho = rho;
        }
        public Ssvi(double[] parameters)
        {
            alpha = parameters[0];
            beta = parameters[1];
            gamma = parameters[2];
            eta = parameters[3];
            rho = parameters[4];
        }

        public double VolatilityDup(double T, double k)
        {
            if (T == 0) // At t = 0 there can be no volatilty. (This is to stop Phi = Infty)
                return 0;
            double dwdk = (OmegaSsvi(T, k + h) - OmegaSsvi(T, k)) / h; // Use finite differences
            if (Math.Abs(dwdk) < 1e-6)         // check dw/dk ~ 0    
                return Math.Sqrt((OmegaSsvi(T + h, k) - OmegaSsvi(T, k)) / h);
            else
            {
                double d2wdk2 = (OmegaSsvi(T, k + 2 * h) - 2 * OmegaSsvi(T, k + h) + OmegaSsvi(T, k)) / (h * h);
                double omega = OmegaSsvi(T, k);
                double g = (Math.Pow(1 - k * dwdk / (2 * omega), 2) - (dwdk * dwdk / 4) * (1 / 4 + 1 / omega) + d2wdk2 / 2);
                return Math.Sqrt((OmegaSsvi(T + h, k) - OmegaSsvi(T, k)) / (h * g));
            }                
        }

        public double OmegaSsvi(double t, double k)
        {
            double theta = Theta(t);
            double phi = Phi(theta, eta, gamma);
            return theta / 2 * (1 + rho * phi * k + Math.Sqrt(Math.Pow(phi * k + rho, 2) + (1 - rho) * (1 - rho) ));
        }

        private double Phi(double theta, double eta, double gamma)
        {
            if (eta <= 0 || gamma <= 0 || gamma >= 1)
               throw new ArgumentOutOfRangeException("Must have eta > 0 and 0 < gamma < 1");

            return eta / (Math.Pow(theta, gamma) * Math.Pow(1 + theta, 1 - gamma));
        }

        public double Theta(double t)
        {           
            return alpha * alpha * (Math.Exp(beta * beta * t) - 1);
        }

        public double GetAlpha() { return alpha; }
        public double GetBeta() { return beta; }
        public double GetEta() { return eta; }
        public double GetGamma() { return gamma; }
        public double GetRho() { return rho; }



    }
}
