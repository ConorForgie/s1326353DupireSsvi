using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dupire
{
    public class CalibrationFailedException : Exception
    {
        public CalibrationFailedException()
        {
        }
        public CalibrationFailedException(string message)
            : base(message)
        {
        }
    }

    public enum CalibrationOutcome { NotStarted, FinishedOK, FailedMaxItReached, FailedOtherReason };

    public struct OptionMarketData
    {
        public double maturity;
        public double strike;
        public string type;
        public double marketPrice;
    }

    public class SsviCalibrator
    {
        private const double defaultAccuracy = 1e-4;
        private const int defaultMaxIterations = 500;
        private double accuracy;
        private int maxIterations;

        private LinkedList<OptionMarketData> marketOptionsList;
        private double r0; // initial interest rate, this is observed, no need to calibrate to options
        private double S0; // underlying initial stock price

        private CalibrationOutcome outcome;

        private double[] calibratedParams;


        public SsviCalibrator()
        {
            accuracy = defaultAccuracy;
            maxIterations = defaultMaxIterations;
            marketOptionsList = new LinkedList<OptionMarketData>();
            r0 = 0;
            calibratedParams = new double[] { 0.2, 1, 0.1, 0.1, 0.1 };
        }

        public SsviCalibrator(double r0, double S0, double accuracy, int maxIterations)
        {
            this.r0 = r0;
            this.S0 = S0;
            this.accuracy = accuracy;
            this.maxIterations = maxIterations;
            marketOptionsList = new LinkedList<OptionMarketData>();
            calibratedParams = new double[] { 0.2, 1, 0.1, 0.1, 0.1 };
        }

        public void SetGuessParameters(double alpha, double beta, double gamma, double eta, double rho)
        {
            calibratedParams = new double[] { alpha, beta, gamma, eta, rho };
        }

        public void AddObservedOption(double maturity, double strike, double mktPrice, string type)
        {
            OptionMarketData observedOption;
            observedOption.maturity = maturity;
            observedOption.type = type;
            observedOption.strike = strike;
            observedOption.marketPrice = mktPrice;
            marketOptionsList.AddLast(observedOption);
        }

        // Calculate difference between implied volatility and model volatility
        public double CalcMeanSquareErrorBetweenThetaAndMarketATM(Ssvi m)
        {
            double meanSqErr = 0;
            int nElements = 0;
            foreach (OptionMarketData option in marketOptionsList)
            {
                double maturity = option.maturity;
                string type = option.type;
                double theta = m.Theta(maturity);
                double bsImpliedVolatility;
                if (option.type == "Call" || option.type == "call" || option.type == "C" || option.type == "c")
                {
                    bsImpliedVolatility = BlackScholesImpliedVolatility.CalculateImpliedVolCall(option.marketPrice, S0, 0, r0, maturity);
                }
                else if (option.type == "Put" || option.type == "put" || option.type == "P" || option.type == "p")
                {
                    bsImpliedVolatility = BlackScholesImpliedVolatility.CalculateImpliedVolPut(option.marketPrice, S0, 0, r0, maturity);
                }
                else
                    throw new CalibrationFailedException("Option type in marketOptionsList was not Call, call, C, c, Put, put, P or p");

                double difference = theta - bsImpliedVolatility * bsImpliedVolatility * maturity;
                meanSqErr += difference * difference;
                nElements++;
            }
            if (nElements > 0)
                return meanSqErr / nElements;
            else
                throw new ArgumentNullException("There is no market data to use for error calculation.");
        }

        public double CalcMeanSquareErrorBetweenOmegaAndMarketNonATM(Ssvi m)
        {
            double meanSqErr = 0;
            int nElements = 0;
            foreach (OptionMarketData option in marketOptionsList)
            {
                double maturity = option.maturity;
                string type = option.type;
                double strike = option.strike;
                double k = Math.Log(strike * Math.Exp(-r0) / S0);
                double omega = m.OmegaSsvi(maturity, k);
                double bsImpliedVolatility;
                if (option.type == "Call" || option.type == "call" || option.type == "C" || option.type == "c")
                {
                    bsImpliedVolatility = BlackScholesImpliedVolatility.CalculateImpliedVolCall(option.marketPrice, S0, k, r0, maturity);
                }
                else if (option.type == "Put" || option.type == "put" || option.type == "P" || option.type == "p")
                {
                    bsImpliedVolatility = BlackScholesImpliedVolatility.CalculateImpliedVolPut(option.marketPrice, S0, k, r0, maturity);
                }
                else
                    throw new CalibrationFailedException("Option type in marketOptionsList was not Call, call, C, c, Put, put, P or p");
                double difference = omega - bsImpliedVolatility * bsImpliedVolatility * maturity;
                meanSqErr += difference * difference;
                nElements++;
            }
            if (nElements > 0)
                return meanSqErr / nElements;
            else
                throw new ArgumentNullException("There is no market data to use for error calculation.");
        }

        // Used by Alglib minimisation algorithm
        public void CalibrationObjectiveFunctionATM(double[] ssviParams, ref double func, object obj)
        {
            Ssvi m = new Ssvi(ssviParams);
            func = CalcMeanSquareErrorBetweenThetaAndMarketATM(m);
        }


        public void CalibrateAlphaAndBetaATM()
        {
            outcome = CalibrationOutcome.NotStarted;

            double[] initialParams = new double[Ssvi.numModelParams];
            calibratedParams.CopyTo(initialParams, 0);
            double epsg = accuracy;
            double epsf = accuracy; 
            double epsx = accuracy;
            double diffstep = 1.0e-6;
            int maxits = maxIterations;
            double stpmax = 0.05;

            alglib.minlbfgsstate state;
            alglib.minlbfgsreport rep;
            alglib.minlbfgscreatef(1,initialParams, diffstep, out state);
            alglib.minlbfgssetcond(state, epsg, epsf, epsx, maxits);
            alglib.minlbfgssetstpmax(state, stpmax);

            alglib.minlbfgsoptimize(state, CalibrationObjectiveFunctionATM, null, null);
            double[] resultParams = new double[Ssvi.numModelParams];
            alglib.minlbfgsresults(state, out resultParams, out rep);

            if (rep.terminationtype == 1			// relative function improvement is no more than EpsF.
                || rep.terminationtype == 2			// relative step is no more than EpsX.
                || rep.terminationtype == 4)        // gradient norm is no more than EpsG
            {
                outcome = CalibrationOutcome.FinishedOK;
                calibratedParams = resultParams;
            }
            else if (rep.terminationtype == 5)
            {	// MaxIts steps was taken
                outcome = CalibrationOutcome.FailedMaxItReached;
                calibratedParams = resultParams;
            }
            else
            {
                outcome = CalibrationOutcome.FailedOtherReason;
                throw new CalibrationFailedException("Alpha and Beta calibration failed badly.");
            }
        }


        // Used by Alglib minimisation algorithm
        public void CalibrationObjectiveFunctionNonATM(double[] x, ref double func, object obj)
        {
            Ssvi m = new Ssvi(x);
            func = CalcMeanSquareErrorBetweenOmegaAndMarketNonATM(m);
        }

        public void CalibrateEtaGammaRhoNonATM()
        {
            outcome = CalibrationOutcome.NotStarted;

            double[] initialParams = new double[Ssvi.numModelParams];
            calibratedParams.CopyTo(initialParams, 0);
            double epsg = accuracy;
            double epsf = accuracy;
            double epsx = accuracy;
            double diffstep = 1.0e-6;
            int maxits = maxIterations;
            double stpmax = 0.05;


            alglib.minlbfgsstate state;
            alglib.minlbfgsreport rep;
            alglib.minlbfgscreatef(1, initialParams, diffstep, out state);
            alglib.minlbfgssetcond(state, epsg, epsf, epsx, maxits);
            alglib.minlbfgssetstpmax(state, stpmax);


            alglib.minlbfgsoptimize(state, CalibrationObjectiveFunctionNonATM, null, null);
            double[] resultParams = new double[Ssvi.numModelParams];
            alglib.minlbfgsresults(state, out resultParams, out rep);

            if (rep.terminationtype == 1            // relative function improvement is no more than EpsF.
                || rep.terminationtype == 2         // relative step is no more than EpsX.
                || rep.terminationtype == 4)        // gradient norm is no more than EpsG
            {
                outcome = CalibrationOutcome.FinishedOK;
                calibratedParams = resultParams;
            }
            else if (rep.terminationtype == 5)
            {   // MaxIts steps was taken
                outcome = CalibrationOutcome.FailedMaxItReached;
                calibratedParams = resultParams;
            }
            else
            {
                outcome = CalibrationOutcome.FailedOtherReason;
                throw new CalibrationFailedException("Ssvi model calibration failed badly.");
            }
        }

        public void GetCalibrationStatusAlphaBeta(ref CalibrationOutcome calibOutcome, ref double pricingError)
        {
            calibOutcome = outcome;
            Ssvi m = new Ssvi(calibratedParams);
            pricingError = CalcMeanSquareErrorBetweenThetaAndMarketATM(m);
        }

        public void GetCalibrationStatus(ref CalibrationOutcome calibOutcome, ref double pricingError)
        {
            calibOutcome = outcome;
            Ssvi m = new Ssvi(calibratedParams);
            pricingError = CalcMeanSquareErrorBetweenOmegaAndMarketNonATM(m);
        }

        // Must ensure calibration of all Parameters alpha, beta, eta, gamma & rho
        public Ssvi GetCalibratedModel()
        {
            Ssvi m = new Ssvi(calibratedParams);
            return m;
        }

    }    

}