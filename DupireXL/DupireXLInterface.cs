using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using Dupire;


namespace DupireXL
{

    /// <summary>
    ///  The class that should implement the Excel interface. Please remember
    ///  to not change any of the function signatures. If you want to modify
    ///  the signature of a function: don't! Create a new function with a new 
    ///  name - with your own new function do as you like. If in doubt what is 
    ///  meant by function signature, then please refer to 
    ///  https://en.wikipedia.org/wiki/Type_signature 
    /// 
    ///  Remember to use try / catch here and report the errors!
    /// </summary>
    public partial class DupireXLInterface
    {
        private const string FunctionError = "Function Failed";

        [ExcelFunction(IsThreadSafe = true, Description = "About DupireXL function")]
        public static string AboutDupireXL() { return "Dupire local vol Excel Interface for OOPA 2016/17."; }

        [ExcelFunction(IsThreadSafe = true, Category = "Dupire", Description = "Returns the surname of the author of the code.")]
        public static string StudentSurname() { return "Forgie"; }

        [ExcelFunction(IsThreadSafe = true, Category = "Dupire", Description = "Returns the first name of the author of the code.")]
        public static string StudentFirstName() { return "Conor"; }

        [ExcelFunction(IsThreadSafe = true, Category = "Dupire", Description = "Returns the student ID of the author of the code.")]
        public static string StudentID() { return "s1326353"; }

        [ExcelFunction(IsThreadSafe = true,
            Category = "Dupire",
            Description = "Gets implied volatility from SSVI parametrized volatility surface.")]
        public static object DupireSsviVolatility([ExcelArgument(Description = "the current risky asset price")] double underlyingPrice,
                                                [ExcelArgument(Description = "constant continuously compounded rate of return")] double riskFreeRate,
                                                [ExcelArgument(Description = "parameter in theta(t) := alpha^2(e^(beta^2 t)  - 1)")]double alpha,
                                                [ExcelArgument(Description = "parameter in theta(t) := alpha^2(e^(beta^2 t)  - 1)")]double beta,
                                                [ExcelArgument(Description = "SSVI parameter")]double gamma,
                                                [ExcelArgument(Description = "SSVI parameter")]double eta,
                                                [ExcelArgument(Description = "SSVI parameter")]double rho,
                                                [ExcelArgument(Description = "option maturity (time to expiry)")]double maturity,
                                                [ExcelArgument(Description = "option strike")]double strike)
        {
            // This is useful for methods that take a while to run.
            if (ExcelDnaUtil.IsInFunctionWizard()) return null;

            try
            {
                double k = Math.Log(strike * Math.Exp(-riskFreeRate * maturity) / underlyingPrice);
                Ssvi SsviSurface = new Ssvi(alpha, beta, gamma, eta, rho);
                return Math.Sqrt(SsviSurface.OmegaSsvi(maturity, k) / maturity);
            }
            catch (Exception e)
            {
                XLInterfaceBase.AddErrorMessage("DupireSsviVolatility error: " + e.Message);
            }

            return FunctionError;
        }



        [ExcelFunction(IsThreadSafe = true,
            Category = "Dupire",
            Description = "Calculates the local volatility function.")]
        public static object DupireSsviCalculateSigmaDup([ExcelArgument(Description = "the current risky asset price")] double S0,
                                                [ExcelArgument(Description = "constant continuously compounded rate of return")] double riskFreeRate,
                                                [ExcelArgument(Description = "parameter in theta(t) := alpha^2(e^(beta^2 t)  - 1)")]double alpha,
                                                [ExcelArgument(Description = "parameter in theta(t) := alpha^2(e^(beta^2 t)  - 1)")]double beta,
                                                [ExcelArgument(Description = "SSVI parameter")]double gamma,
                                                [ExcelArgument(Description = "SSVI parameter")]double eta,
                                                [ExcelArgument(Description = "SSVI parameter")]double rho,
                                                [ExcelArgument(Description = "The value of S for which we want the Dupire Local Vol")]double S,
                                                [ExcelArgument(Description = "The value of t for which we want the Dupire Local Vol")]double t)

        {
            try
            {
                Ssvi SsviSurface = new Ssvi(alpha, beta, gamma, eta, rho);
                double k = Math.Log(S * Math.Exp(-riskFreeRate * t) / S0);
                return SsviSurface.VolatilityDup(t, k);
            }
            catch (Exception e)
            {
                XLInterfaceBase.AddErrorMessage("DupireSsviCalculateSigmaDup error: " + e.Message);
            }

            return FunctionError;
        }

        [ExcelFunction(IsThreadSafe = true,
            Category = "Dupire",
            Description = "Obtains European call/put price by finding the implied vol in SSVI surface.")]
        public static object DupireSsviEuropeanVanillaOptionPrice([ExcelArgument(Description = "the current risky asset price")] double S0,
                                                [ExcelArgument(Description = "constant continuously compounded rate of return")] double riskFreeRate,
                                                [ExcelArgument(Description = "parameter in theta(t) := alpha^2(e^(beta^2 t)  - 1)")]double alpha,
                                                [ExcelArgument(Description = "parameter in theta(t) := alpha^2(e^(beta^2 t)  - 1)")]double beta,
                                                [ExcelArgument(Description = "SSVI parameter")]double gamma,
                                                [ExcelArgument(Description = "SSVI parameter")]double eta,
                                                [ExcelArgument(Description = "SSVI parameter")]double rho,
                                                [ExcelArgument(Description = "option maturity (time to expiry)")]double maturity,
                                                [ExcelArgument(Description = "option strike")]double strike,
                                                [ExcelArgument(Description = "option type, 'C' for call 'P' for put")]string type)
        {
            try
            {
                Ssvi SsviSurface = new Ssvi(alpha, beta, gamma, eta, rho);
                double k = Math.Log(strike * Math.Exp(-riskFreeRate * maturity) / S0);
                double sigmaBs = Math.Sqrt(SsviSurface.OmegaSsvi(maturity, k) / maturity);

                if (type == "Call" || type == "call" || type == "C" || type == "c")
                {
                    return BlackScholesFormula.CalculateCallOptionPrice(sigmaBs, S0, k, riskFreeRate, maturity);
                }
                else if (type == "Put" || type == "put" || type == "P" || type == "p")
                {
                    return BlackScholesFormula.CalculatePutOptionPrice(sigmaBs, S0, k, riskFreeRate, maturity);
                }
                else
                    throw new ArgumentException("Input must be either Call, call, C, c, Put, put, P or p.");
            }
            catch (Exception e)
            {
                XLInterfaceBase.AddErrorMessage("DupireSsviEuropeanVanillaOptionPrice error: " + e.Message);
            }

            return FunctionError;
        }



        [ExcelFunction(IsThreadSafe = true,
            Category = "Dupire",
            Description = "Performs calibration of ATM part of SSVI surface from ATM European call / put prices.")]
        public static object[,] CalibrateDupireSsviParametersATM([ExcelArgument(Description = "A 2x2 XL Range with guess values for alpha and beta")] object guessModelParameters,
                                                    [ExcelArgument(Description = "constant continuously compounded rate of return")] double riskFreeRate,
                                                    [ExcelArgument(Description = "the current risky asset price")] double underlyingPrice,
                                                    [ExcelArgument(Description = "An Nx1 XL Range of option strikes")] object strikes,
                                                    [ExcelArgument(Description = "An Nx1 XL Range of option maturities")] object maturities,
                                                    [ExcelArgument(Description = "An Nx1 XL Range of option either C or P for Calls and Puts")] object type,
                                                    [ExcelArgument(Description = "An Nx1 XL Range of option prices")]object observedPrices,
                                                    [ExcelArgument(Description = "Accuracy parameter for calibration")]double accuracy,
                                                    [ExcelArgument(Description = "Number of iterations parameter for calibration")]int maxIterations)
        {
            try
            {
                if (ExcelDnaUtil.IsInFunctionWizard()) return null;

                double[] strikesArray = XLInterfaceBase.ConvertToVector<double>(strikes);
                double[] maturitesArray = XLInterfaceBase.ConvertToVector<double>(maturities);
                string[] typeArray = XLInterfaceBase.ConvertToVector<string>(type);
                double[] observedPricesArray = XLInterfaceBase.ConvertToVector<double>(observedPrices);

                if (strikesArray.Length != maturitesArray.Length
                    || maturitesArray.Length != typeArray.Length
                    || typeArray.Length != observedPricesArray.Length)
                {
                    // must improve error message display
                    XLInterfaceBase.AddErrorMessage("CalibrateSsviParameters: strikes, maturites, types and observedPrices must be of same length.");
                    return null;
                }

                KeyValuePair<string, double>[] paramPairs = XLInterfaceBase.ConvertToKeyValuePairs(guessModelParameters);
                if (paramPairs.Length != 2)
                {
                    XLInterfaceBase.AddErrorMessage("CalibrateSsviParameters: guessModelParameters must be two key - value pairs.");
                    return null;
                }

                double alpha = 0; double beta = 0;
                for (int paramIdx = 0; paramIdx < paramPairs.Length; ++paramIdx)
                {
                    KeyValuePair<string, double> pair = paramPairs[paramIdx];
                    string key = pair.Key;
                    double value = pair.Value;
                    if (key.Equals("alpha", StringComparison.CurrentCultureIgnoreCase))
                        alpha = value;
                    else if (key.Equals("beta", StringComparison.CurrentCultureIgnoreCase))
                        beta = value;
                    else
                    {
                        XLInterfaceBase.AddErrorMessage("CalibrateSsviParameters: guessModelParameters: unknown key: " + key);
                        return null;
                    }
                }

                SsviCalibrator ssviCalibrator = new SsviCalibrator(riskFreeRate, underlyingPrice, accuracy, maxIterations);
                ssviCalibrator.SetGuessParameters(alpha, beta, 0.1, 0.1, 0.1);

                int numObservedOptions = strikesArray.Length;
                for (int optionIdx = 0; optionIdx < numObservedOptions; ++optionIdx)
                {
                    ssviCalibrator.AddObservedOption(maturitesArray[optionIdx], strikesArray[optionIdx], observedPricesArray[optionIdx], typeArray[optionIdx]);
                }
                ssviCalibrator.CalibrateAlphaAndBetaATM();
                CalibrationOutcome outcome = CalibrationOutcome.NotStarted;
                double calibrationError = 0;
                ssviCalibrator.GetCalibrationStatusAlphaBeta(ref outcome, ref calibrationError);

                Ssvi calibratedModel = ssviCalibrator.GetCalibratedModel();

                // for output
                const int numCols = 2;
                const int numRows = 4;
                object[,] output = new object[numRows, numCols];
                output[0, 0] = "Alpha"; output[0, 1] = calibratedModel.GetAlpha();
                output[1, 0] = "Beta"; output[1, 1] = calibratedModel.GetBeta();
                output[2, 0] = "Minimizer Status";
                if (outcome == CalibrationOutcome.FinishedOK)
                    output[2, 1] = "OK";
                else if (outcome == CalibrationOutcome.FailedMaxItReached)
                    output[2, 1] = "Reached max. num. iterations.";
                else if (outcome == CalibrationOutcome.FailedOtherReason)
                    output[2, 1] = "Failed.";
                else
                    output[2, 1] = "Unknown outcome.";

                output[3, 0] = "Pricing error"; output[3, 1] = Math.Sqrt(calibrationError);
                return output;
            }
            catch (Exception e)
            {
                XLInterfaceBase.AddErrorMessage("CalibrateParameters: unknown error: " + e.Message);
            }

            return null;
        }

        [ExcelFunction(IsThreadSafe = true,
            Category = "Dupire",
            Description = "Performs calibration of wings of SSVI surface from European call / put prices.")]
        public static object[,] CalibrateDupireSsviParametersNonATM(
            [ExcelArgument(Description = "A 3x2 XL Range with guess values for gamma, eta, rho")] object guessModelParameters,
            [ExcelArgument(Description = "constant continuously compounded rate of return")] double riskFreeRate,
            [ExcelArgument(Description = "the current risky asset price")] double underlyingPrice,
            [ExcelArgument(Description = "the SSVI alpha from ATM calibration")] double alpha,
            [ExcelArgument(Description = "the SSVI beta from ATM calibration")] double beta,
            [ExcelArgument(Description = "An Nx1 XL Range of option strikes")] object strikes,
            [ExcelArgument(Description = "An Nx1 XL Range of option maturities")] object maturities,
            [ExcelArgument(Description = "An Nx1 XL Range of option either C or P for Calls and Puts")] object type,
            [ExcelArgument(Description = "An Nx1 XL Range of option prices")]object observedPrices,
            [ExcelArgument(Description = "Accuracy parameter for calibration")]double accuracy,
            [ExcelArgument(Description = "Number of iterations parameter for calibration")]int maxIterations)
        {
            try
            {
                if (ExcelDnaUtil.IsInFunctionWizard()) return null;

                double[] strikesArray = XLInterfaceBase.ConvertToVector<double>(strikes);
                double[] maturitesArray = XLInterfaceBase.ConvertToVector<double>(maturities);
                string[] typeArray = XLInterfaceBase.ConvertToVector<string>(type);
                double[] observedPricesArray = XLInterfaceBase.ConvertToVector<double>(observedPrices);

                if (strikesArray.Length != maturitesArray.Length
                    || maturitesArray.Length != typeArray.Length
                    || typeArray.Length != observedPricesArray.Length)
                {
                    XLInterfaceBase.AddErrorMessage("CalibrateSsviParameters: strikes, maturites, types and observedPrices must be of same length.");
                    return null;
                }

                KeyValuePair<string, double>[] paramPairs = XLInterfaceBase.ConvertToKeyValuePairs(guessModelParameters);
                if (paramPairs.Length != 3)
                {
                    XLInterfaceBase.AddErrorMessage("CalibrateSsviParameters: guessModelParameters must be three key - value pairs.");
                    return null;
                }

                double eta = 0; double gamma = 0; double rho = 0;
                for (int paramIdx = 0; paramIdx < paramPairs.Length; ++paramIdx)
                {
                    KeyValuePair<string, double> pair = paramPairs[paramIdx];
                    string key = pair.Key;
                    double value = pair.Value;
                    if (key.Equals("eta", StringComparison.CurrentCultureIgnoreCase))
                        eta = value;
                    else if (key.Equals("gamma", StringComparison.CurrentCultureIgnoreCase))
                        gamma = value;
                    else if (key.Equals("rho", StringComparison.CurrentCultureIgnoreCase))
                        rho = value;
                    else
                    {
                        XLInterfaceBase.AddErrorMessage("CalibrateSsviParameters: guessModelParameters: unknown key: " + key);
                        return null;
                    }
                }

                SsviCalibrator ssviCalibrator = new SsviCalibrator(riskFreeRate, underlyingPrice, accuracy, maxIterations);
                ssviCalibrator.SetGuessParameters(alpha, beta, gamma, eta, rho);

                int numObservedOptions = strikesArray.Length;
                for (int optionIdx = 0; optionIdx < numObservedOptions; ++optionIdx)
                {
                    ssviCalibrator.AddObservedOption(maturitesArray[optionIdx], strikesArray[optionIdx], observedPricesArray[optionIdx], typeArray[optionIdx]);
                }
                ssviCalibrator.CalibrateEtaGammaRhoNonATM();
                CalibrationOutcome outcome = CalibrationOutcome.NotStarted;
                double calibrationError = 0;
                ssviCalibrator.GetCalibrationStatus(ref outcome, ref calibrationError);

                Ssvi calibratedModel = ssviCalibrator.GetCalibratedModel();

                // for output
                const int numCols = 2;
                const int numRows = 5;
                object[,] output = new object[numRows, numCols];
                output[0, 0] = "Gamma"; output[0, 1] = calibratedModel.GetGamma();
                output[1, 0] = "Eta"; output[1, 1] = calibratedModel.GetEta();
                output[2, 0] = "Rho"; output[2, 1] = calibratedModel.GetRho();

                output[3, 0] = "Minimizer Status";
                if (outcome == CalibrationOutcome.FinishedOK)
                    output[3, 1] = "OK";
                else if (outcome == CalibrationOutcome.FailedMaxItReached)
                    output[3, 1] = "Reached max. num. iterations.";
                else if (outcome == CalibrationOutcome.FailedOtherReason)
                    output[3, 1] = "Failed.";
                else
                    output[3, 1] = "Unknown outcome.";

                output[4, 0] = "Pricing error"; output[4, 1] = Math.Sqrt(calibrationError);
                return output;
            }
            catch (Exception e)
            {
                XLInterfaceBase.AddErrorMessage("CalibrateParameters: unknown error: " + e.Message);
            }

            return null;
        }



        [ExcelFunction(IsThreadSafe = true,
            Category = "Dupire",
            Description = "Uses Monte Carlo to price an European option using Dupire local volatility surface.")]
        public static object DupireSsviEuropeanPriceWithMC([ExcelArgument(Description = "the current risky asset price")] double S0,
                                                [ExcelArgument(Description = "constant continuously compounded rate of return")] double riskFreeRate,
                                                [ExcelArgument(Description = "parameter in theta(t) := alpha^2(e^(beta^2 t)  - 1)")]double alpha,
                                                [ExcelArgument(Description = "parameter in theta(t) := alpha^2(e^(beta^2 t)  - 1)")]double beta,
                                                [ExcelArgument(Description = "SSVI parameter")]double gamma,
                                                [ExcelArgument(Description = "SSVI parameter")]double eta,
                                                [ExcelArgument(Description = "SSVI parameter")]double rho,
                                                [ExcelArgument(Description = "option maturity (time to expiry)")]double maturity,
                                                [ExcelArgument(Description = "option strike")]double strike,
                                                [ExcelArgument(Description = "option type, 'C' for call 'P' for put")]string type,
                                                [ExcelArgument(Description = "number of time steps e.g. 252")]int numSteps,
                                                [ExcelArgument(Description = "number of trials e.g. 10 000")]int numTrials)
        {
            if (ExcelDnaUtil.IsInFunctionWizard()) return null;

            try
            {
                double[] ssviParams = new double[] { alpha, beta, gamma, eta, rho };
                MonteCarloPricingLocalVol pricer = new MonteCarloPricingLocalVol(riskFreeRate, numTrials, numSteps, ssviParams);
                if (type == "Call" || type == "call" || type == "C" || type == "c")
                    return pricer.CalculateEurCallOptionPrice(S0, strike, maturity);
                else if (type == "Put" || type == "put" || type == "P" || type == "p")
                    return pricer.CalculateEurPutOptionPrice(S0, strike, maturity);
                else
                    throw new ArgumentException("Input must be either Call, call, C, c, Put, put, P or p.");
            }
            catch (Exception e)
            {
                XLInterfaceBase.AddErrorMessage("DupireSsviEuropeanPriceWithMC error: " + e.Message);
            }
            return null;
        }

        [ExcelFunction(IsThreadSafe = true,
            Category = "Dupire",
            Description = "Uses Monte Carlo to price an Asian call / put option using Dupire local volatility surface.")]
        public static object DupireSsviAsianOptionPriceMC(
            [ExcelArgument(Description = "the current risky asset price")] double S0,
            [ExcelArgument(Description = "constant continuously compounded rate of return")] double riskFreeRate,
            [ExcelArgument(Description = "parameter in theta(t) := alpha^2(e^(beta^2 t)  - 1)")]double alpha,
            [ExcelArgument(Description = "parameter in theta(t) := alpha^2(e^(beta^2 t)  - 1)")]double beta,
            [ExcelArgument(Description = "SSVI parameter")]double gamma,
            [ExcelArgument(Description = "SSVI parameter")]double eta,
            [ExcelArgument(Description = "SSVI parameter")]double rho,
            [ExcelArgument(Description = "option maturity (time to expiry)")]double maturity,
            [ExcelArgument(Description = "option strike")]double strike,
            [ExcelArgument(Description = "an Nx1 XL range of values over which the average is taken")]object monitoringTimes,
            [ExcelArgument(Description = "option type, 'C' for call 'P' for put")]string type,
            [ExcelArgument(Description = "number of time steps e.g. 252")]int numSteps,
            [ExcelArgument(Description = "number of trials e.g. 10 000")]int numTrials)

        {
            try
            {
                double[] monitoringT = XLInterfaceBase.ConvertToVector<double>(monitoringTimes);
                double[] ssviParams = new double[] { alpha, beta, gamma, eta, rho };
                MonteCarloPricingLocalVol McPricer = new MonteCarloPricingLocalVol(riskFreeRate, numTrials, numSteps, ssviParams);
                if (type == "Call" || type == "call" || type == "C" || type == "c")
                    return McPricer.CalculateAsianCallOptionPrice(S0, strike, monitoringT, maturity);
                else if (type == "Put" || type == "put" || type == "P" || type == "p")
                    return McPricer.CalculateAsianPutOptionPrice(S0, strike, monitoringT, maturity);
                else
                    throw new ArgumentException("Input must be either Call, call, C, c, Put, put, P or p.");
            }
            catch (Exception e)
            {
                XLInterfaceBase.AddErrorMessage("DupireSsviAsianOptionPriceMC error: " + e.Message);
            }
            return null;
        }

        [ExcelFunction(IsThreadSafe = true,
            Category = "Dupire",
            Description = "Uses Monte Carlo to price a Lookback option using Dupire local volatility surface.")]
        public static object DupireSsviLookbackOptionPriceMC(
            [ExcelArgument(Description = "the current risky asset price")] double S0,
            [ExcelArgument(Description = "constant continuously compounded rate of return")] double riskFreeRate,
            [ExcelArgument(Description = "parameter in theta(t) := alpha^2(e^(beta^2 t)  - 1)")]double alpha,
            [ExcelArgument(Description = "parameter in theta(t) := alpha^2(e^(beta^2 t)  - 1)")]double beta,
            [ExcelArgument(Description = "SSVI parameter")]double gamma,
            [ExcelArgument(Description = "SSVI parameter")]double eta,
            [ExcelArgument(Description = "SSVI parameter")]double rho,
            [ExcelArgument(Description = "option maturity (time to expiry)")]double maturity,
            [ExcelArgument(Description = "number of time steps e.g. 252")]int numSteps,
            [ExcelArgument(Description = "number of trials e.g. 10 000")]int numTrials)
        {

            if (ExcelDnaUtil.IsInFunctionWizard()) return null;
            try
            {
                double[] ssviParams = new double[] { alpha, beta, gamma, eta, rho };
                MonteCarloPricingLocalVol McPricer = new MonteCarloPricingLocalVol(riskFreeRate, numTrials, numSteps, ssviParams);
                return McPricer.CalculateLookbackOptionPrice(S0, maturity);
            }
            catch (Exception e)
            {
                XLInterfaceBase.AddErrorMessage("DupireSsviLookbackOptionPriceMC error: " + e.Message);
            }
            return null;
        }

        [ExcelFunction(IsThreadSafe = true,
           Category = "Dupire",
           Description = "Uses Monte Carlo to price a Barrier option using Dupire local volatility surface.")]
        public static object DupireSsviBarrierPriceWithMC([ExcelArgument(Description = "the current risky asset price")] double S0,
                                               [ExcelArgument(Description = "constant continuously compounded rate of return")] double riskFreeRate,
                                               [ExcelArgument(Description = "parameter in theta(t) := alpha^2(e^(beta^2 t)  - 1)")]double alpha,
                                               [ExcelArgument(Description = "parameter in theta(t) := alpha^2(e^(beta^2 t)  - 1)")]double beta,
                                               [ExcelArgument(Description = "SSVI parameter")]double gamma,
                                               [ExcelArgument(Description = "SSVI parameter")]double eta,
                                               [ExcelArgument(Description = "SSVI parameter")]double rho,
                                               [ExcelArgument(Description = "option maturity (time to expiry)")]double maturity,
                                               [ExcelArgument(Description = "option strike")]double strike,
                                               [ExcelArgument(Description = "option barrier")]double barrier,
                                               [ExcelArgument(Description = "option type, 'C' for call 'P' for put")]string CallOrPut,
                                               [ExcelArgument(Description = "option type, 'U' for up 'D' for down")]string UpOrDown,
                                               [ExcelArgument(Description = "option type, 'I' for in 'O' for out")]string InOrOut,
                                               [ExcelArgument(Description = "number of time steps e.g. 252")]int numSteps,
                                               [ExcelArgument(Description = "number of trials e.g. 10 000")]int numTrials)
        {
            if (ExcelDnaUtil.IsInFunctionWizard()) return null;

            try
            {
                double[] ssviParams = new double[] { alpha, beta, gamma, eta, rho };
                MonteCarloPricingLocalVol pricer = new MonteCarloPricingLocalVol(riskFreeRate, numTrials, numSteps, ssviParams);
                if (CallOrPut == "C")
                    return pricer.CalculateCallBarrierOptionPrice(S0, strike, maturity, UpOrDown, InOrOut, barrier);
                else if (CallOrPut == "P")
                    return pricer.CalculatePutBarrierOptionPrice(S0, strike, maturity, UpOrDown, InOrOut, barrier);
                else
                    throw new ArgumentException("Input must be either C or P.");
            }
            catch (Exception e)
            {
                XLInterfaceBase.AddErrorMessage("DupireSsviEuropeanPriceWithMC error: " + e.Message);
            }
            return null;
        }

        [ExcelFunction(IsThreadSafe = true,
            Category = "Dupire",
            Description = "Uses Monte Carlo to price a Lookback option using Dupire local volatility surface.")]
        public static double[,] DupireSsviGeneratePaths(
            [ExcelArgument(Description = "the current risky asset price")] double S0,
            [ExcelArgument(Description = "constant continuously compounded rate of return")] double riskFreeRate,
            [ExcelArgument(Description = "parameter in theta(t) := alpha^2(e^(beta^2 t)  - 1)")]double alpha,
            [ExcelArgument(Description = "parameter in theta(t) := alpha^2(e^(beta^2 t)  - 1)")]double beta,
            [ExcelArgument(Description = "SSVI parameter")]double gamma,
            [ExcelArgument(Description = "SSVI parameter")]double eta,
            [ExcelArgument(Description = "SSVI parameter")]double rho,
            [ExcelArgument(Description = "option maturity (time to expiry)")]double maturity,
            [ExcelArgument(Description = "number of time steps e.g. 252")]int numSteps,
            [ExcelArgument(Description = "number of trials e.g. 10 000")]int numTrials)
        {

            if (ExcelDnaUtil.IsInFunctionWizard()) return null;
            try
            {
                double[] ssviParams = new double[] { alpha, beta, gamma, eta, rho };
                double[,] paths = MonteCarloPaths.GenerateMcPaths(numTrials, numSteps, maturity, riskFreeRate, S0, ssviParams);
                return paths;
            }
            catch (Exception e)
            {
                XLInterfaceBase.AddErrorMessage("DupireSsviGeneratePaths error: " + e.Message);
            }
            return null;
        }
    }
}
