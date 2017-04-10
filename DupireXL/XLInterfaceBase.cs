using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using ExcelDna.Integration;

namespace DupireXL
{
    /// <summary>
    /// Class that provides basic functionality including error message reporting.
    /// In theory, you shouldn't need to change this part.
    /// </summary>
    public class XLInterfaceBase
    {
        static LinkedList<string> errorMessages;


        static XLInterfaceBase()
        {
            errorMessages = new LinkedList<string>();    
        }

        public static void AddErrorMessage(string m)
        {
            errorMessages.AddFirst(m);
        }

        [ExcelFunction(Description = "Clear Error Messages for DupireXL.")]
        public static void ClearErrorMessages()
        {
            errorMessages.Clear();
        }

        [ExcelFunction(Description = "Display Error Messages for DupireXL.")]
        public static object[,] GetLatestErrors([ExcelArgument(Description = "number of error messages to display")] int number)
        {
            if (number <= 0)
            {
                string[,] toDisplay = new string[1, 1];
                toDisplay[0, 0] = "GetLatestErrors: You must enter a positive number.";
                return toDisplay;
            }
            else
            {
                string[,] toDisplay = new string[number, 1];
                int msgIdx = 0;
                foreach (string errorMsg in errorMessages)
                {
                    toDisplay[msgIdx, 0] = errorMsg;
                    ++msgIdx;
                    if (msgIdx >= number)
                        break;
                }

                for (; msgIdx < number; ++msgIdx)
                {
                    toDisplay[msgIdx, 0] = "";
                }
                return toDisplay;
            }
        }

        internal static bool ItIsCall(string s)
        {
            if (s.Equals("call", StringComparison.CurrentCultureIgnoreCase) || s == "C" || s == "c")
                return true;
            else if (s.Equals("put", StringComparison.CurrentCultureIgnoreCase) || s == "P" || s == "p")
                return false;
            else
            {
                throw new System.ArgumentException("Type must be one of call / put or c / p or C / P; this: " + s + " was not recognised.");

            }
        }

        

        internal static T ConvertTo<T>(object In)
        {
            try
            {
                return (T)In;
            }
            catch (Exception e)
            {
                errorMessages.AddFirst("Could not convert object to "+typeof(T).ToString());
                throw e;
            }
        }



        internal static T[] ConvertToVector<T>(object In)
        {
            T[] V;
            try
            {
                object[] InVec;               
                if (In.GetType() == typeof(object) || In.GetType() == typeof(T))
                {
                    V = new T[1];
                    V[0] = ConvertTo<T>(In);
                    return V;
                }
                else if (In.GetType() == typeof(object[]))
                {
                    InVec = (object[])In;
                    int length = InVec.GetLength(0);
                    V = new T[length];
                    for (int i = 0; i < length; i++)
                    {
                        V[i] = ConvertTo<T>(InVec[i]);
                    }
                    return V;
                }
                else if (In.GetType() == typeof(object[,]))
                {
                    object[,] InM = (object[,])In;
                    int rows = InM.GetLength(0);
                    V = new T[rows];
                    for (int i = 0; i < rows; i++)
                    {
                        V[i] = ConvertTo<T>(InM[i, 0]);
                    }
                    return V;
                }
                else
                {
                    errorMessages.AddFirst("Could not convert input to vector of type "+typeof(T).ToString());
                    return null;
                }


            }
            catch (Exception)
            {
                errorMessages.AddFirst("Could not convert input to vector of type " + typeof(T).ToString());
                return null;
            }
        }



        internal static T[,] ConvertToMatrix<T>(object In)
        {
            T[,] M;
            try
            {
                object[,] InM = (object[,])In;
                int rows = InM.GetLength(0);
                int cols = InM.GetLength(1);
                
                M = new T[rows, cols];
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                        M[i, j] = ConvertTo<T>(InM[i, j]);
                }
                return M;
            }
            catch (Exception)
            {
                errorMessages.AddFirst("Could not convert input to matrix of type " + typeof(T).ToString());
                return null;
            }
        }

        internal static KeyValuePair<string, double>[] ConvertToKeyValuePairs(object In)
        {
            KeyValuePair<string, double>[] keyValPairs;
            try
            {
                object[,] In2D = (object[,])In;
                int rows = In2D.GetLength(0);
                int cols = In2D.GetLength(1);
                if (cols != 2)
                {
                    Console.WriteLine("Need two colums!");
                    return null;
                }
                keyValPairs = new KeyValuePair<string, double>[rows];
                for (int i = 0; i < rows; i++)
                {
                    string key = ConvertTo<string>(In2D[i, 0]);
                    double value = ConvertTo<double>(In2D[i, 1]);
                    KeyValuePair<string, double> pair = new KeyValuePair<string, double>(key, value);
                    keyValPairs[i] = pair;
                }
                return keyValPairs;
            }
            catch (Exception)
            {
                errorMessages.AddFirst("Could create key - value pair.");
                return null;
            }

        }
    }
}
