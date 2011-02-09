using System;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

namespace Demo
{
    /// <summary>
    /// Demo UDF functions.
    /// </summary>
    [Guid("846A2806-5064-4a5c-96F2-A0E5E0B74656"),
     ClassInterface(ClassInterfaceType.AutoDual),
     ProgId("Demo.UDF"),
     ComVisible(true)]
    public class UDF : IDTExtensibility2
    {
        #region Constants

        public const int MIN_N = 1;
        public const int MAX_N = 50;
        public const int ERROR = -1;

        #endregion

        #region Fields

        /// <summary>
        /// Singleton instance of this class.
        /// </summary>
        private static readonly UDF instance = new UDF();

        /// <summary>
        /// Data type to denote a missing argument. C# does not
        /// handle missing arguments when calling methods from
        /// the Excel object model and so needs to use this
        /// dummy argument.
        /// </summary>
        private static readonly object missing = Type.Missing;

        /// <summary>
        /// Reference to a running instance of Excel.
        /// </summary>
        private static Application m_excel;

        #endregion // Fields

        #region Register/Unregister

        public virtual void OnConnection(object app,
                                         ext_ConnectMode mode, object add_inst,
                                         ref Array custom)
        {
            m_excel = (Application)app;
        }

        public virtual void OnDisconnection(ext_DisconnectMode mode,
                                            ref Array custom)
        {
        }

        public virtual void OnAddInsUpdate(ref Array custom)
        {
        }

        public virtual void OnBeginShutdown(ref Array custom)
        {
        }

        public virtual void OnStartupComplete(ref Array custom)
        {
        }

        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            Registry.ClassesRoot.CreateSubKey(
                GetSubKeyName(t, "Programmable"));
            RegistryKey key = Registry.ClassesRoot.OpenSubKey(
                GetSubKeyName(t, "InprocServer32"), true);
            key.SetValue("", String.Format("{0}\\mscoree.dll",
                                           Environment.SystemDirectory),
                         RegistryValueKind.String);
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t)
        {
            Registry.ClassesRoot.DeleteSubKey(
                GetSubKeyName(t, "Programmable"));
        }

        private static string GetSubKeyName(Type t, String sub_key)
        {
            return String.Format("CLSID\\{{{0}}}\\{1}",
                                 t.GUID.ToString().ToUpper(), sub_key);
        }

        #endregion // Register/Unregister

        #region Supporting methods

        /// <summary>
        /// Property to support singleton pattern for this class.
        /// </summary>
        public static UDF Instance
        {
            get { return instance; }
        }

        #endregion // Supporting methods

        #region UDFs

        /// <summary>
        /// Generates the first 'n' numbers in the Fibonacci
        /// sequence.
        /// </summary>
        /// <param name="n">Number of terms in the sequence
        /// to generate.</param>
        /// <returns>First n numbers in the Fibonacci
        /// sequence.</returns>
        private int[] Fibonacci(int n)
        {
            int[] results;

            // Limit the length of the sequence
            // that can be requested.
            if (n >= MIN_N && n <= MAX_N)
            {
                results = new int[n];

                int auxiliar = 0;
                int previous = 0;
                int current = 1;
                int i = 0;
                while (i < n)
                {
                    if (i == 0)
                    {
                        results[i] = 0;
                    }
                    else if (i == 1)
                    {
                        results[i] = 1;
                    }
                    else
                    {
                        auxiliar = previous;
                        previous = current;
                        current = auxiliar + current;
                        results[i] = current;
                    }

                    i++;
                }
            }
            else // Return an error code.
            {
                results = new int[1];

                results[0] = ERROR;
            }

            return results;
        }

        public Object FibonacciNth(int n)
        {
            int[] numbers = new int[n];
            numbers = Fibonacci(n);

            if (numbers.Length==1 && numbers[0]==ERROR)
            {
                return "N-th term requested must be in the range 1..50";
            }
            else
            {
                return numbers[n - 1];
            }
        }

        #endregion // UDFs
    }
}