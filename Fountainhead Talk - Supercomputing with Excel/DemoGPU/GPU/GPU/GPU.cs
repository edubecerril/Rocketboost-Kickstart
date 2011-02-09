using System;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

namespace Demo
{
    /// <summary>
    /// Demo GPU UDF functions.
    /// </summary>
    [Guid("7F1843AF-95B5-403f-B244-30B86A1A3D30"),
     ClassInterface(ClassInterfaceType.AutoDual),
     ProgId("Demo.GPU"),
     ComVisible(true)]
    public class GPU : IDTExtensibility2
    {
        #region Fields

        /// <summary>
        /// Singleton instance of this class.
        /// </summary>
        private static readonly GPU instance = new GPU();

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
        public static GPU Instance
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
            if (n >= 1 && n <= 50)
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
            else
            {
                results = new int[1];

                results[0] = -1;
            }

            return results;
        }

        public Object udfFibonacciNth(int n)
        {
            int[] numbers = new int[n];
            numbers = Fibonacci(n);

            return numbers[n - 1];
        }

        public Object gpuIsamax(double[] x)
        {
            Cublas blas = new Cublas();

            float[] vector = new float[x.Length];

            for (int i=0; i<x.Length; i++)
            {
                vector[i] = Convert.ToSingle(x[i]);
            }

            return blas.Isamax(vector);
        }

        #endregion // UDFs
    }
}