using System;
using System.Runtime.InteropServices;
using cublasStatus = System.UInt32;

namespace Demo
{
    /// <summary>
    /// Class that wraps Nvidia's CUBLAS functions.
    /// </summary>
    public class Cublas : IDisposable
    {
        //========================================

        #region Constants

        // Symbolic constants defined in cublas.h header file.
        /// <summary>Const: CUBLAS_STATUS_SUCCESS</summary>
        public const cublasStatus CUBLAS_STATUS_SUCCESS = 0x00000000;
        /// <summary>Const: CUBLAS_STATUS_NOT_INITIALIZED</summary>
        public const cublasStatus CUBLAS_STATUS_NOT_INITIALIZED = 0x00000001;
        /// <summary>Const: CUBLAS_STATUS_ALLOC_FAILED</summary>
        public const cublasStatus CUBLAS_STATUS_ALLOC_FAILED = 0x00000003;
        /// <summary>Const: CUBLAS_STATUS_INVALID_VALUE</summary>
        public const cublasStatus CUBLAS_STATUS_INVALID_VALUE = 0x00000007;
        /// <summary>Const: CUBLAS_STATUS_ARCH_MISMATCH</summary>
        public const cublasStatus CUBLAS_STATUS_ARCH_MISMATCH = 0x00000008;
        /// <summary>Const: CUBLAS_STATUS_MAPPING_ERROR</summary>
        public const cublasStatus CUBLAS_STATUS_MAPPING_ERROR = 0x0000000B;
        /// <summary>Const: CUBLAS_STATUS_EXECUTION_FAILED</summary>
        public const cublasStatus CUBLAS_STATUS_EXECUTION_FAILED = 0x0000000D;
        /// <summary>Const: CUBLAS_STATUS_INTERNAL_ERROR</summary>
        public const cublasStatus CUBLAS_STATUS_INTERNAL_ERROR = 0x0000000E;

        /// <summary>
        /// Given a cublas status code returns the symbolic name for
        /// that code.
        /// </summary>
        /// <param name="status">cublasStatus code.</param>
        /// <returns>Symbolic name for the status code.</returns>
        public static String cublasStatusSymbolicName(cublasStatus status)
        {
            switch (status)
            {
                case CUBLAS_STATUS_SUCCESS:
                    return "CUBLAS_STATUS_SUCCESS";
                case CUBLAS_STATUS_NOT_INITIALIZED:
                    return "CUBLAS_STATUS_NOT_INITIALIZED";
                case CUBLAS_STATUS_ALLOC_FAILED:
                    return "CUBLAS_STATUS_ALLOC_FAILED";
                case CUBLAS_STATUS_INVALID_VALUE:
                    return "CUBLAS_STATUS_INVALID_VALUE";
                case CUBLAS_STATUS_ARCH_MISMATCH:
                    return "CUBLAS_STATUS_ARCH_MISMATCH";
                case CUBLAS_STATUS_MAPPING_ERROR:
                    return "CUBLAS_STATUS_MAPPING_ERROR";
                case CUBLAS_STATUS_EXECUTION_FAILED:
                    return "CUBLAS_STATUS_EXECUTION_FAILED";
                case CUBLAS_STATUS_INTERNAL_ERROR:
                    return "CUBLAS_STATUS_INTERNAL_ERROR";
                default:
                    return "UNKNOWN";
            }
        }

        #endregion // Constants

        //========================================

        #region Fields

        /// <summary>
        /// Singleton instance of this class.
        /// </summary>
        public static Cublas Instance { get; private set; }

        #endregion // Fields

        //========================================

        #region Constructors and Destructors

        /// <summary>
        /// Static constructor - called only once.
        /// </summary>
        static Cublas()
        {
            Instance = new Cublas();

            cublasStatus status = cublasInit();
            TestForError(status, "Call to CublasInit() failed.");
        }

        /// <summary>
        /// Dispose of unmanaged resources for this class.
        /// </summary>
        public void Dispose()
        {
            cublasStatus status = cublasShutdown();
            TestForError(status, "Call to CublasShutdown() failed.");
        }

        /// <summary>
        /// Cleanup for managed code of class.
        /// </summary>
        ~Cublas()
        {
        }

        #endregion // Constructors and Destructors

        //========================================

        #region CUBLAS helper functions

        // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

        // These two functions are private because we
        // want the assembly to manage initialization
        // and shutdown of CUBLAS at load/unload time
        // rather than rely on the caller to do this.

        /// <summary>
        /// Initialize CUBLAS library.
        /// </summary>
        /// <returns>cublasStatus code.</returns>
        [DllImport("cublas.dll", CallingConvention = CallingConvention.StdCall)]
        private static extern cublasStatus cublasInit();

        /// <summary>
        /// Shutdown CUBLAS library.
        /// </summary>
        /// <returns>cublasStatus code.</returns>
        [DllImport("cublas.dll", CallingConvention = CallingConvention.StdCall)]
        private static extern cublasStatus cublasShutdown();

        // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

        /// <summary>
        /// Creates an object in GPU memory.
        /// </summary>
        /// <remarks>
        /// On exit the out parameter devicePtr is a pointer
        /// to memory on the GPU device and so cannot be
        /// dereferenced by code running on the CPU.
        /// </remarks>
        /// <param name="n">Number of elements.</param>
        /// <param name="elemSize">Size (in bytes) of a single element.</param>
        /// <param name="devicePtr">Pointer to memory on the GPU device.</param>
        /// <returns>cublasStatus code.</returns>
        [DllImport("cublas.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern cublasStatus cublasAlloc(
            int n,
            int elemSize,
            out IntPtr devicePtr
            );

        /// <summary>
        /// Returns the last error that occurred on invocation
        /// of any of the CUBLAS core functions.
        /// </summary>
        [DllImport("cublas.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern cublasStatus cublasGetError();

        /// <summary>
        /// Frees an object previously allocated memory using
        /// culblasAlloc().
        /// </summary>
        /// <param name="devicePtr">Pointer to memory on the GPU device.</param>
        /// <returns>cublasStatus code.</returns>
        [DllImport("cublas.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern cublasStatus cublasFree(
            IntPtr devicePtr
            );

        /// <summary>
        /// Copies 'n' elements from a vector 'x' in CPU memory
        /// to a vector 'y' in GPU memory.
        /// </summary>
        /// <param name="n">Number of elements.</param>
        /// <param name="elemSize">Size of elements.</param>
        /// <param name="x">Vector in CPU memory.</param>
        /// <param name="incx">Storage spacing between consecutive
        /// elements for source vector x[].</param>
        /// <param name="y">Vector in GPU memory.</param>
        /// <param name="incy">Storage spacing between consecutive
        /// elements for destination vector y[].</param>
        /// <returns>cublasStatus code.</returns>
        [DllImport("cublas.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern cublasStatus cublasSetVector(
            int n,
            int elemSize,
            float[] x,
            int incx,
            IntPtr y,
            int incy
            );

        #endregion // CUBLAS helper functions

        //========================================

        #region Single-Precision BLAS1 Functions

        /// <summary>
        /// Rocketboost function wrapper for cublasIsmax().
        /// </summary>
        /// <param name="n">Number of elements in the input vector.</param>
        /// <param name="x">Single precision vector with n elements.</param>
        /// <param name="incx">Storage spacing between elements of x.</param>
        /// <returns>Returns the smallest index (returns zero if n&lt;=0 or incx&lt;=0).</returns>
        [DllImport("cublas.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern int cublasIsamax(
            int n,
            IntPtr x,
            int incx);

        #endregion // Single-Precision BLAS1 Functions

        //========================================

        #region Single-Precision Complex BLAS1 Functions

        #endregion // Single-Precision Complex BLAS1 Functions

        //========================================

        #region Double-Precision BLAS1 Functions

        #endregion // Double-Precision BLAS1 Functions

        //========================================

        #region Double-Precision Complex BLAS1 Functions

        #endregion // Double-Precision Complex BLAS1 Functions

        //========================================

        #region Single-Precision BLAS2 Functions

        #endregion // Single-Precision BLAS2 Functions

        //========================================

        #region Single-Precision Complex BLAS2 Functions

        #endregion // Single-Precision Complex BLAS2 Functions

        //========================================

        #region Double-Precision BLAS2 Functions

        #endregion // Double-Precision BLAS2 Functions

        //========================================

        #region Double-Precision Complex BLAS2 Functions

        #endregion // Double-Precision Complex BLAS2 Functions

        //========================================

        #region Single-Precision BLAS3 Functions

        #endregion // Single-Precision BLAS3 Functions

        //========================================

        #region Single-Precision Complex BLAS3 Functions

        #endregion // Single-Precision Complex BLAS3 Functions

        //========================================

        #region Double-Precision BLAS3 Functions

        #endregion // Double-Precision BLAS3 Functions

        //========================================

        #region Double-Precision Complex BLAS3 Functions

        #endregion // Double-Precision Complex BLAS3 Functions

        //========================================

        #region Error handling methods

        /// <summary>
        /// Test for a general error condition. If the
        /// condition is met, throw exception.
        /// </summary>
        /// <param name="condition">Throw exception only
        /// if 'condition' is true.</param>
        /// <param name="message">Error message to
        /// attach to exception.</param>
        private static void TestForError(Boolean condition, String message)
        {
            if (condition) throw new Exception(message);
        }

        /// <summary>
        /// Test for a CUBLAS error condition. If the return
        /// status from a CUBLAS function call is not success,
        /// throw exception.
        /// </summary>
        /// <param name="status">Throw exception only
        /// if 'status' not success.</param>
        /// <param name="message">Error message to
        /// attach to exception.</param>
        private static void TestForError(cublasStatus status, String message)
        {
            if (status != CUBLAS_STATUS_SUCCESS)
            {
                String errorMessage = message + " " +
                    cublasStatusSymbolicName(status);

                throw new Exception(errorMessage);
            }
        }

        /// <summary>
        /// Test for a CUBLAS error condition. If the return
        /// status from a CUBLAS function call is not success,
        /// throw exception, but only after freeing memory
        /// previously allocated on GPU device.
        /// </summary>
        /// <param name="status">Throw exception only
        /// if 'status' not success.</param>
        /// <param name="ptr">Pointer to previously allocated
        /// GPU device memory.</param>
        /// <param name="message">Error message to
        /// attach to exception.</param>
        private static void TestForError(cublasStatus status, IntPtr ptr, String message)
        {
            if (status != CUBLAS_STATUS_SUCCESS)
            {
                String errorMessage = message + " " +
                    cublasStatusSymbolicName(status);

                if (!ptr.Equals(null) && !ptr.Equals(0) && cublasFree(ptr) != CUBLAS_STATUS_SUCCESS)
                {
                    errorMessage =
                        "Call to cublasFree() failed. " + message + " " +
                        cublasStatusSymbolicName(status);
                }

                throw new Exception(errorMessage);
            }
        }

        #endregion // Error handing methods

        //========================================

        #region Compute methods (functions)

        /// <summary>
        /// Convenience function for BLAS1 function cublasIsamax().
        /// </summary>
        /// <param name="x">Single precision vector.</param>
        /// <returns>Returns the smallest index of the maximum
        /// magnitude element.
        /// </returns>
        public int Isamax(params float[] x)
        {
            TestForError(x.Length == 0, "Vector of length zero passed to Isamax() function.");

            // Allocate memory on the device.
            cublasStatus status;
            IntPtr ptr;
            status = cublasAlloc(x.Length, sizeof(float), out ptr);
            TestForError(status, "Call to cublasAlloc() failed. Unable to allocate memory on GPU device.");

            // Set vector in device memory.
            status = cublasSetVector(x.Length, sizeof(float), x, 1, ptr, 1);
            TestForError(status, ptr, "Call to cublasSetVector() failed. Unable to copy data from host to GPU device.");

            // Call CUBLAS function.
            int idx = cublasIsamax(x.Length, ptr, 1);
            status = cublasGetError();
            TestForError(status, "Call to cublasIsamax() failed.");

            // Free vector in device memory.
            status = cublasFree(ptr);
            TestForError(status, "Call to cublasFree() failed. Unable to deallocate memory on GPU device.");

            return idx;
        }

        #endregion // Compute methods (functions)

        //========================================
    }
}
