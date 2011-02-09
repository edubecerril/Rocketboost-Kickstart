using System;
using System.Runtime.InteropServices;

namespace Demo
{
    /// <summary>
    /// In-process (.DLL) Excel RTD server.
    /// </summary>
    [Guid("4BC5FA82-4425-4685-A14F-138AA66F962F"),
     ProgId("Demo.RTD"),
     ComVisible(true)]
    public partial class RTD : IRtdServer
    {
        private const String rtdName = "Demo.RTD";
    }
}