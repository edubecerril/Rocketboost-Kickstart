using System;
using System.Runtime.InteropServices;

namespace MyRTD
{
    /// <summary>
    /// In-process (.DLL) Excel RTD server.
    /// </summary>
    [Guid("8B8172AC-EA72-4111-AC8C-B891F1F944A1"),
     ProgId("MyRTD.RTD"),
     ComVisible(true)]
    public partial class RTD : IRtdServer
    {
        private const String rtdName = "MyRTD.RTD";
    }
}