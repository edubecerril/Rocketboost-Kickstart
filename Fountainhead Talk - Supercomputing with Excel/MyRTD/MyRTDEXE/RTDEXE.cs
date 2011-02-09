using System;
using System.Runtime.InteropServices;

namespace MyRTD
{
    /// <summary>
    /// Out-of-process (.EXE) Excel RTD server.
    /// </summary>
    [Guid("63159552-FC4B-413f-A4C0-12E8B2F10C4B"),
     ProgId("MyRTDEXE.RTD"),
     ComVisible(true)]
    public partial class RTD : IRtdServer
    {
        private const String rtdName = "MyRTDEXE.RTD";
    }
}