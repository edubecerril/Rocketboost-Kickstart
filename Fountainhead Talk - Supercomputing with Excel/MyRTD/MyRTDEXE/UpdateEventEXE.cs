using System;
using System.Runtime.InteropServices;

namespace MyRTD
{
    /// <summary>
    /// This class is useful when testing RTD servers
    /// from clients other than Excel.
    /// 
    /// When Excel calls the RTD ServerStart() method,
    /// it passes a callback through which the RTD can
    /// notify Excel that fresh real-time data is ready
    /// -- a tap on the shoulder of Excel, so to speak.
    /// When Excel is ready, it will ask the RTD server
    /// for the data by calling the RTD's RefreshData()
    /// method.
    /// </summary>
    [Guid("BE0974AA-9060-4032-B485-7AE5DAA04DD9"),
     ProgId("MyRTDEXE.UpdateEvent"),
     ComVisible(true)]
    public partial class UpdateEvent : IRTDUpdateEvent
    {
        private const String eventName = "MyRTDEXE.UpdateEvent";
    }
}
