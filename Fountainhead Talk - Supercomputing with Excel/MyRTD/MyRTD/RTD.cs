using System;
using System.Windows.Forms;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace MyRTD
{
    /// <summary>
    /// RTD server.
    /// </summary>
    public partial class RTD : IRtdServer
    {
        #region Fields

        /// <summary>
        /// Callback to Excel to notify it that updated data is available.
        /// </summary>
        private IRTDUpdateEvent m_callback;

        /// <summary>
        /// Timer object to update topic with data.
        /// </summary>
        private Timer m_timer;

        /// <summary>
        /// Topic ID. This simple RTD supports only one topic.
        /// </summary>
        private int m_topic;

        #endregion // Fields

        #region IRtdServer Interface

        #region ServerStart

        /// <summary>
        /// Called when Excel RTD connects to this automation component.
        /// </summary>
        /// <param name="callbackObject">[In] This is the object through which
        /// this RTD server notifies Excel that real-time data updates are
        /// available.</param>
        /// <returns>1 to denote "success".</returns>
        public int ServerStart(IRTDUpdateEvent callbackObject)
        {
            // Used by this component to notify Excel that
            // new data is available.
            m_callback = callbackObject;

            // Set timer to trigger every 2 seconds.
            m_timer = new Timer();
            m_timer.Tick += new EventHandler(TimerEventHandler);
            m_timer.Interval = 2000; // 2 seconds

            return 1;
        }

        #endregion // ServerStart

        #region ServerTerminate

        /// <summary>
        /// Called when Excel RTD disconnects from this automation component.
        /// </summary>
        public void ServerTerminate()
        {
            if (null != m_timer)
            {
                m_timer.Dispose();
                m_timer = null;
            }
        }

        #endregion // ServerTerminate

        #region ConnectData

        /// <summary>
        /// Called each time Excel RTD has a unique topic to request.
        /// </summary>
        /// <param name="topicId">[In] Excel's internal topic ID.</param>
        /// <param name="topics">[In] List of topics (parameters) passed by
        /// Excel RTD function call.</param>
        /// <param name="newValue">[In/Out] On input the new_value parameter
        /// indicates whether Excel already has a value to initially display.
        /// So if Excel has a cached value to display then new_value will be
        /// false. If Excel does not have a cached value to display then
        /// new_value will be true, indicating that a new value is needed.
        /// On output the new_value parameter indicates whether Excel should
        /// use the returned value or not. If new_value is false then Excel
        /// will ignore the value returned by ConnectData(). If it doesn’t
        /// have a previously cached value then it displays the “#N/A”
        /// warning. Of course this is replaced with an actual value once
        /// it receives one via RefreshData(). If new_value is true then Excel
        /// will immediately replace whatever value it may already have with
        /// the value returned by ConnectData().
        /// <returns>Topic data value, if available. Otherwise, it returns
        /// a default message used to indicate that a computation is
        /// under way. Usually something like "RB.XL: Calculating ...".</returns>
        public object ConnectData(int topicId, ref object[] topics, ref bool newValue)
        {
            newValue = true; // Do not allow Excel to use cached value for topic.

            m_topic = topicId;
            m_timer.Start();

            return GetTime();
        }

        #endregion // ConnectData

        #region DisconnectData

        /// <summary>
        /// Unsubscribe from a topic. This notification comes from
        /// Excel when it no longer has any RTD() entries in cells
        /// for the topic in question.
        /// </summary>
        /// <param name="topicId">[In] Excel's internal topic ID.</param>
        public void DisconnectData(int topicId)
        {
            m_timer.Stop();
        }

        #endregion // DisconnectData

        #region RefreshData

        /// <summary>
        /// Excel will call this method to get new data.
        /// </summary>
        /// <param name="topicCount">[Out] Number of topics for
        /// which data is being returned.</param>
        /// <returns>Array of topic IDs and associated data.</returns>
        public object[,] RefreshData(ref int topicCount)
        {
            object[,] data = new object[2,1];
            data[0, 0] = m_topic;
            data[1, 0] = GetTime();

            topicCount = 1;

            m_timer.Start();

            return data;
        }

        #endregion // RefreshData

        #region Heartbeat

        /// <summary>
        /// Called by Excel from time to time to check this RTD server is alive.
        /// </summary>
        /// <returns>1 to indicate RTD server is alive.</returns>
        public int Heartbeat()
        {
            return 1;
        }

        #endregion // Heartbeat

        #endregion // IRtdServer Interface

        private void TimerEventHandler(object sender,
                                       EventArgs args)
        {
            m_timer.Stop();
            m_callback.UpdateNotify();
        }

        private string GetTime()
        {
            return rtdName + " : " + DateTime.Now.ToString("hh:mm:ss:ff");
        }
    }
}
