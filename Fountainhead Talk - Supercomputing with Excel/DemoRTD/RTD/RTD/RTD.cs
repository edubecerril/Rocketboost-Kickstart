using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Demo
{
    /// <summary>
    /// Compute RTD server.
    /// </summary>
    public partial class RTD : IRtdServer
    {
        #region ComputeJob

        /// <summary>
        /// Delegated compute job for background worker.
        /// </summary>
        /// <param name="sender">Object that triggered event.</param>
        /// <param name="DoWorkEventArgs">Arguments passed to compute job.</param>
        private delegate void ComputeJob(object sender, DoWorkEventArgs e);

        #endregion // ComputeJob

        #region JobCompleted

        /// <summary>
        /// Event handler called when a compute job completes.
        /// </summary>
        /// <param name="sender">Object that triggered event.</param>
        /// <param name="e">Information on the outcome of a completed job.</param>
        public void JobCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // There are three possible outcomes for a compute job:
            //      1) An exception was thrown.
            //      2) The job was canceled.
            //      3) The job completed successfully.
            if (e.Error != null)
            {
                // Deal with error.
            }
            else if (e.Cancelled)
            {
                // Deal with cancel.
            }
            else
            {
                // Result data.
                ComputedResult data = (ComputedResult)e.Result;

                // Update lookup table of computed results.
                if (m_topics.ContainsKey(data.topic_id)) m_topics[data.topic_id] = data.result.ToString();
            }

            // Notify Excel that new data is available.
            m_callback.UpdateNotify();
        }

        #endregion // JobCompleted

        #region ComputedResult

        /// <summary>
        /// Computed result returned by a compute job.
        /// </summary>
        private class ComputedResult
        {
            public ComputedResult(int topicID, object computed_result)
            {
                topic_id = topicID;
                result = computed_result;
            }

            public int topic_id { get; set; }

            public object result { get; set; }
        }

        #endregion // ComputedResult

        #region Fields

        /// <summary>
        /// Callback to Excel to notify it that updated data is available.
        /// </summary>
        private IRTDUpdateEvent m_callback;

        /// <summary>
        /// Background worker threads.
        /// </summary>
        private static BackgroundWorker m_worker;

        /// <summary>
        /// Lookup table of computed results.
        /// </summary>
        private Dictionary<int, String> m_topics;

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

            // Results for each topic ID.
            m_topics = new Dictionary<int, String>();

            // Background worker thread pool.s
            m_worker = new BackgroundWorker();

            return 1;
        }

        #endregion // ServerStart

        #region ServerTerminate

        /// <summary>
        /// Called when Excel RTD disconnects from this automation component.
        /// </summary>
        public void ServerTerminate()
        {
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
            if (newValue) // Excel is requesting a new value.
            {
                // Does this server have a cached value for
                // the requested topic ID?
                if (m_topics.ContainsKey(topicId))
                {
                    // Still calculating ...
                    if (m_topics[topicId] == null)
                    {
                        return "Calculating ...";
                    }
                    else // Return calculated value.
                    {
                        return m_topics[topicId];
                    }
                }
                else // No cached value. Dispatch compute job.
                {
                    // Add a new null placeholder entry in the lookup
                    // table of computed results.
                    m_topics.Add(topicId, null);

                    // At a minimum, a function name must be passed
                    // as the first topic, and usually there will be
                    // additional parameters too.
                    if (topics.Length >= 1)
                    {
                        // First topic (parameter) is the name of the function
                        // to be computed.
                        String functionName = topics.GetValue(0).ToString();

                        // Compute job delegate.
                        ComputeJob Job;

                        // Lookup function to compute.
                        switch (functionName)
                        {
                            case "Fibonacci":
                                Job = funcFibonacci;
                                break;
                            default:
                                return "Function not found.";
                        }

                        // Dispatch compute job.
                        BackgroundWorker worker = new BackgroundWorker();
                        worker.WorkerSupportsCancellation = true;
                        worker.DoWork += new DoWorkEventHandler(Job);
                        worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(JobCompleted);

                        // Include (preappend) topicID in the list of topics
                        // (parameters).
                        String[] parameters = new String[topics.Length + 1];
                        parameters[0] = topicId.ToString();
                        topics.CopyTo(parameters, 1);

                        // Pass topics (parameters) and start worker thread.
                        worker.RunWorkerAsync(parameters);

                        // Return a message informing the user that
                        // a compute job has been dispatched. So the
                        // user will see somthing like "Calculating ..."
                        // in a cell in Excel until a calculated value
                        // is returned sometime later.
                        return "Calculating ...";
                    }
                    else // Too few topics (parameters) passed.
                    {
                        m_topics[topicId] = "Too few parameters.";

                        return "Too few parameters.";
                    }
                }
            }
            else // Excel already has a cached value to display.
            {
                newValue = false;

                return 1; // Dummy return value. Ignored by Excel.
            }
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
            m_topics.Remove(topicId);
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
            object[,] data = new object[2, m_topics.Count];

            int idx = 0;
            foreach (int topic in m_topics.Keys)
            {
                data[0, idx] = topic;

                if (m_topics[topic] != null)
                {
                    data[1, idx] = m_topics[topic];
                }
                else // Still calculating ...
                {
                    data[1, idx] = "Calculating ...";
                }

                ++idx;
            }

            topicCount = m_topics.Count;

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

        #region funcFibonacci

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

        /// <summary>
        /// Function (delegate) that returns the n-th number
        /// in the Fibonacci sequence.
        /// </summary>
        /// <param name="sender">Object that delegated to this method.</param>
        /// <param name="e">Arguments (topics) as array of strings.</param>
        private void funcFibonacci(object sender, DoWorkEventArgs e)
        {
            String[] args = e.Argument as String[];

            // args are in this order: topicID, function name, parameter
            int topic_id = int.Parse(args[0]);
            int n; Int32.TryParse(args[2],out n); // parameter

            // Pause this thread for 3 seconds to simulate a "long
            // running" background job. This way "Calculating ..."
            // appears in the calling cell for a few moments before
            // the result is returned. Otherwise, there is no
            // "Calculating ..." and the result is returned instantly.
            // To make this function run at full speed just comment
            // out the next line.
            System.Threading.Thread.Sleep(10000);

            try
            {
                int[] numbers = new int[n];
                numbers = Fibonacci(n);

                ComputedResult data = new ComputedResult(topic_id, numbers[n-1]);

                e.Result = data;
            }
            catch
            {
                e.Result = "funcFibonacci() failed.";
            }
        }

        #endregion // funcFibonacci
    }
}
