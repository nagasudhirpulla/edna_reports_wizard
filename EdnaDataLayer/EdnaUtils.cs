using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InStep.eDNA.EzDNAApiNet;

namespace EdnaDataLayer
{
    public class EdnaUtils
    {
        public static List<EdnaPointResult> FetchHistoricalPointData(EdnaPoint pnt, DateTime startTime, DateTime endTime, string strategy, int FetchPeriodicitySecs)
        {
            try
            {
                int nret = 0;
                uint s = 0;
                double dval = 0;
                DateTime timestamp = DateTime.Now;
                string status = "";
                TimeSpan period = TimeSpan.FromSeconds(FetchPeriodicitySecs);
                // History request initiation
                if (strategy == FetchStrategy.Raw)
                { nret = History.DnaGetHistRaw(pnt.PointId, startTime, endTime, out s); }
                else if (strategy == FetchStrategy.Snap)
                { nret = History.DnaGetHistSnap(pnt.PointId, startTime, endTime, period, out s); }
                else if (strategy == FetchStrategy.Average)
                { nret = History.DnaGetHistAvg(pnt.PointId, startTime, endTime, period, out s); }
                else if (strategy == FetchStrategy.Min)
                { nret = History.DnaGetHistMin(pnt.PointId, startTime, endTime, period, out s); }
                else if (strategy == FetchStrategy.Max)
                { nret = History.DnaGetHistMax(pnt.PointId, startTime, endTime, period, out s); }
                else if (strategy == FetchStrategy.Interpolated)
                { nret = History.DnaGetHistInterp(pnt.PointId, startTime, endTime, period, out s); }

                // Get history values
                List<EdnaPointResult> historyResults = new List<EdnaPointResult>();
                while (nret == 0)
                {
                    nret = History.DnaGetNextHist(s, out dval, out timestamp, out status);
                    if (status != null)
                    {
                        historyResults.Add(new EdnaPointResult(dval, status, timestamp));
                    }
                }
                return historyResults;
            }
            catch (Exception e)
            {
                // Todo send this to console printing of the dashboard
                Console.WriteLine($"Error while fetching history data of point {pnt.PointId} ({pnt.Name})");
                Console.WriteLine($"The exception is {e}");
            }
            return null;
        }
    }
}
