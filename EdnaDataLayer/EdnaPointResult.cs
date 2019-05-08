using System;

namespace EdnaDataLayer
{
    public class EdnaPointResult
    {
        // Vale of the data point
        public double Val_ { get; set; }

        // Data Quality of the result
        public string DataQuality_ { get; set; }

        // Result TimeStamp
        public DateTime ResultTime_ { get; set; }

        // Result units
        public string Units_ { get; set; }

        public EdnaPointResult(double val, string DataQuality, DateTime ResultTime)
        {
            Val_ = val;
            DataQuality_ = DataQuality;
            ResultTime_ = ResultTime;
        }
    }
}
