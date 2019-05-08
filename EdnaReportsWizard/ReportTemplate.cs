using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using EdnaDataLayer;
using OfficeOpenXml;
using VariableTime;
using Excel = Microsoft.Office.Interop.Excel;

namespace EdnaReportsWizard
{
    public class ReportTemplate
    {
        public List<EdnaPoint> EdnaPoints { get; set; }
        public int ResolutionSecs { get; set; } = 60;
        public VarTime StartTime { get; set; }
        public VarTime EndTime { get; set; }
        public string FetchStrategy { get; set; } = EdnaDataLayer.FetchStrategy.Snap;
        public string TimeStampFormat { get; set; } = "dd-MM-yyyy HH:mm:ss";
        public string TimeStampHeading { get; set; } = "Timestamp";
        public int NumRowsToSkip { get; set; } = 0;
        public string reportFolderPath { get; set; } = @"C:\Users\Nagasudhir\Downloads";
        public string ReportName { get; set; } = "Report";
        public bool IsReportTimeSuffixPresent { get; set; } = true;
        public string ReportTimeSuffixFormat { get; set; } = "_dd_MM_yyyy";
        public string Extension { get; set; } = ReportExtension.Xlsx;

        public List<List<EdnaPointResult>> FetchReportData()
        {
            List<List<EdnaPointResult>> ednaPointsData = new List<List<EdnaPointResult>>();
            foreach (EdnaPoint pnt in EdnaPoints)
            {
                List<EdnaPointResult> historyResults = EdnaUtils.FetchHistoricalPointData(pnt, StartTime.GetTime(), EndTime.GetTime(), FetchStrategy, ResolutionSecs);
                ednaPointsData.Add(historyResults);
            }
            return ednaPointsData;
        }

        public string DeriveReportFileName()
        {
            string reportName = ReportName;
            string reportSuffix = "";
            if (IsReportTimeSuffixPresent)
            {
                try
                {
                    // derive report suffix from Start Time
                    reportSuffix = StartTime.GetTime().ToString(ReportTimeSuffixFormat);
                }
                catch (Exception e)
                {
                    reportSuffix = "";
                    Console.WriteLine(e.Message);
                }
            }
            return reportName + reportSuffix + "." + Extension;
        }

        public void CreateFolderIfRequired(string reportFolderPath)
        {
            // check if dest folder exists and create if abasent
            if (!Directory.Exists(reportFolderPath))
            {
                Directory.CreateDirectory(reportFolderPath);
                Console.WriteLine("Created the destination folder " + reportFolderPath);
            }
        }

        public void DeleteExistingFile(string reportFileNameWithPath)
        {
            // check if required file exists and delete if already exists
            if (File.Exists(reportFileNameWithPath))
            {
                File.Delete(reportFileNameWithPath);
                Console.WriteLine("Deleted existing file " + reportFileNameWithPath);
            }
        }

        public void GenerateReport()
        {
            // todo use async await
            List<List<EdnaPointResult>> pntsData = FetchReportData();
            //check if we have data
            if (pntsData.Count > 0)
            {
                // create list of timestamps
                List<DateTime> timeStamps = new List<DateTime>();

                // create list of headers
                List<string> headers = new List<string>();
                foreach (EdnaPoint pnt in EdnaPoints)
                {
                    headers.Add(pnt.Name);
                }

                // create report folder if required
                CreateFolderIfRequired(reportFolderPath);

                // derive the report name
                string reportFileName = DeriveReportFileName();
                string reportFileNameWithPath = Path.Combine(reportFolderPath, reportFileName);

                // deleting existing report file if present
                DeleteExistingFile(reportFileNameWithPath);

                // create report file
                ExcelPackage excel = new ExcelPackage(new FileInfo(reportFileNameWithPath));
                List<string> sheetNames = excel.Workbook.Worksheets.Select(ws => ws.Name).ToList();

                // create sheet if not present
                ExcelWorksheet sheet;
                if (sheetNames.Contains("Sheet1"))
                { sheet = excel.Workbook.Worksheets["Sheet1"]; }
                else
                { sheet = excel.Workbook.Worksheets.Add("Sheet1"); }

                int currentRow = 1;
                currentRow += NumRowsToSkip;

                // Dump Headers in excel
                sheet.SetValue(currentRow, 1, TimeStampHeading);
                for (int headerIter = 0; headerIter < headers.Count; headerIter++)
                { sheet.SetValue(currentRow, 2 + headerIter, headers[headerIter]); }
                currentRow += 1;

                // dump data from results
                if (pntsData[0] is List<EdnaPointResult> && pntsData[0].Count > 0)
                {
                    // get all the timestamps
                    for (int timeIter = 0; timeIter < pntsData[0].Count; timeIter++)
                    {
                        // dump the timestamp
                        string timeStamp = pntsData[0][timeIter].ResultTime_.ToString(TimeStampFormat);
                        sheet.SetValue(currentRow, 1, timeStamp);

                        // dump the points values
                        for (int pntIter = 0; pntIter < pntsData.Count; pntIter++)
                        {
                            sheet.SetValue(currentRow, 2 + pntIter, pntsData[pntIter][timeIter].Val_);
                        }
                        currentRow += 1;
                    }
                }

                // save the report file
                if (Extension == ReportExtension.Xlsx)
                {
                    excel.Save();
                }
                else if (Extension == ReportExtension.Csv)
                {
                    EpplusCsvConverter.ConvertToCsv(excel, reportFileNameWithPath);
                }
            }
        }
    }
}
