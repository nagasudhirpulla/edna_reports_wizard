using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Reflection;
using OfficeOpenXml;
using System.IO;
using EdnaReportsWizard;
using EdnaDataLayer;
using VariableTime;

namespace TestScripts
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void TestExcelCreation_Click(object sender, RoutedEventArgs e)
        {
            // send excel file for user to download - https://www.ryadel.com/en/asp-net-generate-excel-files-programmatically-epplus-guide-tutorial-mvc-core/amp/

            ExcelPackage excel = new ExcelPackage(new FileInfo("C:\\Users\\Nagasudhir\\Downloads\\test_excel.xlsx"));
            List<string> sheetNames = excel.Workbook.Worksheets.Select(ws => ws.Name).ToList();

            // create sheet if not present
            ExcelWorksheet sheet;
            if (sheetNames.Contains("Sheet1"))
            {
                sheet = excel.Workbook.Worksheets["Sheet1"];
            }
            else
            {
                sheet = excel.Workbook.Worksheets.Add("Sheet1");
            }

            sheet.Cells["A1"].Value = 12;
            sheet.SetValue(2, 3, 501);
            sheet.SetValue(3, 4, "Sudhir");
            // output the XLSX file
            excel.Save();
            MessageBox.Show("Saved");
        }

        private void TestReportCreation_Click(object sender, RoutedEventArgs e)
        {
            // create a test report template
            ReportTemplate template = new ReportTemplate();

            // set the template variables
            template.EdnaPoints = new List<EdnaPoint>() {   new EdnaPoint() { Name="Point1", PointId="id1"},
                                                            new EdnaPoint() { Name="Point2", PointId="id2"},
                                                            new EdnaPoint() { Name="Point3", PointId="id3"}};
            DateTime absTime = DateTime.Now.AddHours(-1);
            template.StartTime = new VarTime() { AbsoluteTime = absTime };
            template.EndTime = new VarTime() { AbsoluteTime = absTime.AddHours(0.5) };
            template.reportFolderPath = @"C:\Users\Nagasudhir\Downloads";
            template.ReportName = "test template";
            template.IsReportTimeSuffixPresent = true;
            template.Extension = ReportExtension.Csv;
            template.ResolutionSecs = 60;

            // generate the report
            template.GenerateReport();
        }
    }
}
