using System;
using System.Linq;
using MST.Engine.Utility;
using System.Data;
using System.IO;
using OfficeOpenXml;
using System.Diagnostics;
using OpenQA.Selenium.Support.PageObjects;
using MST.Engine.Model.Initialization;
using OpenQA.Selenium;
using NLog;

namespace MST.Engine.Model.LineAnalyzer
{
    class CheckPoint
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        [FindsBy(How = How.Id, Using = "btnGoHome")]
        //[CacheLookup]
        private IWebElement GoDashboradButton { get; set; }

        [FindsBy(How = How.ClassName, Using = "messageContainer")]
        //[CacheLookup]
        private IWebElement MessageContainer { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/div[3]/div[2]/form/div[3]/input")]
        //[CacheLookup]
        private IWebElement dialogButton { get; set; }

        private static string status;
        private static string errorMessage;
        private static string EndTime;
        private static Stopwatch watch;
        private static string ElapsedTime;
        public void CheckResult()
        {
            try
            {                
                EndTime = DateTime.Now.ToString("yyyyMMddHHmmss");
                watch = BusinessFlows.watch;
                watch.Stop();
                ElapsedTime = watch.Elapsed.TotalMinutes.ToString("0.00");
                status = ScreenShotUtility.CompareScreen();
                errorMessage = BusinessFlows.ErrorMessage;
                if (status != "Failure")
                {
                    //status = CompareCsvFile();
                }
                SaveStatusToReport();
            }
            catch(Exception e)
            {
                logger.Fatal("Fail To Check Result.");
                logger.Fatal(e.Message);
            }
            finally
            {
                GoDashboradPage();
            }
        }

        public void GoDashboradPage()
        {
            PageFactory.InitElements(Initialize.Driver, this);
            if (MessageContainer.IsDisplayed("Message Container"))
            {                             
                dialogButton.ClickOnIt("Dialog Button");               
            }
            PageFactory.InitElements(Initialize.Driver, this);
            GoDashboradButton.ClickOnIt("Go Dashborad Button");
        }

        public void SaveStatusToReport()
        {
            try
            {
                logger.Debug("------|Start To Save Status To Report File");
                string SuiteName = Path.GetFileName(InitializeParamter.DataDriverFile);
                string CaseName = BusinessFlows.CaseName;
                string StartTime = BusinessFlows.StartTime;

                DataTable report = new DataTable();
                report = DataTableUtility.GetDataTable();
                report.Rows.Add(SuiteName, CaseName, status,
                    StartTime, EndTime, ElapsedTime, errorMessage, ScreenShotUtility.finalImageFile);
                FileInfo newFile = new FileInfo(InitializeParamter.ReportFile);
                using (ExcelPackage pck = new ExcelPackage(newFile))
                {
                    var sheetQuery = from sheet in pck.Workbook.Worksheets
                                     where sheet.Name == InitializeParamter.RunDay
                                     select sheet;
                    if (sheetQuery.Count() == 0)
                    {
                        ExcelWorksheet ws = pck.Workbook.Worksheets.Add(InitializeParamter.RunDay);
                        ws.Cells["A1"].LoadFromDataTable(report, true);
                        pck.Save();
                    }
                    else
                    {
                        sheetQuery.First().Cells["A1"].
                            LoadFromDataTable(report, true);
                        pck.Save();
                    }
                }
            }
            catch(Exception e)
            {
                logger.Fatal("Fail To Save Status To Report");
                logger.Fatal(e.Message);
            }
        }



        //public static string CompareCsv()
        //{
        //string csvPath = Path.GetDirectoryName(RunTest.ReportFile) + @"\Csv\";
        ////string baselineFile = csvPath + @"\Baseline\0_Multi-chronology1.csv";
        ////string actualFile= csvPath + @"\Actual\0_Multi-chronology1.csv";
        //string downloadPath = @"C:\Users\ghma\Downloads\";
        //var directory = new DirectoryInfo(downloadPath);
        //var query = from f in directory.GetFiles()
        //            where f.Name.StartsWith(TemplatesPage.templateName)
        //            orderby f.LastWriteTime descending
        //            select f;
        //if (query.Count() > 0)
        //{
        //    string actualFile = query.First().FullName;
        //    string baselineFile = csvPath + "Baseline\\" + TemplatesPage.templateName + ".csv";
        //    List<string> B_lines = File.ReadAllLines(baselineFile).ToList();
        //    List<string> A_lines = File.ReadAllLines(actualFile).ToList();
        //    //A有B没有
        //    var INA = A_lines.Except(B_lines);
        //    //B有A没有
        //    var INB = B_lines.Except(A_lines);
        //    if (INA.Count() > 0)
        //    {
        //        return "Failure";
        //    }
        //    else
        //    {
        //        return "Pass";
        //    }
        //}
        //else
        //{
        //    return "Failure";
        //}


        //}
    }
}
