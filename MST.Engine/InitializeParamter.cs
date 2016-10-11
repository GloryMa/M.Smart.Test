using EPPlus.Extensions;
using NLog;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MST.Engine
{
    class InitializeParamter
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        public static string DataDriverFile { get; set; }
        public static string ReportFile { get; set; }
        public static string FMIndexUrl { get; set; }
        public static string ImageBaselinePath { get; set; }
        public static string RunDay { get; set; }
        public static string mailTo { get; set; }
        public static string mailCc { get; set; }
        public static string mailSubject { get; set; }
        public static string htmlBody { get; set; }
        public static string mailFrom { get; set; }
        public static string mailFromDisplay { get; set; }
        public static string mailAttach { get; set; }
        public static string mailSmtpClient { get; set; }       

        public static void InitializeParameter()
        {
            try
            {
                logger.Info("======Start To Initialize Parameters...====");
                var package = new ExcelPackage(new FileInfo(@"Config\ENV.xlsx"));
                DataSet ds = Extensions.ToDataSet(package, true);
                //Env Sheet
                DataTable envTable = ds.Tables["Env"];
                var envDict = envTable.AsEnumerable().ToDictionary(
                    r => r.Field<string>("Key"),
                    r => r.Field<string>("Value"));
                DataDriverFile = envDict["DataDriverFile"];
                ReportFile = envDict["ReportFile"];
                FMIndexUrl = envDict["FMIndexUrl"];
                ImageBaselinePath = envDict["ImageBaselinePath"];
                //RunDay = envDict["RunDay"];

                RunDay = DateTime.Today.ToString("yyyyMMdd");

                logger.Debug("Run Day is " + RunDay);

                //Mail Sheet
                DataTable mailTable = ds.Tables["Mail"];
                var Maildict = mailTable.AsEnumerable().ToDictionary(
                    r => r.Field<string>("Key"),
                    r => r.Field<string>("Value"));
                mailTo = Maildict["mailTo"];
                mailCc = Maildict["mailCc"];
                mailSubject = Maildict["mailSubject"];
                htmlBody = Maildict["htmlBody"];
                mailFrom = Maildict["mailFrom"];
                mailFromDisplay = Maildict["mailFromDisplay"];
                mailAttach = Maildict["mailAttach"];
                mailSmtpClient = Maildict["mailSmtpClient"];                
                logger.Info("---------------End-------------------------");
            }
            catch(Exception e)
            {
                logger.Fatal("Initialize Paramter Fail.");
                logger.Fatal(e.Message);
                Environment.Exit(-1);
            }
        }
    }

}
