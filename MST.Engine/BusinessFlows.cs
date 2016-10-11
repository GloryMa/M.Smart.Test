using EPPlus.Extensions;
using MST.Engine.Model.Initialization;
using MST.Engine.Model.LineAnalyzer;
using MST.Engine.Utility;
using NLog;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;



namespace MST.Engine
{
    public class BusinessFlows
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        public static string CaseName;
        private static TreeNode<string> SuiteTree;
        public static string StartTime;
        public static string EndTime;
        public static double ElapsedTime;
        public static string ErrorMessage;
        public static Stopwatch watch;
        public static string BuildVersion;

        public static void Run(string buildVersion)
        {            
            //0. Set Post Parameter
            BuildVersion = buildVersion;
            //1. Initialize Parameter via ENV.xlsx           
            InitializeParamter.InitializeParameter();
            //2. Set Up Environment 

            //3. Collect Suite Information via Data Driver File
            GenerateSuiteTree();
            //4. Execute Testing Based On Suite               
            ExecuteSuite();
            //5. Collect And Analysis Result 
            //6. Generate Test Report
            //7. Send Email  
            SendMail();
        }

        private static void SendMail()
        {
            try
            {
                logger.Info("=========Start To Send Mail...==============");
                string htmlBody = InitializeParamter.htmlBody;
                string mailAttach = InitializeParamter.mailAttach;
                mailAttach = InitializeParamter.ReportFile;
                DateTime runDate = DataTimeConvert(InitializeParamter.RunDay, "yyyyMMdd");
                var package = new ExcelPackage(new FileInfo(InitializeParamter.ReportFile));
                DataSet ds = Extensions.ToDataSet(package, true);
                DataTable reportTable = ds.Tables[InitializeParamter.RunDay];
                IEnumerable<DataRow> queryViaDate =
                    from rundate in reportTable.AsEnumerable()
                    where DataTimeConvert(rundate.Field<string>("Start Time"), "yyyyMMddHHmmss").Date == runDate.Date
                    select rundate;
                // Create a table from the query.
                if (queryViaDate.Count() > 0)
                {
                    DataTable DailyRecord = queryViaDate.CopyToDataTable<DataRow>();
                    htmlBody = ConvertDataTableToHTML(DailyRecord);
                }
                string strbody = ReplaceText(htmlBody);
                List<object> mailInfoList = new List<object>();
                mailInfoList.Add(InitializeParamter.mailTo);
                mailInfoList.Add(InitializeParamter.mailCc);
                mailInfoList.Add(InitializeParamter.mailSubject + ": " + BuildVersion);
                mailInfoList.Add(strbody);
                mailInfoList.Add(InitializeParamter.mailFrom);
                mailInfoList.Add(InitializeParamter.mailFromDisplay);
                mailInfoList.Add(mailAttach);
                mailInfoList.Add(InitializeParamter.mailSmtpClient);
                MST.Email.NetMail.SendMail(mailInfoList);
                logger.Info("---------------End-------------------------");
            }
            catch(Exception e)
            {
                logger.Fatal("Send Mail Fail...");
                logger.Fatal(e.Message);
                Environment.Exit(-1);
            }
        }

        public static string ConvertDataTableToHTML(DataTable dt)
        {
            string html = "<table id=\"customers\">";
            //add header row
            html += "<tr>";
            for (int i = 0; i < dt.Columns.Count; i++)
                html += "<th>" + dt.Columns[i].ColumnName + "</th>";
            html += "</tr>";
            //add rows
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                html += "<tr>";
                for (int j = 0; j < dt.Columns.Count; j++)
                    html += "<td>" + dt.Rows[i][j].ToString() + "</td>";
                html += "</tr>";
            }
            html += "</table>";
            return html;
        }

        public static string ReplaceText(string body)
        {
            string path = string.Empty;
            path = ("index.html");
            if (path == string.Empty)
            {
                return string.Empty;
            }
            StreamReader sr = new StreamReader(path);
            string str = string.Empty;
            str = sr.ReadToEnd();
            str = str.Replace("$Body$", body);
            return str;
        }

        private static DateTime DataTimeConvert(string dateString,string dateFormat)
        {
            DateTime dt = DateTime.ParseExact(dateString,
                dateFormat, null, System.Globalization.DateTimeStyles.AllowWhiteSpaces);
            return dt;
        }

        private static void GenerateSuiteTree()
        {
            try
            {
                logger.Info("======Start To Generate Suite Tree...======");
                Dictionary<string, string> dic = new Dictionary<string, string>();
                var package = new ExcelPackage(new FileInfo(InitializeParamter.DataDriverFile));
                DataSet ds = Extensions.ToDataSet(package, true);
                DataTable table = ds.Tables[0];

                /*---------------Tree Structure----------
                //|Suites|
                //      --|Case 1|
                //               --|Model 1|
                //                        --|Action 1|
                //                                  --|Parameter 1|
                ---------------------------------------*/
                string root = "Suites";//|Suites|           
                SuiteTree = new TreeNode<string>(root);

                for (int r = 0; r < table.Rows.Count; r++)
                {
                    string isCheck = table.Rows[r][0].ToString();
                    if (isCheck == "YES")
                    {
                        string description = table.Rows[r][2].ToString();
                        var CaseNode = SuiteTree.AddChild(description);//--|Case 1|
                        string model = table.Rows[r][1].ToString();
                        var modelNode = CaseNode.AddChild(model);//--|Model 1|
                        for (int c = 3; c < table.Columns.Count; c++)
                        {
                            {
                                string setting = table.Rows[r][c].ToString();
                                if ((setting != null && setting != @"" && setting != "\"\""))
                                {
                                    string[] actionsInfo = setting.Split(';');
                                    foreach (var item in actionsInfo)
                                    {
                                        string[] parameters = item.Split(',');
                                        if (parameters.Count() > 1)
                                        {
                                            string action = parameters[0];
                                            var actionNode = modelNode.AddChild(action);//--|Action 1|
                                            for (int i = 1; i < parameters.Count(); i++)
                                            {
                                                var parameterNode = actionNode.AddChild(parameters[i]); //--| Parameter 1 |
                                            }
                                        }
                                        else
                                        {
                                            string action = item;
                                            var actionNode = modelNode.AddChild(action);//--|Action 1|
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                logger.Info("---------------End-------------------------");
            }
            catch (Exception e)
            {
                logger.Fatal("Generate Suite Tree Fail.");
                logger.Fatal(e.Message);
                Environment.Exit(-1);
            }
        }

        private static void ExecuteSuite()
        {
            logger.Info("=========Start To Execute Suite...=========");
            logger.Info("Total Case: " + SuiteTree.Count());
            int itemIndex = 0;
            foreach (var item in SuiteTree)
            {
                ++itemIndex;               
                ErrorMessage = "";
                watch = new Stopwatch();
                watch.Start();
                StartTime = DateTime.Now.ToString("yyyyMMddHHmmss");
                string caseName = item.Data;
                CaseName = caseName;
                string modelName = item.Children.First().Data;
                logger.Info("--["+itemIndex+@"/"+ SuiteTree.Count() +"]"+"Start Case: " + CaseName);

                foreach (var action in item.Children.First().Children)
                {                    
                    try
                    {                        
                        string actionName = action.Data;
                        logger.Info("----|Start Action: " + actionName);
                        object[] o = new object[] { };
                        o = action.Children.Select(x => x.Data as object).ToArray();

                        var pages = Assembly.GetExecutingAssembly().GetTypes()
                              .Where(t => t.Namespace == "MST.Engine.Model." + modelName)
                              .ToList();

                        var actions = from p in pages
                                      from a in Type.GetType(p.FullName).
                              GetMethods(BindingFlags.Public | BindingFlags.Instance)
                                      where a.Name == actionName
                                      select new { a,p} ;                      

                        if (actions.Count() > 0)
                        {
                            Type t = Type.GetType(actions.First().p.FullName);
                            var instance = Activator.CreateInstance(t);
                            t.InvokeMember(actions.First().a.Name, BindingFlags.Public | BindingFlags.Instance
                                | BindingFlags.InvokeMethod, null, instance, o);                            
                        }
                        else
                        {
                            logger.Error("Could Not Found Action: " + actionName +
                                ", Please Check DataDriver File Or Add The Action On Code");
                            break;
                        }                       
                    }
                    catch (Exception e)
                    {
                        EndTime = DateTime.Now.ToString("yyyyMMddHHmmss");
                        watch.Stop();
                        ElapsedTime = watch.Elapsed.TotalMinutes;
                        CheckPoint check = new CheckPoint();
                        ErrorMessage = e.Message;
                        check.CheckResult();
                        break;
                    }
                }
                watch.Stop();
                EndTime = DateTime.Now.ToString("yyyyMMddHHmmss");
                ElapsedTime = watch.Elapsed.TotalMinutes;
            }
            logger.Info("---------------End-------------------------");
        }
    }

}
