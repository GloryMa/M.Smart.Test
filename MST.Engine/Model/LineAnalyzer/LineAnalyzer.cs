using MST.Engine.Model.Initialization;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System.Threading;
using MST.Engine.Utility;
using System.Diagnostics;
using System;
using NLog;

namespace MST.Engine.Model.LineAnalyzer
{
    class LineAnalyzer
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        [FindsBy(How = How.Id, Using = "btnExportData")]
        //[CacheLookup]
        private IWebElement ExportButton { get; set; }

        [FindsBy(How = How.ClassName, Using = "spinner")]
        //[CacheLookup]
        private IWebElement Spinner { get; set; }

        [FindsBy(How = How.ClassName, Using = "messageContainer")]
        //[CacheLookup]
        private IWebElement MessageContainer { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/div[3]/div[2]/form/div[3]/input")]
        //[CacheLookup]
        private IWebElement dialogButton { get; set; }


        public void Wait(string second)
        {
            Thread.Sleep(1000 * int.Parse(second));
        }
        public void ExportCsv()
        {
            CheckSpinner();            
            //Console.Read();
            Thread.Sleep(3000);
            //Error Alart Or Elapsed More Than 5 Minutes===> Return             
            if(MessageContainer.IsDisplayed("Message Container"))
            {
                logger.Fatal("Crash Issue On Loading Result");
                BusinessFlows.ErrorMessage = "Crash Issue On Loading Result";
                //dialogButton.ClickOnIt("Dialog Button");
                return;
            }
            else if(BusinessFlows.ErrorMessage == "Result Could Not Load After 5 Minutes")
            {
                return;
            }                         
            else
            {
                PageFactory.InitElements(Initialize.Driver, this);
                ExportButton.ClickOnIt("Export Button");
                Thread.Sleep(5000);
            }      
        }  

        
        
        public void CheckSpinner()
        {
            Thread.Sleep(5000);
            PageFactory.InitElements(Initialize.Driver, this);
            Stopwatch watcher = new Stopwatch();
            while (Spinner.IsDisplayed("Spinner"))
            {
                Thread.Sleep(5000);
                watcher.Start();
                if(!Spinner.IsDisplayed("Spinner")||watcher.Elapsed.TotalMinutes>5)
                {
                    if(watcher.Elapsed.TotalMinutes > 5)
                    {
                        BusinessFlows.ErrorMessage = "Result Could Not Load After 5 Minutes";
                        logger.Warn("Result Could Not Load After 5 Minutes");
                    }
                    watcher.Stop();
                    break;
                }
            }
        }           
    }
}
