using NLog;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using System;
using System.Collections.Generic;

namespace MST.Engine.Model.Initialization
{
    class Initialize
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        private static readonly IDictionary<string, IWebDriver> Drivers = new Dictionary<string, IWebDriver>();
        private static IWebDriver driver;

        public static IWebDriver Driver
        {
            get
            {
                if (driver == null)
                    throw new NullReferenceException("The WebDriver browser instance was not initialized. You should first call the method InitBrowser.");
                return driver;
            }
            private set
            {
                driver = value;
            }
        }


        public void InitBrowser(string browserName)
        {
            logger.Info("Start to Initalize Browser");
            switch (browserName)
            {               
                case "Firefox":
                    logger.Debug("Firefox");
                    if (driver == null)
                    {
                        driver = new FirefoxDriver();
                        Drivers.Add("Firefox", Driver);
                    }
                    break;

                case "IE":
                    if (driver == null)
                    {
                        driver = new InternetExplorerDriver(@"\Driver");
                        Drivers.Add("IE", Driver);
                    }
                    break;

                case "Chrome":
                    if (driver == null)
                    {
                        logger.Debug("Chrome");
                        string path = System.IO.Directory.GetCurrentDirectory();
                        logger.Debug(path);
                        driver = new ChromeDriver(path+@"\Driver");                       
                        Drivers.Add("Chrome", Driver);
                    }
                    break;
            }
        }

        public void NavigateToHome()
        {
            //Driver.Manage().Cookies.DeleteAllCookies();
            string url = InitializeParamter.FMIndexUrl;
            logger.Info("Navigate To Home");
            Driver.Navigate().GoToUrl(url);
            Driver.Manage().Window.Maximize();
            //Driver.Manage().Window.Size = new Size(1440, 900);
        }

        public void CloseAllDrivers()
        {
            logger.Info("Close All Drivers");
            foreach (var key in Drivers.Keys)
            {
                Drivers[key].Close();
                Drivers[key].Quit();
            }
        }
    }
}
