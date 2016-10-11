using MST.Engine.Model.Initialization;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using MST.Engine.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;

namespace MST.Engine.Model.LineAnalyzer
{
    class Dashborad
    {      
        [FindsBy(How = How.Id, Using = "btnLogout")]
        //[CacheLookup]
        private IWebElement LogoutButton { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='wrapper']/nav/div/ul/li[10]/a/img")]
        //[CacheLookup]
        private IWebElement GoTemplateButton { get; set; }
       
        public void GoTemplatePage()
        {
            CheckGoTemplateButton();
            GoTemplateButton.ClickOnIt("Go Template Button");
        }


        public void Logout()
        {
            PageFactory.InitElements(Initialize.Driver, this);
            LogoutButton.ClickOnIt("Logout Button");
            if (WebDriverExtensions.isAlertPresent(Initialize.Driver))
            {
                Initialize.Driver.SwitchTo().Alert().Accept();
            }
        }

        public void CheckGoTemplateButton()
        {            
            Thread.Sleep(3000);
            PageFactory.InitElements(Initialize.Driver, this);
            Stopwatch watcher = new Stopwatch();
            while (!GoTemplateButton.IsDisplayed("Go Template Button"))
            {
                Thread.Sleep(3000);
                watcher.Start();
                if (GoTemplateButton.IsDisplayed("Go Template Button") || watcher.Elapsed.TotalMinutes > 5)
                {
                    watcher.Stop();
                    break;
                }
            }
        }

    }
}
