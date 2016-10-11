using MST.Engine.Model.Initialization;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System.Linq;
using System.Threading;
using MST.Engine.Utility;

namespace MST.Engine.Model.LineAnalyzer
{
    class Template
    {
        public static string templateName;

        [FindsBy(How = How.Id, Using = "run")]
        //[CacheLookup]
        private IWebElement RunButton { get; set; }

        public void SelectTemplate(string item)
        {
            templateName = item;
            Thread.Sleep(1000);
            PageFactory.InitElements(Initialize.Driver, this);
            //string xpathString = ".//*[text() = '" + item + "']";
            string xpathString = ".//*[contains(text(), '" + item + "')]";
            var templates = Initialize.Driver.FindElements(By.XPath(xpathString));
            var template = from t in templates
                           where t.Text.Split('(')[0] == item
                           select t;
            template.First().ClickOnIt("Template Item Link");
        }
        public void Run()
        {
            Thread.Sleep(1000);
            PageFactory.InitElements(Initialize.Driver, this);
            RunButton.ClickOnIt("Run Button");
        }
    }
}
