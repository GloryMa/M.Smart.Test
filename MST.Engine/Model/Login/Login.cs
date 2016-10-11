using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using MST.Engine.Model.Initialization;
using MST.Engine.Utility;

namespace MST.Engine.Model.Login
{
    class Login
    {
        [FindsBy(How = How.Id, Using = "user")]
        //[CacheLookup]
        private IWebElement UserName { get; set; }

        [FindsBy(How = How.Id, Using = "pwd")]
        //[CacheLookup]
        private IWebElement Password { get; set; }

        [FindsBy(How = How.Id, Using = "submit")]
        //[CacheLookup]
        private IWebElement Submit { get; set; }

        public void LoginFMWeb(string usename, string password)
        {
            PageFactory.InitElements(Initialize.Driver, this);
            UserName.EnterText(usename, "UserName");
            Password.EnterText(password, "Password");
            Submit.ClickOnIt("Submit Button");
            if (WebDriverExtensions.isAlertPresent(Initialize.Driver))
            {
                Initialize.Driver.SwitchTo().Alert().Accept();
            }
        }        
    }
}
