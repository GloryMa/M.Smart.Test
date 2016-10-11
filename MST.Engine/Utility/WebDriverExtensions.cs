using NLog;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MST.Engine.Utility
{
    public static class WebDriverExtensions
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        public static void EnterText(this IWebElement element, string text, string elementName)
        {
            element.Clear();
            element.SendKeys(text);
            logger.Debug("------|"+text + " entered in the " + elementName + " field.");
        }

        public static bool IsDisplayed(this IWebElement element, string elementName)
        {
            bool result;
            try
            {
                result = element.Displayed;
                logger.Debug("------|" + elementName + " is Displayed.");
            }
            catch (Exception)
            {
                result = false;
                logger.Debug("------|" + elementName + " is not Displayed.");
            }

            return result;
        }

        public static void ClickOnIt(this IWebElement element, string elementName)
        {
            element.Click();
            logger.Debug("------|" + "Clicked on " + elementName);
        }

        public static void SelectByText(this IWebElement element, string text, string elementName)
        {
            SelectElement oSelect = new SelectElement(element);
            oSelect.SelectByText(text);
            logger.Debug("------|" + text + " text selected on " + elementName);
        }

        public static void SelectByIndex(this IWebElement element, int index, string elementName)
        {
            SelectElement oSelect = new SelectElement(element);
            oSelect.SelectByIndex(index);
            logger.Debug("------|" + index + " index selected on " + elementName);
        }

        public static void SelectByValue(this IWebElement element, string text, string elementName)
        {
            SelectElement oSelect = new SelectElement(element);
            oSelect.SelectByValue(text);
            logger.Debug("------|" + text + " value selected on " + elementName);
        }

        public static bool isAlertPresent(IWebDriver WebDriver)
        {
            try
            {
                Thread.Sleep(1000);
                WebDriver.SwitchTo().Alert();
                return true;
            }
            catch (NoAlertPresentException e)
            {
                logger.Trace("------|" + e.Message);
                return false;
            }
        }
    }
}
