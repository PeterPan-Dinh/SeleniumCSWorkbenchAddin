using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumGendKS.Core.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumGendKS.Core.BaseClass
{
    public class BasePageElementMap
    {
        protected IWebDriver? browser;
        protected WebDriverWait? browserWait;

        public BasePageElementMap()
        {
            this.browser = Driver.Browser;
            this.browserWait = Driver.BrowserWait;
        }

        public void SwicthToDefault()
        {
            this.browser.SwitchTo().DefaultContent();
        }
    }
}
