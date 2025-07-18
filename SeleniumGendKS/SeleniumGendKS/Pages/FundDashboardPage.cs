using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using SeleniumGendKS.Core.BaseClass;
using SeleniumGendKS.Core.Selenium;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Xml.XPath;

namespace SeleniumGendKS.Pages
{
    internal class FundDashboardPage : BasePageElementMap
    {
        // Initiate variables
        internal static readonly XDocument xdoc = LoginPage.xdoc;
        internal static string managerNameAlbourne = xdoc.XPathSelectElement("config/fundAlbourne").Attribute("managerName").Value;
        internal static string fundNameAlbourne = xdoc.XPathSelectElement("config/fundAlbourne").Attribute("fundName").Value;
        internal static string managerNameManual = xdoc.XPathSelectElement("config/fundManual").Attribute("managerName").Value;
        internal static string fundNameManual = xdoc.XPathSelectElement("config/fundManual").Attribute("fundName").Value;
        internal static WebDriverWait? wait;
        internal static string exposure = "Exposure";
        internal static string fundReturns = "Fund Returns";
        internal static string fundAum = "Fund AUM";
        internal static string rowCounts = "Row Counts";
        internal static string startDate = "Start Date";
        internal static string endDate = "End Date";
        internal static string dealInformation = "Deal Information";
        internal static string fundInformation = "Fund Information";
        internal static string cashFlowInformation = "Cash Flow Information";
        internal static string rowCount = "Row Count";
        internal static string asOfDate = "As of Date";
        internal static string yearFirmFounded = "Year Firm Founded";
        internal static string strategyHeadquarter = "Strategy Headquarter";
        internal static string businessContact	 = "Business Contact";
        internal static string contactPhone = "Contact Phone";
        internal static string contactEmail = "Contact Email";
        internal static string firm = "Firm";
        internal static string strategy = "Strategy";
        internal static string vintageYear = "Vintage Year";
        internal static string assetClass = "Asset Class";
        internal static string investmentStage = "Investment Stage";
        internal static string industryFocus = "Industry Focus";
        internal static string geographicFocus = "Geographic Focus";
        internal static string fundSizeM = "Fund Size (M)";
        internal static string currency = "Currency";

        // Initiate the By objects for elements
        internal static By searchFundInputTxt = By.XPath(@"//input[@placeholder='Search']");
        internal static By loadingIcon = By.XPath(@"//div[contains(@class,'progress-spinner')]");
        internal static By fundNameTitle = By.XPath(@"//div[@class='title']");
        internal static By fundInfoTab = By.Id("p-tabpanel-0-label");
        internal static By managerInfoTab = By.XPath(@"//span[contains(.,'Manager Info')]");
        internal static By fundListTab = By.XPath(@"//span[contains(.,'Fund List')]");
        internal static By dataStatusTab = By.XPath(@"//span[contains(.,'Data Status​')]");
        internal static By addNewFundIcon = By.XPath(@"//button[contains(@class,'add-fund')]");
        internal static By editFundPencilIcon = By.XPath(@"//button[contains(@icon,'pencil')]");
        internal static By expandAllBtn = By.XPath(@"//button[@label='Expand All']");
        internal static By collapseAllBtn = By.XPath(@"//button[@label='Collapse All']");
        internal static By fundNameReturnOfResults(string fundName) => By.XPath(@"//div[.='" + fundName + "']");
        internal static By fundNameReturnOfResultsWithItemSource(string fundName, string sourceIcon) => By.XPath(@"//div[.='" + fundName + "']/preceding-sibling::span[.='" + sourceIcon + "']");
        internal static By fundInfoLabel(string fundInfo) => By.XPath(@"//td[contains(., '" + fundInfo + "')]");
        internal static By fundInfoTxt(string fundInfo) => By.XPath(@"//td[contains(., '" + fundInfo + "')]//following-sibling::td");
        internal static By sectionTitleContent(string title, string content) => By.XPath(@"//div[@class='section-title' and contains(.,'" + title + "')]//following-sibling::div/div//div[contains(.,'" + content + "')]//following-sibling::div");
        internal static By sectionTitleContentNoData(string title) => By.XPath(@"//div[@class='section-title' and contains(.,'" + title + "')]//following-sibling::div/div");
        internal static By listStrategiesTxt(int number) => By.XPath(@"//div[@class='funds-container']/div[1+" + number + "]//div[@class='section']/div[1]");
        internal static By listFundsNameTxt(int groupNumber, int number) => By.XPath(@"//div[@class='funds-container']/div[1+" + groupNumber + "]/div[@class='section']/div[2]/div[" + number + "]//div[@class='section-title']");
        internal static By listFundsInfoTxt(int groupNumber, int number, string fieldLabel) => By.XPath(@"//div[@class='funds-container']/div[1+" + groupNumber + "]/div[@class='section']/div[2]/div[" + number + "]//table//td[.='" + fieldLabel + "']/following-sibling::td");
        internal static By toastMessage(string text) => By.XPath(@"//p-toastitem[contains(@class,'toastAnimation')]//div[.='" + text + "']");

        // Initiate the elements
        public IWebElement inputTxtSearchFund(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(d=>d.FindElement(searchFundInputTxt));
        }
        public IWebElement returnOfResultsFundName(int timeoutInSeconds, string fundName)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(fundNameReturnOfResults(fundName)));
        }
        public IWebElement returnOfResultsFundNameWithItemSource(int timeoutInSeconds, string fundName, string sourceIcon)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(fundNameReturnOfResultsWithItemSource(fundName, sourceIcon)));
        }
        public IWebElement titleFundName(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(fundNameTitle));
        }
        public IWebElement tabFundInfo(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(fundInfoTab));
        }
        public IWebElement tabManagerInfo(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(managerInfoTab));
        }
        public IWebElement tabFundList(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(fundListTab));
        }
        public IWebElement tabDataStatus(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(dataStatusTab));
        }
        public IWebElement iconAddNewFund(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(addNewFundIcon));
        }
        public IWebElement iconPencilEditFund(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(editFundPencilIcon));
        }
        public IWebElement buttonExpandAll(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(expandAllBtn));
        }
        public IWebElement buttonCollapseAll(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(collapseAllBtn));
        }
        public IWebElement labelFundInfo(int timeoutInSeconds, string label)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(fundInfoLabel(label)));
        }
        public IWebElement textFundInfo(int timeoutInSeconds, string fundInfo)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(fundInfoTxt(fundInfo)));
        }
        public IWebElement textSectionTitleContent(int timeoutInSeconds, string title, string content)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(sectionTitleContent(title, content)));
        }
        public IWebElement textSectionTitleContentNoData(int timeoutInSeconds, string title)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(sectionTitleContentNoData(title)));
        }
        public IWebElement textListStrategies(int timeoutInSeconds, int number)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(listStrategiesTxt(number)));
        }
        public IWebElement textListFundsName(int timeoutInSeconds, int groupNumber, int number)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(listFundsNameTxt(groupNumber, number)));
        }
        public IWebElement textListFundsInfo(int timeoutInSeconds, int groupNumber, int number, string fieldLabel)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(listFundsInfoTxt(groupNumber, number, fieldLabel)));
        }    
        public IWebElement toastMessageAlert(int timeoutInSeconds, string text)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(toastMessage(text)));
        }
    }

    internal sealed class FundDashboardAction : BasePage<FundDashboardAction, FundDashboardPage>
    {
        #region FundDashboardAction constructor
        private FundDashboardAction() {}
        #endregion

        #region FundDashboard Items action
        // Wait for loading Spinner icon to disappear
        public FundDashboardAction WaitForLoadingIconToDisappear(int timeoutInSeconds, By element)
        {
            if (timeoutInSeconds > 0)
            {
                WebDriverWait wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
                wait.Until(ExpectedConditions.ElementIsVisible(element));
                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(element));
            }
            return this;
        }

        // Wait for element visible
        public FundDashboardAction WaitForElementVisible(int timeoutInSeconds, By element)
        {
            if (timeoutInSeconds > 0)
            {
                WebDriverWait wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
                wait.Until(ExpectedConditions.ElementIsVisible(element));
            }
            return this;
        }

        // Checking element exists or not
        public bool IsElementPresent(By by)
        {
            try
            {
                Driver.Browser.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }

        // (wait method 2) Wait for element load Done (can apply for performance)
        public FundDashboardAction WaitingElementLoadDone(int timeout, By by)
        {
            int times = 0;
            while ((!IsElementPresent(by)) && (timeout < times))
            {
                times++;
                System.Threading.Thread.Sleep(1000);
            }

            if (times == timeout)
            {
                Console.WriteLine("The Element failed to load in " + timeout + " seconds");
                return this;
            }
            else
            {
                //Console.WriteLine("The Element Load Done in " + times + " seconds");
                return this;
            }
        }

        public FundDashboardAction InputNameToSearchFund(int timeoutInSeconds, string fundName)
        {
            this.Map.inputTxtSearchFund(timeoutInSeconds).Clear();
            this.Map.inputTxtSearchFund(timeoutInSeconds).SendKeys(fundName);
            return this;
        }

        public FundDashboardAction ClickFundNameReturnOfResults(int timeout, string fundName)
        {
            Actions actions = new Actions(Driver.Browser);
            actions.MoveToElement(this.Map.returnOfResultsFundName(timeout, fundName));
            actions.Perform();
            this.Map.returnOfResultsFundName(timeout, fundName).Click();
            return this;
        }

        public FundDashboardAction ClickFundNameReturnOfResults(int timeout, string fundName, string sourceIcon)
        {
            Actions actions = new Actions(Driver.Browser);
            actions.MoveToElement(this.Map.returnOfResultsFundNameWithItemSource(timeout, fundName, sourceIcon));
            actions.Perform();
            this.Map.returnOfResultsFundNameWithItemSource(timeout, fundName, sourceIcon).Click();
            return this;
        }

        public string FundNameTitleGetText(int timeout)
        {
            return this.Map.titleFundName(timeout).Text;
        }

        public string FundInfoGetLabel(int timeout, string label)
        {
            return this.Map.labelFundInfo(timeout, label).Text;
        }
        public string FundInfoGetText(int timeout, string fundInfo)
        {
            return this.Map.textFundInfo(timeout, fundInfo).Text;
        }

        public FundDashboardAction ClickFundInfoTab(int timeoutInSeconds)
        {
            this.Map.tabFundInfo(timeoutInSeconds).Click();
            return this;
        }

        public FundDashboardAction ClickManagerInfoTab(int timeoutInSeconds)
        {
            this.Map.tabManagerInfo(timeoutInSeconds).Click();
            return this;
        }

        public FundDashboardAction ClickFundListTab(int timeoutInSeconds)
        {
            this.Map.tabFundList(timeoutInSeconds).Click();
            return this;
        }

        public FundDashboardAction ClickDataStatusTab(int timeoutInSeconds)
        {
            //this.Map.tabDataStatus(timeoutInSeconds).Click(); // --> element click intercepted
            // Try with javascript
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].click();", this.Map.tabDataStatus(timeoutInSeconds));
            return this;
        }

        public FundDashboardAction ClickAddNewFundIcon(int timeoutInSeconds)
        {
            this.Map.iconAddNewFund(timeoutInSeconds).Click();
            return this;
        }

        public FundDashboardAction ClickEditFundPencilIcon(int timeoutInSeconds)
        {
            this.Map.iconPencilEditFund(timeoutInSeconds).Click();
            return this;
        }

        public FundDashboardAction ClickExpandAllButton(int timeoutInSeconds)
        {
            this.Map.buttonExpandAll(timeoutInSeconds).Click();
            return this;
        }

        public FundDashboardAction ClickCollapseAllButton(int timeoutInSeconds)
        {
            this.Map.buttonCollapseAll(timeoutInSeconds).Click();
            return this;
        }

        public string SectionTitleContentGetText(int timeout, string title, string content)
        {
            return this.Map.textSectionTitleContent(timeout, title, content).Text;
        }

        public string SectionTitleContentNoDataGetText(int timeout, string title)
        {
            return this.Map.textSectionTitleContentNoData(timeout, title).Text;
        }

        public string ListStrategiesGetText(int timeout, int number)
        {
            return this.Map.textListStrategies(timeout, number).Text;
        }

        public string ListFundsNameGetText(int timeout, int groupNumber, int number)
        {
            return this.Map.textListFundsName(timeout, groupNumber, number).Text;
        }

        public string ListFundsInfoGetText(int timeout, int groupNumber, int number, string fieldLabel)
        {
            return this.Map.textListFundsInfo(timeout, groupNumber, number, fieldLabel).Text;
        }

        public string toastMessageAlertGetText(int timeout, string text)
        {
            return this.Map.toastMessageAlert(timeout, text).Text;
        }
        #endregion

        #region FundDashboard Built-in actions
        public FundDashboardAction InputNameToSearchFund(int timeoutInSeconds, string searchInput, string managerName, string fundName, string sourceIcon)
        {
            InputNameToSearchFund(timeoutInSeconds, searchInput)
            .WaitForElementVisible(timeoutInSeconds, FundDashboardPage.fundNameReturnOfResultsWithItemSource(managerName, sourceIcon))
            .ClickFundNameReturnOfResults(timeoutInSeconds, managerName, sourceIcon)
            .WaitForElementVisible(timeoutInSeconds, FundDashboardPage.fundNameReturnOfResultsWithItemSource(fundName, sourceIcon))
            .ClickFundNameReturnOfResults(timeoutInSeconds, fundName, sourceIcon)
            .WaitForLoadingIconToDisappear(timeoutInSeconds, FundDashboardPage.loadingIcon);
            return this;
        }
        #endregion
    }
}
