using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumGendKS.Core.Selenium;
using SeleniumGendKS.Core.BaseClass;
using System;
using System.Xml.Linq;
using SeleniumExtras.WaitHelpers;
using OpenQA.Selenium.Interactions;
using System.Xml;

namespace SeleniumGendKS.Pages
{
    internal class NavigationPage : BasePageElementMap
    {
        internal static readonly XDocument xdoc = LoginPage.xdoc;
        internal static WebDriverWait? wait;
        //internal static string? frameNameInstacne;

        // Initiate the By objects for elements
        internal static By fundInfoTab = By.Id("p-tabpanel-0-label");
        internal static By dataStatusTab = By.Id("p-tabpanel-1-label");
        internal static By editFundPencilIcon = By.XPath(@"//span[contains(@class,'pencil')]");
        internal static By addNewFundButton = By.XPath(@"//button[contains(@class,'button-rounded add-fund-fixed')]");
        internal static By uploadPDFFileIcon = By.XPath(@"//i[@class='pi pi-upload']");
        internal static By uploadCSVFileButton = By.XPath(@"//button[contains(.,'Upload File')]");
        internal static By userInputSection = By.XPath(@"//span[.='User Input']");
        internal static By dateSectionInUserInput = By.XPath(@"//div[@class='section-title' and contains(.,'Date Section')]");
        internal static By crbmSectionInUserInput = By.XPath(@"//div[@class='section-title' and contains(.,'Custom Risk Benchmark Modelling')]");
        internal static By feeModelSessionInUserInput = By.XPath(@"//div[@class='section-title' and contains(.,'Fee Model Section')]");
        internal static By frameIdFADAddIn = By.XPath(@"//iframe[@title='Office Add-in FAD Add-in']");
        internal static By frameIdFADAddInCurrentInstance(string frameNameInstance) => By.XPath(@"//iframe[contains(@src,'workbench." + frameNameInstance + "') or contains(@src,'workbench-" + frameNameInstance + "')]");
        internal static By frameOfficeExcelOnline = By.XPath(@"//iframe[@id='WebApplicationFrame' or @id='WacFrame_Excel_0']"); //By.Id("WebApplicationFrame"); // By.CssSelector("iframe[id='WebApplicationFrame'")
        internal static By frameDialogAddinExcelOnline = By.Id("InsertDialog");
        internal static By dialogNotificationExcelOnline(string text) => By.XPath(@"//label[.='"+ text +"']");
        internal static By dontShowButtonInReviewMenuExcelOnline = By.XPath("//button[@aria-label=\"Don't show\"]");
        internal static By okInvalidSheetButtonExcelOnline = By.Id("DialogActionButton");
        internal static By gotItBtnExcelOnline = By.XPath(@"//button[@aria-label='Got it']");
        internal static By homeMenuExcelOnline = By.XPath(@"//span[contains(@class, 'ms-Button-label pivot-header-label label') and .='Home']");
        internal static By menuItemExcelOnline(string name) => By.XPath(@"//div[@role='menuitem' and contains(@name,'" + name + "')]");
        internal static By insertMenuExcelOnline = By.XPath(@"//span[contains(@class, 'ms-Button-label pivot-header-label label') and .='Insert']"); 
        internal static By fileMenuExcelOnline = By.XPath(@"//button[contains(@data-unique-id,'FileMenu')]");
        internal static By moreFileOptionButtonExcelOnline = By.XPath(@"//button[@data-unique-id='FileMenuOverflowButton']");
        internal static By saveAsButtonExcelOnline = By.XPath(@"//button[@data-unique-id='FileSaveAsPage']");
        internal static By downloadACopyButtonExcelOnline = By.XPath(@"//button[@data-unique-id='DownloadACopy']");
        internal static By officeAddInsButtonExcelOnline = By.XPath(@"//button[@id='InsertAppsForOffice']");
        internal static By addinsButtonExcelOnline = By.XPath(@"//button[@aria-label='Add-ins']");
        //internal static By popularAddinsPopup = By.XPath(@"//div[@id='WACDialogPanel']"); 
        internal static By popularAddinsPopup = By.XPath(@"//div[@aria-label='Add-ins' and @tabindex='0']"); // using this when ExcelOnline enable 'More Addins Button' 
        internal static By moreAddinsButtonExcelOnline = By.XPath(@"//button[@aria-label='More Add-ins']");
        internal static By advancedButtonExcelOnline = By.XPath("//button[@aria-label='Advanced...']");
        internal static By myADDINSTabExcelOnline = By.Id(@"MY ADD-INS");
        internal static By uploadMyAddInlinkExcelOnline = By.XPath(@"//span[@id='UploadMyAddin']");
        internal static By browserFileMyAddInExcelOnline = By.Id("BrowserFile");
        internal static By uploadButtonUploadAddInExcelOnline = By.XPath(@"//input[@id='DialogInstall']");
        internal static By fadAddinButtonExcelOnline = By.XPath(@"//button[@aria-label='FAD Add-in']");
        internal static By notificationCloseButtonExcelOnline = By.XPath(@"//button[contains(@id,'-closeButton')]");
        internal static By workBenchLeftPaneExcelOnline = By.XPath(@"//div[@class='ewa-ptp-nov2-flex' or @class='ewa-ptp-nov2-flex ewa-ptp-nov2']"); //By.XPath(@"//div[@class='ewa-ptp-nov2-flex ewa-ptp-nov2']");
        internal static By workBenchSheetName(string sheetName) => By.XPath(@"//ul[@class='ewa-stb-tabs']/li/a[@aria-label='" + sheetName + "']");
        internal static By deleteSheetAtRightClickExcelOnline = By.XPath(@"//button[@name='Delete']");

        // Initiate the elements
        public IWebElement tabFundInfo(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(fundInfoTab));
        }
        public IWebElement tabDataStatus(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(dataStatusTab));
        }
        public IWebElement pencilIconEditFund(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(editFundPencilIcon));
        }
        public IWebElement buttonAddNewFund(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(addNewFundButton));
        }
        public IWebElement iconUploadPDFFile(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(uploadPDFFileIcon));
        }
        public IWebElement buttonuploadCSVFile(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(uploadCSVFileButton));
        }
        public IWebElement sectionUserInput(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(userInputSection));
        }
        public IWebElement userInputDateSection(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(dateSectionInUserInput));
        }
        public IWebElement userInputCRBMSection(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(crbmSectionInUserInput));
        }
        public IWebElement userInputFeeModelSession(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(feeModelSessionInUserInput));
        }
        public IWebElement imageIconExcelOnlineInstance(int timeoutInSeconds, int instanceIcon)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(@"//*[@id='AddinControl" + instanceIcon.ToString() + "']/span/i/img")));
        }
        public IWebElement itemMenuExcelOnline(int timeoutInSeconds, string name)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(menuItemExcelOnline(name)));
        }
        public IWebElement menuHomeExcelOnline(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(homeMenuExcelOnline));
        }
        public IWebElement menuInsertExcelOnline(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(insertMenuExcelOnline));
        }
        public IWebElement menuFileExcelOnline(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(fileMenuExcelOnline));
        }
        public IWebElement buttonMoreFileOptionExcelOnline(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(moreFileOptionButtonExcelOnline));
        }
        public IWebElement buttonSaveAsExcelOnline(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(saveAsButtonExcelOnline));
        }
        public IWebElement buttonDownloadACopyExcelOnline(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(downloadACopyButtonExcelOnline));
        }
        public IWebElement buttonOfficeAddInsExcelOnline(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(officeAddInsButtonExcelOnline));
        }
        public IWebElement buttonAddinsExcelOnline(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementToBeClickable(addinsButtonExcelOnline));
        }
        public IWebElement buttonMoreAddinsExcelOnline(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(moreAddinsButtonExcelOnline));
        }
        public IWebElement buttonAdvancedAddinsExcelOnline(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(advancedButtonExcelOnline));
        }
        public IWebElement tabMYADDINSExcelOnline(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(myADDINSTabExcelOnline));
        }
        public IWebElement linkUploadMyAddInExcelOnline(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(uploadMyAddInlinkExcelOnline));
        }
        public IWebElement myAddInBrowserFileExcelOnline(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(d=>d.FindElement(browserFileMyAddInExcelOnline));
        }
        public IWebElement buttonUploadUploadAddInExcelOnline(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(uploadButtonUploadAddInExcelOnline));
        }
        public IWebElement buttonFADAddInExcelOnline(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(fadAddinButtonExcelOnline));
        }
        public IWebElement buttonFADAddInInstanceExcelOnline(int timeoutInSeconds, int instance)
        {
            int InstanceNumber = 9 + instance;
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(@"//div["+ InstanceNumber.ToString() + "][contains(@class,'groupContainer-')]//button[@aria-label='FAD Add-in']")));
        }
        public IWebElement buttonCloseNotificationExcelOnline(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(notificationCloseButtonExcelOnline));
        }
        public IWebElement buttonDontShowInReviewMenuExcelOnline(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(dontShowButtonInReviewMenuExcelOnline));
        }
        public IWebElement buttonOkInvalidSheetExcelOnline(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(okInvalidSheetButtonExcelOnline));
        }
        public IWebElement buttonDeleteSheetAtRightClickExcelOnline(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(deleteSheetAtRightClickExcelOnline));
        }
        public IWebElement buttonGotItExcelOnline(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(gotItBtnExcelOnline));
        }
        public IWebElement byElement(int timeoutInSeconds, By by)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(by));
        }
    }

    internal class NavigationAction : BasePage<NavigationAction, NavigationPage>
    {
        #region NavigationPageAction constructor
        private NavigationAction(){}
        #endregion

        #region Navigation Page Items action
        // Wait for a frame to load Done then switch to frame (x)
        public NavigationAction SwitchToFrameWithWaitMethod(int timeoutInSeconds, By? element = null)
        {
            //check if have a specific element, if no then will get default element at Config.xml
            if (element == null) element = NavigationPage.frameIdFADAddInCurrentInstance(LoginPage.instanceName);

            if (timeoutInSeconds > 0)
            {
                WebDriverWait wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
                wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt(element));
            }
            return this;
        }

        // Wait for Element on Form to disappear
        public NavigationAction WaitForElementOnFormToDisappear(int timeoutInSeconds, By element)
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
        public NavigationAction WaitForElementVisible(int timeoutInSeconds, By element)
        {
            if (timeoutInSeconds > 0)
            {
                WebDriverWait wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
                wait.Until(ExpectedConditions.ElementIsVisible(element));
            }
            return this;
        }

        // Wait for element invisible
        public NavigationAction WaitForElementInvisible(int timeoutInSeconds, By element)
        {
            if (timeoutInSeconds > 0)
            {
                WebDriverWait wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(element));
            }
            return this;
        }

        // (wait method 1) Wait for page loading
        public NavigationAction WaitForNoneOfElementLoading(int timeoutInSeconds, By element)
        {
            if (timeoutInSeconds > 0)
            {
                WebDriverWait wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
                wait.Until(d => d.FindElement(element));
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
        public NavigationAction WaitingElementLoadDone(int timeout, By by)
        {
            int times = 0;
            while((!IsElementPresent(by)) && (timeout < times))
            {
                times++;
                System.Threading.Thread.Sleep(1000);
            }

            if(times == timeout)
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

        // Wait for a css attribute to change
        public NavigationAction WaitForACssAttributeChange(int timeoutInSeconds, By element)
        {
            if (timeoutInSeconds > 0)
            {
                WebDriverWait wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
                wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(element));
            }
            return this;
        }

        // Basic actions
        public string IsImageIconExcelOnlineInstanceShown(int timeoutInSeconds, int instance)
        {
            return this.Map.imageIconExcelOnlineInstance(timeoutInSeconds, instance).GetAttribute("src");
        }

        public NavigationAction ClickFundInfoTab(int timeoutInSeconds)
        {
            this.Map.tabFundInfo(timeoutInSeconds).Click();
            return this;
        }

        public NavigationAction ClickDataStatusTab(int timeoutInSeconds)
        {
            this.Map.tabDataStatus(timeoutInSeconds).Click();
            return this;
        }

        public NavigationAction ClickEditFundPencilIcon(int timeoutInSeconds)
        {
            this.Map.pencilIconEditFund(timeoutInSeconds).Click();
            return this;
        }

        public NavigationAction ClickAddNewFundButton(int timeoutInSeconds)
        {
            this.Map.buttonAddNewFund(timeoutInSeconds).Click();
            return this;
        }

        public NavigationAction ClickUploadCSVFileButton(int timeoutInSeconds)
        {
            this.Map.buttonuploadCSVFile(timeoutInSeconds).Click();
            return this;
        }

        public NavigationAction ClickUploadPDFFileIcon(int timeoutInSeconds)
        {
            this.Map.iconUploadPDFFile(timeoutInSeconds).Click();
            return this;
        }

        public NavigationAction ClickUserInputSection(int timeoutInSeconds)
        {
            this.Map.sectionUserInput(timeoutInSeconds).Click();
            return this;
        }

        public NavigationAction ClickDateSectionInUserInput(int timeoutInSeconds)
        {
            this.Map.userInputDateSection(timeoutInSeconds).Click();
            return this;
        }

        public NavigationAction ClickCRBMInUserInput(int timeoutInSeconds)
        {
            this.Map.userInputCRBMSection(timeoutInSeconds).Click();
            return this;
        }

        public NavigationAction ClickFeeModelSessionInUserInput(int timeoutInSeconds)
        {
            this.Map.userInputFeeModelSession(timeoutInSeconds).Click();
            return this;
        }
        public NavigationAction ClickMenuItemExcelOnline(int timeoutInSeconds, string name)
        {
            this.Map.itemMenuExcelOnline(timeoutInSeconds, name).Click();
            return this;
        }
        public NavigationAction ClickHomeMenuExcelOnline(int timeoutInSeconds)
        {
            this.Map.menuHomeExcelOnline(timeoutInSeconds).Click();
            return this;
        }
        public NavigationAction ClickInsertMenuExcelOnline(int timeoutInSeconds)
        {
            this.Map.menuInsertExcelOnline(timeoutInSeconds).Click();
            return this;
        }
        public NavigationAction ClickFileMenuExcelOnline(int timeoutInSeconds)
        {
            this.Map.menuFileExcelOnline(timeoutInSeconds).Click();
            return this;
        }
        public NavigationAction ClickMoreFileOptionButtonExcelOnline(int timeoutInSeconds)
        {
            this.Map.buttonMoreFileOptionExcelOnline(timeoutInSeconds).Click();
            return this;
        }
        public NavigationAction ClickSaveAsButtonExcelOnline(int timeoutInSeconds)
        {
            this.Map.buttonSaveAsExcelOnline(timeoutInSeconds).Click();
            return this;
        }
        public NavigationAction ClickDownloadACopyButtonExcelOnline(int timeoutInSeconds)
        {
            this.Map.buttonDownloadACopyExcelOnline(timeoutInSeconds).Click();
            return this;
        }

        public NavigationAction ClickAddinsButtonExcelOnline(int timeoutInSeconds)
        {
            this.Map.buttonAddinsExcelOnline(timeoutInSeconds).Click();
            return this;
        }

        public NavigationAction ClickAddinsButtonAndWaitPopupExcelOnline(int timeoutInSeconds)
        {
            this.Map.buttonAddinsExcelOnline(timeoutInSeconds).Click();

            int time = 0;
            while (IsElementPresent(NavigationPage.popularAddinsPopup) == false && time < timeoutInSeconds)
            {
                if (IsElementPresent(NavigationPage.popularAddinsPopup) == true)
                { break; }

                if (time == timeoutInSeconds)
                { Console.WriteLine("Timeout - Click Add-in button failed!"); }

                this.Map.buttonAddinsExcelOnline(timeoutInSeconds).Click();
                Thread.Sleep(1000); time++;
            }

            return this;
        }

        public NavigationAction ClickMoreAddinsButtonExcelOnline(int timeoutInSeconds)
        {
            this.Map.buttonMoreAddinsExcelOnline(timeoutInSeconds).Click();
            return this;
        }

        public NavigationAction ClickAdvancedButtonExcelOnline(int timeoutInSeconds)
        {
            this.Map.buttonAdvancedAddinsExcelOnline(timeoutInSeconds).Click();
            return this;
        }

        public NavigationAction ClickMYADDINSTabExcelOnline(int timeoutInSeconds)
        {
            //this.Map.tabMYADDINSExcelOnline(timeoutInSeconds).Click(); // --> element click intercepted

            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].click();", this.Map.tabMYADDINSExcelOnline(timeoutInSeconds));
            return this;
        }

        public NavigationAction ClickOfficeAddInsButtonExcelOnline(int timeoutInSeconds)
        {
            this.Map.buttonOfficeAddInsExcelOnline(timeoutInSeconds).Click();
            return this;
        }

        public NavigationAction ClickUploadMyAddInLinkExcelOnline(int timeoutInSeconds)
        {
            this.Map.linkUploadMyAddInExcelOnline(timeoutInSeconds).Click();
            return this;
        }

        public NavigationAction SelectBrowserFileMyAddInExcelOnline(int timeoutInSeconds, string file)
        {
            this.Map.myAddInBrowserFileExcelOnline(timeoutInSeconds).SendKeys(file);
            return this;
        }

        public NavigationAction ClickUploadButtonUploadAddInExcelOnline(int timeoutInSeconds)
        {
            this.Map.buttonUploadUploadAddInExcelOnline(timeoutInSeconds).Click();
            return this;
        }

        public NavigationAction ClickFADAddInButtonExcelOnline(int timeoutInSeconds)
        {
            //this.Map.buttonFADAddInExcelOnline(timeoutInSeconds).Click(); // --> element click intercepted
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].click();", this.Map.buttonFADAddInExcelOnline(timeoutInSeconds));
            return this;
        }

        public NavigationAction ClickFADAddInButtonInstanceExcelOnline(int timeoutInSeconds, int instance)
        {
            this.Map.buttonFADAddInInstanceExcelOnline(timeoutInSeconds, instance).Click();
            return this;
        }

        public NavigationAction ClickNotificationCloseButtonExcelOnline(int timeoutInSeconds)
        {
            //this.Map.buttonFADAddInExcelOnline(timeoutInSeconds).Click(); // --> element click intercepted
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].click();", this.Map.buttonCloseNotificationExcelOnline(timeoutInSeconds));
            return this;
        }

        public NavigationAction ClickDontShowButtonInReviewMenuExcelOnline(int timeoutInSeconds)
        {
            this.Map.buttonDontShowInReviewMenuExcelOnline(timeoutInSeconds).Click();
            return this;
        }

        public NavigationAction ClicKOkButtonInvalidSheetExcelOnline(int timeoutInSeconds)
        {
            this.Map.buttonOkInvalidSheetExcelOnline(timeoutInSeconds).Click();
            return this;
        }

        public NavigationAction ClickDeleteButtonSheetAtRightClickExcelOnline(int timeoutInSeconds)
        {
            this.Map.buttonDeleteSheetAtRightClickExcelOnline(timeoutInSeconds).Click();
            return this;
        }

        // Built-in Actions
        public void MouseClick(IWebElement target)
        {
            var builder = new Actions(Driver.Browser);
            builder.Click(target);
            builder.Perform();
        }
        public void MouseClick(int timeoutInSeconds, By target)
        {
            var builder = new Actions(Driver.Browser);
            builder.Click(this.Map.byElement(timeoutInSeconds, target));
            builder.Perform();
        }

        public void RightClick(IWebElement target)
        {
            var builder = new Actions(Driver.Browser);
            builder.ContextClick(target);
            builder.Perform();
        }

        public void UploadManifestExcelOnline(string manifestFilePath, string manifest)
        {
            By menuAutomateExcelOnline = By.XPath(@"//button[@id='Automate']");
            By dropdownEditingPencilExcelOnline = By.XPath(@"//div[@id='ModeSwitcher-container']");
            By buttonAnalyzeDataExcelOnline = By.XPath(@"//button[@id='Automate']");
            const string dialogChangeInReviewMenu = @"Discover new changes";

            // This will shift focus back to main (default) content in which frame 'one' lies
            Driver.Browser.SwitchTo().DefaultContent();

            // Wait for frame(x) to load Done and then switch to the frame of Office Document Body (to interact with menu/button of MS Excel Online)
            SwitchToFrameWithWaitMethod(40, NavigationPage.frameOfficeExcelOnline);

            // Wait for Menu/Button of MS Excel Online load Done
            WaitForElementVisible(30, menuAutomateExcelOnline)
            .WaitForElementVisible(30, dropdownEditingPencilExcelOnline)
            .WaitForElementVisible(30, buttonAnalyzeDataExcelOnline);

            // Check if Review menu show dialog "Discover new changes" then click "Don't show" button
            if (IsElementPresent(NavigationPage.dialogNotificationExcelOnline(dialogChangeInReviewMenu)))
            {
                ClickDontShowButtonInReviewMenuExcelOnline(10);
            }

            // Click on Insert menu of MS Excel Online
            ClickInsertMenuExcelOnline(30)
            .ClickOfficeAddInsButtonExcelOnline(10);

            // Wait for frame(x) to load Done and then Switch to the frame (of Office Add In ExcelOnline popup)
            SwitchToFrameWithWaitMethod(20, NavigationPage.frameDialogAddinExcelOnline);

            // Click on Upload My Add In link
            ClickUploadMyAddInLinkExcelOnline(10);

            // Upload Manifest file
            SelectBrowserFileMyAddInExcelOnline(10, manifestFilePath + manifest)
            .ClickUploadButtonUploadAddInExcelOnline(10);
        }

        public void UploadManifestExcelOnlineNoGodaddy(string manifestFilePath, string manifest)
        {
            By dropdownEditingPencilExcelOnline = By.XPath(@"//div[@id='ModeSwitcher-container']");
            By bookSavedTitleLoadDone = By.XPath(@"//i[contains(@data-icon-name,'CloudCheckmark') and @aria-hidden='true']"); // By.XPath(@"//span[@data-unique-id='DocumentTitleSaveStatus']");
            const string dialogChangeInReviewMenu = @"Discover new changes";
            
            // This will shift focus back to main (default) content in which frame 'one' lies
            Driver.Browser.SwitchTo().DefaultContent();

            // Wait for frame(x) to load Done and then switch to the frame of Office Document Body (to interact with menu/button of MS Excel Online)
            SwitchToFrameWithWaitMethod(40, NavigationPage.frameOfficeExcelOnline);

            // Wait for Menu/Button of MS Excel Online load Done
            //WaitForElementVisible(30, menuAutomateExcelOnline)
            WaitForElementVisible(30, dropdownEditingPencilExcelOnline).
            WaitForElementVisible(30, bookSavedTitleLoadDone);

            // Check if Review menu show dialog "Discover new changes" then click "Don't show" button
            if (IsElementPresent(NavigationPage.dialogNotificationExcelOnline(dialogChangeInReviewMenu)))
            {
                ClickDontShowButtonInReviewMenuExcelOnline(10);
            }

            /* old: click on Insert menu 
            // Click on Insert menu of MS Excel Online
            ClickInsertMenuExcelOnline(30)
            .ClickOfficeAddInsButtonExcelOnline(10);

            // Wait for frame(x) to load Done and then Switch to the frame (of Office Add In ExcelOnline popup)
            SwitchToFrameWithWaitMethod(20, NavigationPage.frameDialogAddinExcelOnline);

            // Click on Upload My Add In link
            ClickUploadMyAddInLinkExcelOnline(10);

            // Upload Manifest file
            SelectBrowserFileMyAddInExcelOnline(10, manifestFilePath + manifest)
            .ClickUploadButtonUploadAddInExcelOnline(10);
            */

            // Click on Home menu of MS Excel Online
            ClickHomeMenuExcelOnline(30);
            ClickAddinsButtonAndWaitPopupExcelOnline(30); //ClickAddinsButtonExcelOnline(30);
            WaitForElementVisible(30, NavigationPage.popularAddinsPopup);
            ClickAdvancedButtonExcelOnline(30); //ClickMoreAddinsButtonExcelOnline(30);

            // Wait for frame(x) to load Done and then Switch to the frame (of Office Add In ExcelOnline popup)
            SwitchToFrameWithWaitMethod(20, NavigationPage.frameDialogAddinExcelOnline);

            // Click on 'MY ADD INS' tab 
            //ClickMYADDINSTabExcelOnline(30);

            // Click on Upload My Add In link
            ClickUploadMyAddInLinkExcelOnline(10);

            // Upload Manifest file
            SelectBrowserFileMyAddInExcelOnline(10, manifestFilePath + manifest)
            .ClickUploadButtonUploadAddInExcelOnline(10);
        }

        public NavigationAction DownloadExcelOnlineReport(int timeoutInSeconds)
        {
            // Click on File menu of Excel Online
            ClickFileMenuExcelOnline(timeoutInSeconds); Thread.Sleep(2000);

            // Check if More File Option button is shown then Click on it
            if (IsElementPresent(NavigationPage.moreFileOptionButtonExcelOnline))
            {
                ClickMoreFileOptionButtonExcelOnline(timeoutInSeconds)
                .ClickSaveAsButtonExcelOnline(timeoutInSeconds);
            }
            if (IsElementPresent(NavigationPage.saveAsButtonExcelOnline))
            {
                ClickSaveAsButtonExcelOnline(timeoutInSeconds);
                Thread.Sleep(1000);
                // Click on Download A Copy Button (at Excel Online)
                ClickDownloadACopyButtonExcelOnline(timeoutInSeconds);
            }
            if (IsElementPresent(NavigationPage.menuItemExcelOnline("Create a Copy")))
            {
                ClickMenuItemExcelOnline(timeoutInSeconds, "Create a Copy"); // Click 'Create a Copy' menu
                Thread.Sleep(1000);
                // Click on Download A Copy Button (at Excel Online)
                ClickMenuItemExcelOnline(timeoutInSeconds, "Download a Copy"); // Click 'Download a Copy' menu
            }            
            return this;
        }

        public NavigationAction CheckifSheetExistingThenDelete(int timeoutInSeconds, string sheetName)
        {
            // This will shift focus back to main (default) content in which frame 'one' lies
            Driver.Browser.SwitchTo().DefaultContent();

            // Switch to the frame of Office Document Body (to interact with menu/button of MS Excel Online)
            SwitchToFrameWithWaitMethod(timeoutInSeconds, NavigationPage.frameOfficeExcelOnline);

            // Check if the "xxx" sheet is existing in workbook then delete it
            if (IsElementPresent(NavigationPage.workBenchSheetName(sheetName)))
            {
                IWebElement iwebSheetName = Driver.Browser.FindElement(NavigationPage.workBenchSheetName(sheetName));
                RightClick(iwebSheetName);
                ClickDeleteButtonSheetAtRightClickExcelOnline(timeoutInSeconds);
                ClicKOkButtonInvalidSheetExcelOnline(timeoutInSeconds);
                WaitForElementInvisible(10, NavigationPage.workBenchSheetName(sheetName));
            }

            // Switch to the FAD Add-in frame (to interact with Workbench for Excel Online)
            string? instanceName = null;
            string url = LoginPage.url;
            if (url.Contains("sandbox"))
            {
                instanceName = "sandbox";
            }
            if (url.Contains("conceptia"))
            {
                instanceName = "conceptia";
            }
            SwitchToFrameWithWaitMethod(10, NavigationPage.frameIdFADAddInCurrentInstance(instanceName));
            return this;
        }

        #endregion
    }
}
