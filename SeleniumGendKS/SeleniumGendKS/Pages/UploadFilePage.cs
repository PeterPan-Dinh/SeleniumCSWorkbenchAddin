using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using SeleniumGendKS.Core.BaseClass;
using SeleniumGendKS.Core.Selenium;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Nodes;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace SeleniumGendKS.Pages
{
    internal class UploadFilePage : BasePageElementMap
    {
        // Initiate variables
        /// <summary>
        /// Public Fund
        /// </summary>
        internal static WebDriverWait? wait;
        internal static string currency = "currency";
        internal static string date = "date";
        internal static string share_class_id = "share_class_id";
        internal static string exposure = "exposure";
        internal static string aum = "aum";
        internal static string gross = "gross";
        internal static string net = "net";
        internal static string other_amt_1 = "other_amt_1";
        internal static string other_amt_2 = "other_amt_2";
        internal static string other_amt_3 = "other_amt_3";
        /// <summary>
        /// Private Fund
        /// </summary>
        internal static string attribution_category_1 = "attribution_category_1";
        internal static string attribution_category_2 = "attribution_category_2";
        internal static string attribution_category_3 = "attribution_category_3";
        internal static string company_name = "company_name";
        internal static string contribution = "contribution";
        internal static string custom_weight = "custom_weight";
        internal static string data_as_of_date = "data_as_of_date";
        internal static string distribution = "distribution";
        internal static string entry_date = "entry_date";
        internal static string exit_date = "exit_date";
        internal static string fund = "fund";
        internal static string fund_name = "fund_name";
        internal static string fund_size = "fund_size";
        internal static string gross_irr = "gross_irr";
        internal static string gross_tvpi = "gross_tvpi";
        internal static string net_irr = "net_irr";
        internal static string net_tvpi = "net_tvpi";
        internal static string invested_capital = "invested_capital";
        internal static string realized = "realized";
        internal static string realized_capital = "realized_capital";
        internal static string status = "status";
        internal static string unrealized_fmv = "unrealized_fmv";
        internal static string unrealized_current_nav = "unrealized_current_nav";
        internal static string vintage_year = "vintage_year";

        // Initiate the By objects for elements
        /// <summary>
        /// Upload CSV files - Exposure, Fund AUM, Fund Returns
        /// </summary>
        internal static By dragAndDropFileHereGrid = By.XPath(@"//div[contains(@class,'files-input')]");
        internal static By uploadFileButton = By.XPath(@"//button[contains(.,'Upload File')]");
        internal static By uploadLabelButton(string label) => By.XPath(@"//button[contains(.,'" + label + "')]"); // KS-614 (Public) change Done to Replace (default)
        internal static By uploadFileDropdown = By.XPath(@"//p-splitbutton[@styleclass='upload-options']//button[@icon='pi pi-chevron-down']"); // KS-614
        internal static By uploadItemDropdown(string item) => By.XPath(@"//li[.='" + item + "']");
        internal static By uploadFile = By.XPath(@"//input[@type='file']");
        internal static By labelDropdown(string label) => By.XPath(@"//label[.='" + label + "']/preceding-sibling::p-dropdown");
        internal static By dropdownSelect(string item) => By.XPath(@"//li[@aria-label='" + item + "']");
        internal static By overlayDropdown = By.XPath(@"//div[contains(@class, 'overlayAnimation')]");
        internal static By asOfDateInputTxt = By.XPath(@"//span[contains(@class,'p-calendar')]/input");
        internal static By asOfDateButton = By.XPath(@"//button[contains(@class,'datepicker')]");
        internal static By cancelButton = By.XPath(@"//button[@label='Cancel']");
        internal static By doneButton = By.XPath(@"//button[@label='Done']");
        internal static By toastMessage(int number) => By.XPath(@"//p-toastitem[contains(@class,'toastAnimation')]//div[@role='alert']/div/div["+number+"]");
        internal static By toastMessageInvalidFile = By.XPath(@"//div[contains(@class,'p-message-wrapper')]/span[contains(@class,'message-summary')]");
        internal static By toastMessageInvalidFileDetail = By.XPath(@"//div[contains(@class,'p-message-wrapper')]/span[contains(@class,'message-detail')]");
        internal static By closeToastMessageButton = By.XPath(@"//button[contains(@class,'message-close')]");
        internal static By destinationDropdown(int row) => By.XPath(@"//tbody[@class='ng-star-inserted']/tr[" + row + "]/td/p-dropdown");
        internal static By destinationDropdownItemSelect(int row, string item) => By.XPath(@"//tbody[@class='ng-star-inserted']/tr[" + row + "]/td/p-dropdown//li[@aria-label='" + item + "']");
        internal static By dialogWarning = By.XPath(@"//div[contains(@class,'confirm-dialog')]");
        internal static By buttonInDialog(string label) => By.XPath(@"//div[contains(@class,'p-dialog-footer')]/button[.='" + label + "']");

        // Initiate the elements
        /// <summary>
        /// Upload CSV files - Exposure, Fund AUM, Fund Returns
        /// </summary>
        public IWebElement gridDragAndDropFileHere(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(dragAndDropFileHereGrid));
        }
        public IWebElement buttonUploadFile(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(uploadFileButton));
        }
        public IWebElement buttonUploadLabel(int timeoutInSeconds, string label)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(uploadLabelButton(label)));
        }
        public IWebElement dropdownUploadFile(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(uploadFileDropdown));
        }
        public IWebElement dropdownUploadItem(int timeoutInSeconds, string item)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(uploadItemDropdown(item)));
        }
        public IWebElement fileUpload => Driver.Browser.FindElement(uploadFile);
        public IWebElement dropdownLabel(int timeoutInSeconds, string label)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(labelDropdown(label)));
        }
        public IWebElement selectItemDropdown(int timeoutInSeconds, string item)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(dropdownSelect(item)));
        }
        public IWebElement inputTxtAsOfDate => Driver.Browser.FindElement(asOfDateInputTxt);
        public IWebElement buttonAsOfDate(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(asOfDateButton));
        }
        public IWebElement buttonCancel => Driver.Browser.FindElement(cancelButton);
        public IWebElement buttonDone => Driver.Browser.FindElement(doneButton);
        public IWebElement toastMessageAlert(int timeoutInSeconds,int number)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(toastMessage(number)));
        }
        public IWebElement invalidFileToastMessage(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(toastMessageInvalidFile));
        }
        public IWebElement invalidFileDetailToastMessage(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(toastMessageInvalidFileDetail));
        }
        public IWebElement buttonCloseToastMessage(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(closeToastMessageButton));
        }
        public IWebElement dropdownDestination(int timeoutInSeconds, int row)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(destinationDropdown(row)));
        }
        public IWebElement selectItemDestinationDropdown(int timeoutInSeconds, int row, string item)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(destinationDropdownItemSelect(row, item)));
        }
        public IWebElement buttonInDialogPopup(int timeoutInSeconds, string label)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(buttonInDialog(label)));
        }
    }

    internal sealed class UploadFileAction : BasePage<UploadFileAction, UploadFilePage>
    {
        #region constructor
        private UploadFileAction() { }
        #endregion

        #region Items action
        // Wait for all of elements visible (use for dropdown on-overlay visible)
        public UploadFileAction WaitForAllElementsVisible(int timeoutInSeconds, By element)
        {
            if (timeoutInSeconds > 0)
            {
                WebDriverWait wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
                wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(element));
            }
            return this;
        }

        // Wait for element Invisible (can use for dropdown on-overlay Invisible)
        public UploadFileAction WaitForElementInvisible(int timeoutInSeconds, By element)
        {
            if (timeoutInSeconds > 0)
            {
                WebDriverWait wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(element));
            }
            return this;
        }

        public UploadFileAction ClickUploadFileButton(int timeoutInSeconds)
        {
            //this.Map.buttonUploadLabel(timeoutInSeconds, label).Click(); // --> element click intercepted
            //ScrollIntoViewJs(this.Map.buttonUploadLabel(timeoutInSeconds, label));

            // Try with javascript
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].click();", this.Map.buttonUploadFile(timeoutInSeconds));
            return this;
        }
        public UploadFileAction ClickUploadLabelButton(int timeoutInSeconds, string label)
        {
            // Try with javascript
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].click();", this.Map.buttonUploadLabel(timeoutInSeconds, label));
            return this;
        }
        public UploadFileAction ClickUploadFileDropdown(int timeoutInSeconds)
        {
            // Try with javascript
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].click();", this.Map.dropdownUploadFile(timeoutInSeconds));
            return this;
        }
        public UploadFileAction SelectItemInDropdownUpload(int timeoutInSeconds, string item)
        {
            // Try with javascript
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].click();", this.Map.dropdownUploadItem(timeoutInSeconds, item));
            return this;
        }
        public UploadFileAction ClickDragAndDropFileHereGrid(int timeoutInSeconds)
        {
            this.Map.gridDragAndDropFileHere(timeoutInSeconds).Click();
            return this;
        }
        public UploadFileAction UploadFileInput(string filepath)
        {
            this.Map.fileUpload.SendKeys(filepath);
            return this;
        }
        public UploadFileAction ClickLabelDropdown(int timeoutInSeconds, string label)
        {
            this.Map.dropdownLabel(timeoutInSeconds, label).Click();
            return this;
        }
        public UploadFileAction SelectItemInDropdown(int timeoutInSeconds, string item)
        {
            this.Map.selectItemDropdown(timeoutInSeconds, item).Click();
            return this;
        }
        public UploadFileAction InputAsOfDate(int timeoutInSeconds, string date)
        {
            this.Map.inputTxtAsOfDate.SendKeys(OpenQA.Selenium.Keys.Control + "a");
            this.Map.inputTxtAsOfDate.SendKeys(date);
            // this.Map.buttonAsOfDate(timeoutInSeconds).Click(); --> Element Click Intercepted Exception
            // Try with javascript if Element Click Intercepted Exception
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].click();", this.Map.buttonAsOfDate(timeoutInSeconds));
            WaitForElementInvisible(10, AddEditFundPage.overlayDropdown);
            return this;
        }
        public UploadFileAction ClickCancelButton()
        {
            this.Map.buttonCancel.Click();
            return this;
        }
        public UploadFileAction ClickButtonInDialog(int timeoutInSeconds, string label)
        {
            this.Map.buttonInDialogPopup(timeoutInSeconds, label).Click();
            return this;
        }
        public UploadFileAction ClickDoneButton()
        {
            this.Map.buttonDone.Click();
            return this;
        }
        public UploadFileAction ScrollIntoView(int timeoutInSeconds, string label)
        {
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].scrollIntoView(false);", this.Map.dropdownLabel(timeoutInSeconds, label));
            return this;
        }
        public UploadFileAction ScrollIntoViewJs(IWebElement iwebE) // scroll to element with JavaScript
        {
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].scrollIntoView(false);", iwebE);
            return this;
        }
        public string toastMessageAlertGetText(int timeoutInSeconds, int number)
        {
            return this.Map.toastMessageAlert(timeoutInSeconds, number).Text;
        }
        public string ToastMessageInvalidFileGetText(int timeoutInSeconds)
        {
            return this.Map.invalidFileToastMessage(timeoutInSeconds).Text;
        }
        public string ToastMessageInvalidFileDetailGetText(int timeoutInSeconds)
        {
            return this.Map.invalidFileDetailToastMessage(timeoutInSeconds).Text;
        }
        public UploadFileAction ClickCloseToastMessageButton(int timeoutInSeconds)
        {
            this.Map.buttonCloseToastMessage(timeoutInSeconds).Click();
            return this;
        }
        public UploadFileAction ClickDestinationDropdown(int timeoutInSeconds, int row)
        {
            OpenQA.Selenium.Interactions.Actions actions = new OpenQA.Selenium.Interactions.Actions(Driver.Browser);
            actions.MoveToElement(this.Map.dropdownDestination(timeoutInSeconds, row));
            actions.Click(this.Map.dropdownDestination(timeoutInSeconds, row)).Build().Perform();
            return this;
        }
        public UploadFileAction SelectItemInDestinationDropdown(int timeoutInSeconds, int row, string item)
        {
            ScrollIntoViewJs(this.Map.selectItemDestinationDropdown(timeoutInSeconds, row, item));
            //this.Map.selectItemDestinationDropdown(timeoutInSeconds, row, item).Click(); // --> element click intercepted

            // Try with javascript
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].click();", this.Map.selectItemDestinationDropdown(timeoutInSeconds, row, item));
            return this;
        }
        public UploadFileAction PressDownArrowKeyUntilElementIsVisible(int timeoutInSeconds)
        {
            System.Windows.Forms.SendKeys.SendWait(@"{DOWN}");
            string? item = Driver.Browser.FindElement(By.XPath(@"//tbody[@class='ng-star-inserted']/tr[1]/td/p-dropdown//ul//li")).Text;
            while (item.Contains("No results found") && timeoutInSeconds > 0)
            {
                System.Windows.Forms.SendKeys.SendWait(@"{DOWN}"); System.Threading.Thread.Sleep(1000);
                var refreshItem = Driver.Browser.FindElement(By.XPath(@"//tbody[@class='ng-star-inserted']/tr[1]/td/p-dropdown//ul//li")).Text;
                if (!refreshItem.Contains("No results found")) break;
                timeoutInSeconds--;
            }
            return this;
        }
        #endregion

        #region Built-in actions
        public UploadFileAction ClickAndSelectItemInDropdown(int timeoutInSeconds, string label, string item)
        {
            ScrollIntoView(timeoutInSeconds, label);
            ClickLabelDropdown(timeoutInSeconds, label);
            WaitForAllElementsVisible(10, UploadFilePage.overlayDropdown);
            WaitForAllElementsVisible(10, UploadFilePage.dropdownSelect(item));
            SelectItemInDropdown(timeoutInSeconds, item);
            WaitForElementInvisible(10, UploadFilePage.overlayDropdown);
            return this;
        }
        public UploadFileAction ClickAndSelectItemInDestinationDropdown(int timeoutInSeconds, int row, string item)
        {
            ScrollIntoViewJs(this.Map.dropdownDestination(timeoutInSeconds, row));
            ClickDestinationDropdown(timeoutInSeconds, row);
            WaitForAllElementsVisible(timeoutInSeconds, UploadFilePage.overlayDropdown);
            WaitForAllElementsVisible(timeoutInSeconds, UploadFilePage.destinationDropdownItemSelect(row, item));
            SelectItemInDestinationDropdown(timeoutInSeconds, row, item);
            WaitForElementInvisible(timeoutInSeconds, UploadFilePage.overlayDropdown);
            return this;
        }
        public bool CompareObjects(IOrderedEnumerable<JObject> list1, IOrderedEnumerable<JObject> list2)
        {
            bool result = false;
            foreach (var item1 in list1)
            {
                foreach(var item2 in list2)
                {
                    if (JToken.DeepEquals(item1, item2)!=true)
                    {
                        result = false;
                    }
                    else
                    {
                        result = true;
                    }
                }
            }
            return result;
        }
        public bool CompareJObjectsToString(IOrderedEnumerable<JObject> list1, IOrderedEnumerable<JObject> list2)
        {
            bool result;
            var jobject1 = string.Join(",", list1);
            var jobject2 = string.Join(",", list2);

            if (jobject1.Equals(jobject2))
            {
                result = true;
            }
            else
            {
                result = false;
            }

            return result;
        }
        #endregion
    }
}
