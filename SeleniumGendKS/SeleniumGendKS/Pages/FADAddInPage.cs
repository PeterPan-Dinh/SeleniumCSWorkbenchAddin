using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using SeleniumGendKS.Core.BaseClass;
using SeleniumGendKS.Core.Selenium;
using System.Collections.ObjectModel;
using System.Xml.Linq;

namespace SeleniumGendKS.Pages
{
    internal class FADAddInPage : BasePageElementMap
    {
        // Initiate variables
        internal static WebDriverWait? wait;
        internal static string dataSource = "Data source";
        internal static string dateSelection = " Date Selection ";
        internal static string weightingMethodology = "Weighting Methodology";
        internal static string attributionCategory1 = "Attribution Category 1";
        internal static string attributionCategory2 = "Attribution Category 2";
        internal static string attributionCategory3 = "Attribution Category 3";
        internal static string assetClass = "Asset Class";
        internal static string geography = "Geography";

        // Initiate the By objects for elements
        internal static By userInputSection = By.XPath(@"//span[.='User Input']");
        internal static By userInputSubSection(string sectionName) => By.XPath(@"//span[@class='action']/ancestor::div[@class='section-title' and contains(.,'" + sectionName + "')]");
        internal static By labelDropdown(string label) => By.XPath(@"//label[.='" + label + "']/preceding-sibling::p-dropdown");
        internal static By labelInputTxt(string label) => By.XPath(@"//label[.='" + label + "']/preceding-sibling::input");
        internal static By datePickerLabelButton(string label) => By.XPath(@"//label[contains(.,'" + label + "')]/preceding-sibling::p-calendar//button");
        internal static By datePickerTitleButton(string title) => By.XPath(@"//div[.='" + title + "']/following-sibling::div//button");
        internal static By labelSearchInputTxt(string label) => By.XPath(@"//label[.='" + label + "']/preceding-sibling::p-autocomplete//input");
        internal static By dropdownSelect(string item) => By.XPath(@"//li[@aria-label='" + item + "']");
        internal static By overlayDropdown = By.XPath(@"//div[contains(@class, 'overlayAnimation')]");
        internal static By fieldCheckbox(string label) => By.XPath(@"//label[.='" + label + "']");
        internal static By datePickerPrevButton = By.XPath(@"//button[contains(@class,'datepicker-prev')]");
        internal static By datePickerNextButton = By.XPath(@"//button[contains(@class,'datepicker-next')]");
        internal static By datePickerMonthOnTopButton = By.XPath(@"//button[contains(@class,'p-datepicker-month')]");
        internal static By datePickerYearOnTopButton = By.XPath(@"//button[contains(@class,'p-datepicker-year')]");
        internal static By datePickerMonthOrYearButton(string monthOrYear) => By.XPath(@"//span[.=' "+ monthOrYear + " ']");
        internal static By datePickerDateButton(string date) => By.XPath(@"//span[.='" + date + "' and not(contains(@class,'p-disabled'))]");
        internal static By addCRBMbutton = By.XPath(@"//th[@class='gend-action']/button");
        internal static By deleteCRBMbutton(string number) => By.XPath(@"//tbody[contains(@class,'datatable')]/tr[" + number.ToString() + "]//button");
        internal static By nameCRBMInputTxtRow(int number) => By.XPath(@"//tbody[contains(@class,'datatable')]/tr[" + number.ToString() + "]/td[@class='b-name']//input");
        internal static By nameCRBMLoadingSpinnerRow(int number) => By.XPath(@"//tbody[contains(@class,'datatable')]/tr[" + number.ToString() + "]/td[@class='b-name']//input/following-sibling::i");
        internal static By nameCRBMReturnOfResults(string benchmark) => By.XPath(@"//span[.='" + benchmark + "']");
        internal static By factsetIDCRBMReadOnlyFieldRow(int number) => By.XPath(@"//tbody[contains(@class,'datatable')]/tr[" + number.ToString() + "]/td[@class='faceset']//input");
        internal static By betaCRBMInputNumberRow(int number) => By.XPath(@"//tbody[contains(@class,'datatable')]/tr[" + number.ToString() + "]/td[@class='beta']//input");
        internal static By grossExposureCRBMInputNumberRow(int number) => By.XPath(@"//tbody[contains(@class,'datatable')]/tr[" + number.ToString() + "]/td[@class='exposure']//input");
        internal static By nameCRBMRedErrorMessageRow(int number) => By.XPath(@"//tbody[contains(@class,'datatable')]/tr[" + number.ToString() + "]/following-sibling::tr//div");
        internal static By nameCRBMErrorMessageRow(int number) => By.XPath(@"//tbody[contains(@class,'datatable')]/tr[" + number.ToString() + "]/td[@class='b-name']//div");
        /// Hurdle Benchmark
        internal static By addButtonInXTable(string attrValue) => By.XPath(@"//div[contains(@class,'" + attrValue + "')]//button[@icon='pi pi-plus']");
        internal static By deleteButtonInXTable(string attrValue, string number) => By.XPath(@"//div[contains(@class,'" + attrValue + "')]//tbody/tr[" + number + "]//button");
        internal static By nameBenchmarkXTableInputTxtRow(string attrValue, string number) => By.XPath(@"//div[contains(@class,'" + attrValue + "')]//tbody/tr[" + number + "]/td[@class='b-name']//input");
        internal static By nameBenchmarkXTableSpinnerRow(string attrValue, string number) => By.XPath(@"//div[contains(@class,'" + attrValue + "')]//tbody/tr[" + number + "]/td[@class='b-name']//input/following-sibling::i");
        internal static By exposureBenchmarkXTableInputTxtRow(string attrValue, string number) => By.XPath(@"//div[contains(@class,'" + attrValue + "')]//tbody/tr[" + number + "]/td[@class='exposure']//input");
        /// error message
        internal static By errorInvalidMessage = By.XPath(@"//div[contains(@class,'error invalid-message')]"); // in red text
        internal static By errorInvalidMessageContent(string content) => By.XPath(@"//div[contains(@class,'error invalid-message') and contains(.,'" + content + "')]"); // in red text
        internal static By runButton = By.XPath(@"//p-button[@type='submit']");

        // Initiate the elements
        public IWebElement sectionUserInput => Driver.Browser.FindElement(userInputSection);
        public IWebElement subSectionUserInput(int timeoutInSeconds, string sectionName)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(userInputSubSection(sectionName)));
        }
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
        public IWebElement checkboxField(string label) => Driver.Browser.FindElement(fieldCheckbox(label));
        public IWebElement inputTxtLabel(string label) => Driver.Browser.FindElement(labelInputTxt(label));
        public IWebElement inputTxtSearchLabel(string label) => Driver.Browser.FindElement(labelSearchInputTxt(label));
        public IWebElement buttondatePickerLabel(int timeoutInSeconds, string label)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(datePickerLabelButton(label)));
        }
        public IWebElement buttondatePickerTitle(int timeoutInSeconds, string title)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(datePickerTitleButton(title)));
        }
        public IWebElement buttonDatePickerPrev(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(datePickerPrevButton));
        }
        public IWebElement buttonDatePickerNext(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(datePickerNextButton));
        }
        public IWebElement buttonOnTopDatePickerMonth(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(datePickerMonthOnTopButton));
        }
        public IWebElement buttonOnTopDatePickerYear(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(datePickerYearOnTopButton));
        }
        public IWebElement buttonDatePickerMonthOrYear(int timeoutInSeconds, string monthOrYear)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(datePickerMonthOrYearButton(monthOrYear)));
        }
        public IWebElement buttonDatePickerDate(int timeoutInSeconds, string date)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(datePickerDateButton(date)));
        }
        public IWebElement buttonAddCRBM(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(addCRBMbutton));
        }
        public IWebElement buttonDeleteCRBM(int timeoutInSeconds, string number)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(deleteCRBMbutton(number)));
        }
        public IWebElement inputTxtNameCRBMRow(int timeoutInSeconds, int number)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(d=>d.FindElement(nameCRBMInputTxtRow(number)));
        }
        public IWebElement returnOfResultsNameCRBM(int timeoutInSeconds, string benchmark)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(nameCRBMReturnOfResults(benchmark)));
        }
        public IWebElement readOnlyFieldFactsetIDCRBMRow(int timeoutInSeconds, int number)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(d => d.FindElement(factsetIDCRBMReadOnlyFieldRow(number)));
        }
        public IWebElement inputNumberBetaCRBMRow(int timeoutInSeconds, int number)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(d => d.FindElement(betaCRBMInputNumberRow(number)));
        }
        public IWebElement inputNumberGrossExposureCRBMRow(int timeoutInSeconds, int number)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(d => d.FindElement(grossExposureCRBMInputNumberRow(number)));
        }
        public IWebElement redErrorMessageNameCRBMRow(int timeoutInSeconds, int number)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(nameCRBMRedErrorMessageRow(number)));
        }
        public IWebElement errorMessageNameCRBMRow(int timeoutInSeconds, int number)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(nameCRBMErrorMessageRow(number)));
        }
        public IWebElement buttonAddInXTable(int timeoutInSeconds, string attrValue)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(addButtonInXTable(attrValue)));
        }
        public IWebElement buttonDeleteInXTable(int timeoutInSeconds, string attrValue, string number)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(deleteButtonInXTable(attrValue, number)));
        }
        public IWebElement inputTxtNameBenchmarkXTableRow(int timeoutInSeconds, string attrValue, string number)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(d => d.FindElement(nameBenchmarkXTableInputTxtRow(attrValue, number)));
        }
        public IWebElement inputTxtExposureBenchmarkXTableRow(int timeoutInSeconds, string attrValue, string number)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(d => d.FindElement(exposureBenchmarkXTableInputTxtRow(attrValue, number)));
        }
        public IWebElement messageErrorInvalid(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(errorInvalidMessage));
        }
        public IWebElement buttonRun(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(runButton));
        }
    }

    internal sealed class FADAddInAction : BasePage<FADAddInAction, FADAddInPage>
    {
        #region constructor
        private FADAddInAction() { }
        #endregion

        #region Items action
        // Wait for loading Spinner icon to disappear
        public FADAddInAction WaitForLoadingIconToDisappear(int timeoutInSeconds, By element)
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
        public FADAddInAction WaitForElementVisible(int timeoutInSeconds, By element)
        {
            if (timeoutInSeconds > 0)
            {
                WebDriverWait wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
                wait.Until(ExpectedConditions.ElementIsVisible(element));
            }
            return this;
        }

        // Wait for all of elements visible (use for dropdown on-overlay visible)
        public FADAddInAction WaitForAllElementsVisible(int timeoutInSeconds, By element)
        {
            if (timeoutInSeconds > 0)
            {
                WebDriverWait wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
                wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(element));
            }
            return this;
        }

        // Wait for element Invisible (can use for dropdown on-overlay Invisible)
        public FADAddInAction WaitForElementInvisible(int timeoutInSeconds, By element)
        {
            if (timeoutInSeconds > 0)
            {
                WebDriverWait wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(element));
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

        public FADAddInAction ClickUserInputSection()
        {
            this.Map.sectionUserInput.Click();
            return this;
        }
        public FADAddInAction ClickUserInputSubSection(int timeoutInSeconds, string sectionName)
        {
            this.Map.subSectionUserInput(timeoutInSeconds, sectionName).Click();
            return this;
        }
        public FADAddInAction ClickLabelDropdown(int timeoutInSeconds, string label)
        {
            Actions actions = new Actions(Driver.Browser);
            actions.MoveToElement(this.Map.dropdownLabel(timeoutInSeconds, label));
            actions.Perform();
            //this.Map.dropdownLabel(timeoutInSeconds, label).Click();
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].click();", this.Map.dropdownLabel(timeoutInSeconds, label));
            return this;
        }
        public FADAddInAction SelectItemInDropdown(int timeoutInSeconds, string item)
        {
            ScrollIntoView(this.Map.selectItemDropdown(timeoutInSeconds, item));
            //this.Map.selectItemDropdown(timeoutInSeconds, item).Click(); --> element click intercepted
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].click();", this.Map.selectItemDropdown(timeoutInSeconds, item));
            return this;
        }
        public FADAddInAction ClickFieldCheckbox(string label)
        {
            ScrollIntoView(this.Map.checkboxField(label));
            // this.Map.checkboxField(label).Click(); --> element click intercepted
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].click();", this.Map.checkboxField(label));
            return this;
        }
        public FADAddInAction InputTxtLabelField(string label, string text)
        {
            this.Map.inputTxtLabel(label).Clear();
            this.Map.inputTxtLabel(label).SendKeys(text);
            return this;
        }
        public FADAddInAction InputTxtSearchLabelField(string label, string text)
        {
            this.Map.inputTxtSearchLabel(label).Clear();
            this.Map.inputTxtSearchLabel(label).SendKeys(text);
            return this;
        }
        public FADAddInAction ClickDatePickerTitleButton(int timeoutInSeconds, string title)
        {
            //this.Map.buttondatePickerTitle(timeoutInSeconds, title).Click(); --> Issue Element Click Intercepted Exception
            // Try with javascript if Element Click Intercepted Exception
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].click();", this.Map.buttondatePickerTitle(timeoutInSeconds, title));
            WaitForElementVisible(10, AddEditFundPage.overlayDropdown);
            return this;
        }
        public FADAddInAction ClickDatePickerPrevButton(int timeoutInSeconds)
        {
            this.Map.buttonDatePickerPrev(timeoutInSeconds).Click();
            return this;
        }
        public FADAddInAction ClickDatePickerNextButton(int timeoutInSeconds)
        {
            this.Map.buttonDatePickerNext(timeoutInSeconds).Click();
            return this;
        }
        public FADAddInAction ClickDatePickerMonthOnTopButton(int timeoutInSeconds)
        {
            this.Map.buttonOnTopDatePickerMonth(timeoutInSeconds).Click();
            return this;
        }
        public FADAddInAction ClickDatePickerYearOnTopButton(int timeoutInSeconds)
        {
            this.Map.buttonOnTopDatePickerYear(timeoutInSeconds).Click();
            return this;
        }
        public FADAddInAction ClickMonthOrYearInDatePicker(int timeoutInSeconds, string monthOrYear)
        {
            this.Map.buttonDatePickerMonthOrYear(timeoutInSeconds, monthOrYear).Click();
            return this;
        }
        public FADAddInAction ClickDateInDatePicker(int timeoutInSeconds, string date)
        {
            this.Map.buttonDatePickerDate(timeoutInSeconds, date).Click();
            return this;
        }
        public FADAddInAction ClickCRBMAddButton(int timeoutInSeconds)
        {
            // move to element
            ScrollIntoView(this.Map.buttonAddCRBM(timeoutInSeconds));
            Actions actions = new Actions(Driver.Browser);
            actions.MoveToElement(this.Map.buttonAddCRBM(timeoutInSeconds)).DoubleClick().Build().Perform();
            System.Threading.Thread.Sleep(2000);
            ClickCRBMDeleteButton(timeoutInSeconds, "last()");
            System.Threading.Thread.Sleep(500);

            //this.Map.buttonAddCRBM(timeoutInSeconds).Click(); --> element click intercepted
            //IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            //je.ExecuteScript("arguments[0].click();", this.Map.buttonAddCRBM(timeoutInSeconds));
            return this;
        }
        public FADAddInAction ClickCRBMDeleteButton(int timeoutInSeconds, string rowNumber)
        {
            PageDownToScrollDownPage();
            ScrollIntoView(this.Map.buttonDeleteCRBM(timeoutInSeconds, rowNumber));
            //this.Map.buttonDeleteCRBM(timeoutInSeconds, rowNumber).Click();
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].click();", this.Map.buttonDeleteCRBM(timeoutInSeconds, rowNumber));
            System.Threading.Thread.Sleep(500);
            return this;
        }
        public FADAddInAction InputTxtNameCRBMRow(int timeoutInSeconds, int number, string text)
        {
            ScrollIntoView(this.Map.inputTxtNameCRBMRow(timeoutInSeconds, number));
            this.Map.inputTxtNameCRBMRow(timeoutInSeconds, number).Clear();
            this.Map.inputTxtNameCRBMRow(timeoutInSeconds, number).SendKeys(text);
            return this;
        }
        public FADAddInAction ClickNameCRBMReturnOfResults(int timeoutInSeconds, string benchmark)
        {
            //this.Map.returnOfResultsNameCRBM(timeoutInSeconds, benchmark).Click(); //  --> element click intercepted
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].click();", this.Map.returnOfResultsNameCRBM(timeoutInSeconds, benchmark));
            return this;
        }
        public string ReadOnlyFieldBenchMarkIDGetTextRow(int timeoutInSeconds, int number)
        {
            return this.Map.readOnlyFieldFactsetIDCRBMRow(timeoutInSeconds, number).Text;
        }
        public FADAddInAction InputNumberBetaCRBMRow(int timeoutInSeconds, int rowNumber, string text)
        {
            this.Map.inputNumberBetaCRBMRow(timeoutInSeconds, rowNumber).Clear();
            this.Map.inputNumberBetaCRBMRow(timeoutInSeconds, rowNumber).SendKeys(text);
            return this;
        }
        public FADAddInAction InputNumberGrossExposureCRBMRow(int timeoutInSeconds, int number, string text)
        {
            this.Map.inputNumberGrossExposureCRBMRow(timeoutInSeconds, number).Clear();
            this.Map.inputNumberGrossExposureCRBMRow(timeoutInSeconds, number).SendKeys(text);
            return this;
        }
        public string RedErrorMessageNameCRBMGetTextRow(int timeoutInSeconds, int number)
        {
            return this.Map.redErrorMessageNameCRBMRow(timeoutInSeconds, number).Text;
        }
        public string ErrorMessageNameCRBMGetTextRow(int timeoutInSeconds, int number)
        {
            return this.Map.errorMessageNameCRBMRow(timeoutInSeconds, number).Text;
        }
        public string ErrorInvalidMessageGetText(int timeoutInSeconds) // in red text
        {
            return this.Map.messageErrorInvalid(timeoutInSeconds).Text;
        }
        public FADAddInAction ClickAddButtonInXTable(int timeoutInSeconds, string attrValue)
        {
            // move to element
            ScrollIntoView(Map.buttonAddInXTable(timeoutInSeconds, attrValue));
            Actions actions = new Actions(Driver.Browser);
            actions.MoveToElement(Map.buttonAddInXTable(timeoutInSeconds, attrValue)).DoubleClick().Build().Perform();
            Thread.Sleep(2000);
            ClickDeleteButtonInXTable(timeoutInSeconds, attrValue, "last()");
            Thread.Sleep(500);
            return this;
        }
        public FADAddInAction ClickDeleteButtonInXTable(int timeoutInSeconds, string attrValue, string rowNumber)
        {
            PageDownToScrollDownPage();
            ScrollIntoView(Map.buttonDeleteInXTable(timeoutInSeconds, attrValue, rowNumber));
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].click();", Map.buttonDeleteInXTable(timeoutInSeconds, attrValue, rowNumber));
            Thread.Sleep(500);
            return this;
        }
        public FADAddInAction InputTxtNameBenchmarkXTableRow(int timeoutInSeconds, string attrValue, string number, string text)
        {
            ScrollIntoView(Map.inputTxtNameBenchmarkXTableRow(timeoutInSeconds, attrValue, number));
            Map.inputTxtNameBenchmarkXTableRow(timeoutInSeconds, attrValue, number).Clear();
            Map.inputTxtNameBenchmarkXTableRow(timeoutInSeconds, attrValue, number).SendKeys(text);
            return this;
        }
        public FADAddInAction InputTxtExposureBenchmarkXTableRow(int timeoutInSeconds, string attrValue, string number, string text)
        {
            ScrollIntoView(Map.inputTxtExposureBenchmarkXTableRow(timeoutInSeconds, attrValue, number));
            Map.inputTxtExposureBenchmarkXTableRow(timeoutInSeconds, attrValue, number).Clear();
            Map.inputTxtExposureBenchmarkXTableRow(timeoutInSeconds, attrValue, number).SendKeys(text);
            return this;
        }
        public FADAddInAction HomeToScrollUpPage()
        {
            Actions actions = new Actions(Driver.Browser);
            actions.SendKeys(OpenQA.Selenium.Keys.Home).Build().Perform();
            System.Threading.Thread.Sleep(1000);
            return this;
        }
        public FADAddInAction PageUpToScrollUpPage()
        {
            Actions actions = new Actions(Driver.Browser);
            actions.SendKeys(OpenQA.Selenium.Keys.PageUp).Build().Perform();
            System.Threading.Thread.Sleep(1000);
            return this;
        }
        public FADAddInAction PageDownToScrollDownPage()
        {
            Actions actions = new Actions(Driver.Browser);
            actions.SendKeys(OpenQA.Selenium.Keys.PageDown).Build().Perform();
            System.Threading.Thread.Sleep(1000);
            return this;
        }
        public FADAddInAction ScrollIntoView(IWebElement iwebE)
        {
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].scrollIntoView(false);", iwebE);
            return this;
        }
        public FADAddInAction ClickRunButton(int timeoutInSeconds)
        {
            // Scroll to element
            ScrollIntoView(this.Map.buttonRun(timeoutInSeconds));

            // MouseHover and Click
            Actions action = new Actions(Driver.Browser);
            action.MoveToElement(this.Map.buttonRun(timeoutInSeconds)).DoubleClick(this.Map.buttonRun(timeoutInSeconds)).Perform();
            return this;
        }

        // action - verify
        public string IsDatePickerLabelShown(int timeoutInSeconds)
        {
            var wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(@"//div[.=' Date Selection ']/following-sibling::div//label"))).Text;
        }
        #endregion

        #region Built-in actions
        public FADAddInAction ClickAndSelectItemInDropdown(int timeoutInSeconds, string label, string item)
        {    
            if (item == "Custom")
            {
                ClickLabelDropdown(timeoutInSeconds, label);
                WaitForAllElementsVisible(10, FADAddInPage.overlayDropdown);
                WaitForAllElementsVisible(10, FADAddInPage.dropdownSelect(item));
                SelectItemInDropdown(timeoutInSeconds, item);
            }
            else
            {
                ScrollIntoView(this.Map.dropdownLabel(timeoutInSeconds, label));
                ClickLabelDropdown(timeoutInSeconds, label);
                //WaitForAllElementsVisible(10, FADAddInPage.overlayDropdown);
                int time = 0;
                while (IsElementPresent(FADAddInPage.overlayDropdown) == false && time < timeoutInSeconds)
                {
                    if (IsElementPresent(FADAddInPage.overlayDropdown) == true) { break; }
                    if (time == timeoutInSeconds) { Console.WriteLine("Timeout - Click dropdown failed!"); }
                    ScrollIntoView(this.Map.dropdownLabel(timeoutInSeconds, label));
                    this.Map.dropdownLabel(timeoutInSeconds, label).Click(); Thread.Sleep(1000);
                    time++;
                }

                WaitForAllElementsVisible(10, FADAddInPage.dropdownSelect(item));
                SelectItemInDropdown(timeoutInSeconds, item);
                WaitForElementInvisible(10, FADAddInPage.overlayDropdown);
            }
            return this;
        }
        public FADAddInAction ClickToUnCheckTheCheckbox(string label)
        {
            bool isChecked = false;
            isChecked = Driver.Browser.FindElement(AddEditFundPage.fieldCheckbox(label)).GetAttribute("class").Contains("active");
            if (isChecked)
            {
                ClickFieldCheckbox(label);
            }
            return this;
        }
        public FADAddInAction ClickToCheckTheCheckbox(string label)
        {
            bool isChecked = false;
            isChecked = Driver.Browser.FindElement(AddEditFundPage.fieldCheckbox(label)).GetAttribute("class").Contains("active");
            if (isChecked)
            {
                return this;
            }
            else ClickFieldCheckbox(label);
            return this;
        }
        public FADAddInAction InputTxtDatePickerLabel(int timeoutInSeconds, string label, string yyyy, string mmm, string dd)
        {
            // Try with javascript if Element Click Intercepted Exception
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].click();", this.Map.buttondatePickerLabel(timeoutInSeconds, label));
            WaitForElementVisible(10, AddEditFundPage.overlayDropdown);

            // Click Year button in Date-Picker (to select Year, month, date)
            WaitForElementVisible(timeoutInSeconds, FADAddInPage.datePickerYearOnTopButton);
            ClickDatePickerYearOnTopButton(timeoutInSeconds);
            ClickMonthOrYearInDatePicker(timeoutInSeconds, yyyy);
            ClickMonthOrYearInDatePicker(timeoutInSeconds, mmm);
            ClickDateInDatePicker(timeoutInSeconds, dd);
            WaitForElementInvisible(10, AddEditFundPage.overlayDropdown);
            return this;
        }
        public FADAddInAction InputTxtDatePickerTitle(int timeoutInSeconds, string title, string yyyy, string mmm, string dd)
        {
            // Click on Date Picker button
            ClickDatePickerTitleButton(timeoutInSeconds, title);
            WaitForElementVisible(10, AddEditFundPage.overlayDropdown);

            // Click Year button in Date-Picker (to select Year, month, date)
            WaitForElementVisible(timeoutInSeconds, FADAddInPage.datePickerYearOnTopButton);
            ClickDatePickerYearOnTopButton(timeoutInSeconds); Thread.Sleep(500);
            ClickMonthOrYearInDatePicker(timeoutInSeconds, yyyy); Thread.Sleep(500);
            ClickMonthOrYearInDatePicker(timeoutInSeconds, mmm); Thread.Sleep(500);
            ClickDateInDatePicker(timeoutInSeconds, dd); Thread.Sleep(500);
            WaitForElementInvisible(10, AddEditFundPage.overlayDropdown);
            return this;
        }

        public FADAddInAction CheckIfExistingCRBMThenDeleteAll(int timeoutInSeconds, string dataSource)
        {
            Thread.Sleep(1000);
            ReadOnlyCollection<IWebElement>? crbm = Driver.Browser.FindElements(By.XPath("//tbody[contains(@class,'datatable')]/tr"));
            if (crbm.Count == 1 && dataSource == "C")
            {
                ClickCRBMAddButton(timeoutInSeconds);
                System.Threading.Thread.Sleep(500);
                ClickCRBMDeleteButton(timeoutInSeconds, "1");
                System.Threading.Thread.Sleep(500);
            }
            if (crbm.Count > 1 && dataSource == "C") 
            {
                ClickCRBMAddButton(timeoutInSeconds);
                System.Threading.Thread.Sleep(500);
                int getCRBM = Driver.Browser.FindElements(By.XPath("//tbody[contains(@class,'datatable')]/tr")).Count;
                while (getCRBM > 1)
                {
                    ClickCRBMDeleteButton(timeoutInSeconds, "1");
                    System.Threading.Thread.Sleep(500);
                    getCRBM--;
                }
            }
            if (crbm.Count >= 1 && (dataSource == "A" || dataSource == "S"))
            {
                int getCRBM = Driver.Browser.FindElements(By.XPath("//tbody[contains(@class,'datatable')]/tr")).Count;
                while (getCRBM >= 1)
                {
                    ClickCRBMDeleteButton(timeoutInSeconds, "1");
                    System.Threading.Thread.Sleep(500);
                    getCRBM--;
                }
            }
            return this;
        }

        public FADAddInAction WaitforExcelWebRenderDone(int timeoutInSeconds)
        {
            // This will shift focus back to main (default) content in which frame 'one' lies
            Driver.Browser.SwitchTo().DefaultContent();

            // Switch to the frame of Office Document Body (to interact with menu/button of MS Excel Online)
            NavigationAction.Instance.SwitchToFrameWithWaitMethod(timeoutInSeconds, NavigationPage.frameOfficeExcelOnline);

            // Check if the dialog "Invalid sheet" is displayed then click on OK button
            System.Threading.Thread.Sleep(2000);
            if (IsElementPresent(NavigationPage.okInvalidSheetButtonExcelOnline))
            {
                NavigationAction.Instance.ClicKOkButtonInvalidSheetExcelOnline(timeoutInSeconds);
            }

            // Wait for Sheet "WorkBench" move to the last sheet 
            string workBenchSheet = @"//ul[@class='ewa-stb-tabs']/li[last()]/a[@aria-label='WorkBench']";
            //WaitForElementVisible(timeoutInSeconds, By.XPath(workBenchSheet));

            // check if sheet 'WorkBench' not yet moved to the last sheet then wait for it
            if (IsElementPresent(By.XPath(workBenchSheet)))
            {
                WaitForElementInvisible(timeoutInSeconds, By.XPath(workBenchSheet));
            }

            // Check if the dialog "Invalid sheet" is displayed then click on OK button
            if (IsElementPresent(NavigationPage.okInvalidSheetButtonExcelOnline))
            {
                NavigationAction.Instance.ClicKOkButtonInvalidSheetExcelOnline(timeoutInSeconds);
            }

            // Wait for Sheet "Raw Data" move to the last sheet
            string rawDataSheet = @"//ul[@class='ewa-stb-tabs']/li[last()]/a[@aria-label='Raw Data']";
            WaitForElementVisible(timeoutInSeconds, By.XPath(rawDataSheet));

            // Wait for Sheet "GraphData" move to the last sheet
            string graphDataSheet = @"//ul[@class='ewa-stb-tabs']/li[last()]/a[@aria-label='GraphData']";
            WaitForElementVisible(timeoutInSeconds, By.XPath(graphDataSheet));

            // Wait for the title of Workbook (Saving ...) is disappeared
            //string bookSavedTitleLoadDone = @"//span[@data-unique-id='DocumentTitleSaveStatus' and .='Saving...']"; // begin is 'Saved'
            string bookSavedTitleLoadDone = @"//i[contains(@data-icon-name,'ArrowSync') and @aria-hidden='true']"; // begin is 'Saved'
            WaitForElementVisible(timeoutInSeconds, By.XPath(bookSavedTitleLoadDone));
            WaitForElementInvisible(timeoutInSeconds, By.XPath(bookSavedTitleLoadDone));
            System.Threading.Thread.Sleep(11000);
            // Check if the dialog "Invalid sheet" is displayed then click on OK button
            if (IsElementPresent(NavigationPage.okInvalidSheetButtonExcelOnline))
            {
                NavigationAction.Instance.ClicKOkButtonInvalidSheetExcelOnline(timeoutInSeconds); Thread.Sleep(1000);
            }
            Thread.Sleep(1000);
            return this;
        }
        public FADAddInAction WaitforExcelWebRenderDone2(int timeoutInSeconds)
        {
            // This will shift focus back to main (default) content in which frame 'one' lies
            Driver.Browser.SwitchTo().DefaultContent();

            // Switch to the frame of Office Document Body (to interact with menu/button of MS Excel Online)
            NavigationAction.Instance.SwitchToFrameWithWaitMethod(timeoutInSeconds, NavigationPage.frameOfficeExcelOnline);

            // Check if the dialog "Invalid sheet" is displayed then click on OK button
            System.Threading.Thread.Sleep(2000);
            if (IsElementPresent(NavigationPage.okInvalidSheetButtonExcelOnline))
            {
                NavigationAction.Instance.ClicKOkButtonInvalidSheetExcelOnline(timeoutInSeconds);
            }

            // Wait for Sheet "WorkBench" move to the last sheet 
            string workBenchSheet = @"//ul[@class='ewa-stb-tabs']/li[last()]/a[@aria-label='WorkBench']";

            // check if sheet 'WorkBench' not yet moved to the last sheet then wait for it
            if (IsElementPresent(By.XPath(workBenchSheet)))
            {
                WaitForElementInvisible(timeoutInSeconds, By.XPath(workBenchSheet));
            }

            // Check if the dialog "Invalid sheet" is displayed then click on OK button
            if (IsElementPresent(NavigationPage.okInvalidSheetButtonExcelOnline))
            {
                NavigationAction.Instance.ClicKOkButtonInvalidSheetExcelOnline(timeoutInSeconds);
            }

            // Wait for Sheet "GraphData" move to the last sheet
            string graphDataSheet = @"//ul[@class='ewa-stb-tabs']/li[last()]/a[@aria-label='GraphData']";
            WaitForElementVisible(timeoutInSeconds, By.XPath(graphDataSheet)); Thread.Sleep(1000);

            // Wait for the title of Workbook (Saving ...) is disappeared
            string bookSavedTitleLoadDone = "//*[@data-unique-id='DocumentTitleSaveStatus']";
            
            // Check if the status "Saving.../Saved" is displayed then wait for it disappeared
            if (IsElementPresent(By.XPath(bookSavedTitleLoadDone)))
            {
                WaitForElementVisible(timeoutInSeconds, By.XPath(bookSavedTitleLoadDone));
                WaitForElementInvisible(timeoutInSeconds, By.XPath(bookSavedTitleLoadDone));
            }
            Thread.Sleep(11000);

            // Check if the dialog "Invalid sheet" is displayed then click on OK button
            if (IsElementPresent(NavigationPage.okInvalidSheetButtonExcelOnline))
            {
                NavigationAction.Instance.ClicKOkButtonInvalidSheetExcelOnline(timeoutInSeconds); Thread.Sleep(1000);
            }
            Thread.Sleep(1000);
            return this;
        }
        public FADAddInAction WaitforExcelWebRenderGraphDataAtLastSheet(int timeoutInSeconds)
        {
            // This will shift focus back to main (default) content in which frame 'one' lies
            Driver.Browser.SwitchTo().DefaultContent();

            // Switch to the frame of Office Document Body (to interact with menu/button of MS Excel Online)
            NavigationAction.Instance.SwitchToFrameWithWaitMethod(timeoutInSeconds, NavigationPage.frameOfficeExcelOnline);

            // Check if the dialog "Invalid sheet" is displayed then click on OK button
            System.Threading.Thread.Sleep(2000);
            if (IsElementPresent(NavigationPage.okInvalidSheetButtonExcelOnline))
            {
                NavigationAction.Instance.ClicKOkButtonInvalidSheetExcelOnline(timeoutInSeconds);
            }

            // Wait for Sheet "WorkBench" move to the last sheet 
            string workBenchSheet = @"//ul[@class='ewa-stb-tabs']/li[last()]/a[@aria-label='WorkBench']";

            // check if sheet 'WorkBench' not yet moved to the last sheet then wait for it
            if (IsElementPresent(By.XPath(workBenchSheet)))
            {
                WaitForElementInvisible(timeoutInSeconds, By.XPath(workBenchSheet));
            }

            // Check if the dialog "Invalid sheet" is displayed then click on OK button
            if (IsElementPresent(NavigationPage.okInvalidSheetButtonExcelOnline))
            {
                NavigationAction.Instance.ClicKOkButtonInvalidSheetExcelOnline(timeoutInSeconds);
            }

            // Wait for Sheet "GraphData" move to the last sheet
            string graphDataSheet = @"//ul[@class='ewa-stb-tabs']/li[last()]/a[@aria-label='GraphData']";
            WaitForElementVisible(timeoutInSeconds, By.XPath(graphDataSheet)); Thread.Sleep(1000);
            return this;
        }

        public FADAddInAction WaitforExcelWebDxDOutputRenderDone(int timeoutInSeconds)
        {
            // This will shift focus back to main (default) content in which frame 'one' lies
            Driver.Browser.SwitchTo().DefaultContent();

            // Switch to the frame of Office Document Body (to interact with menu/button of MS Excel Online)
            NavigationAction.Instance.SwitchToFrameWithWaitMethod(timeoutInSeconds, NavigationPage.frameOfficeExcelOnline);

            // Wait for Sheet "DxD Output" move to the last sheet 
            string dxDOutputSheet = @"//ul[@class='ewa-stb-tabs']/li[last()]/a[@aria-label='DxD Output']";
            WaitForElementVisible(timeoutInSeconds, By.XPath(dxDOutputSheet));

            // This will shift focus back to main (default) content in which frame 'one' lies
            Driver.Browser.SwitchTo().DefaultContent();

            // Switch to the frame of Office Document Body (to interact with menu/button of MS Excel Online)
            NavigationAction.Instance.SwitchToFrameWithWaitMethod(timeoutInSeconds, NavigationPage.frameOfficeExcelOnline);
            Thread.Sleep(1000);

            // Wait for the title of Workbook (Saving ...) is disappeared
            //string bookSavedTitleLoadDone = @"//span[@data-unique-id='DocumentTitleSaveStatus' and .='Saving...']"; // begin is 'Saved'
            string bookSavedTitleLoadDone = "//*[@data-unique-id='DocumentTitleSaveStatus']";// old: @"//i[contains(@data-icon-name,'ArrowSync') and @aria-hidden='true']"; // begin is 'Saved'
            // Check if the status "Saving.../Saved" is displayed then wait for it disappeared
            if (IsElementPresent(By.XPath(bookSavedTitleLoadDone)))
            {
                WaitForElementVisible(timeoutInSeconds, By.XPath(bookSavedTitleLoadDone));
                WaitForElementInvisible(timeoutInSeconds, By.XPath(bookSavedTitleLoadDone));
            }
            return this;
        }

        public FADAddInAction WaitforExcelWebSinglePrivateReportRenderDone(int timeoutInSeconds)
        {
            // This will shift focus back to main (default) content in which frame 'one' lies
            Driver.Browser.SwitchTo().DefaultContent();

            // Switch to the frame of Office Document Body (to interact with menu/button of MS Excel Online)
            NavigationAction.Instance.SwitchToFrameWithWaitMethod(timeoutInSeconds, NavigationPage.frameOfficeExcelOnline);

            // Wait for the title of Workbook (Saving ...) is disappeared
            //string bookSavedTitleLoadDone = @"//span[@data-unique-id='DocumentTitleSaveStatus' and .='Saving...']"; // begin is 'Saved'
            string bookSavedTitleLoadDone = @"//div[contains(@aria-label,'Saving... ')]"; // begin is 'Saved'
            WaitForElementVisible(timeoutInSeconds, By.XPath(bookSavedTitleLoadDone));
            WaitForElementInvisible(timeoutInSeconds, By.XPath(bookSavedTitleLoadDone));

            // This will shift focus back to main (default) content in which frame 'one' lies
            Driver.Browser.SwitchTo().DefaultContent();

            // Switch to the frame of Office Document Body (to interact with menu/button of MS Excel Online)
            NavigationAction.Instance.SwitchToFrameWithWaitMethod(timeoutInSeconds, NavigationPage.frameOfficeExcelOnline);

            // Wait for Sheet "Single Manager Dashboard Output" move to the last sheet 
            string singlePrivateSheet = @"//ul[@class='ewa-stb-tabs']/li[last()]/a[@aria-label='Single Manager Dashboard Output']";
            WaitForElementVisible(timeoutInSeconds, By.XPath(singlePrivateSheet));

            return this;
        }
        #endregion
    }
}
