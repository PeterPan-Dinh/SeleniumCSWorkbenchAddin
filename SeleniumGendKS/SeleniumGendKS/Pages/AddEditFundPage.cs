using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using SeleniumGendKS.Core.BaseClass;
using SeleniumGendKS.Core.Selenium;
using System;
using System.Xml.Linq;

namespace SeleniumGendKS.Pages
{
    internal class AddEditFundPage : BasePageElementMap
    {
        // Initiate variables
        internal static WebDriverWait? wait;
        internal static string fundNameManager = "Fund Manager Name";
        internal static string firm = "Firm";
        internal static string subAssetClass= "Sub-Asset Class ";
        internal static string latestActualValue = "Latest Actual Value";
        internal static string businessCity = "Business City";
        internal static string businessState = "Business State";
        internal static string businessCountry = "Business Country";
        internal static string businessStreet = "Business Street";
        internal static string businessZip = "Business ZIP";
        internal static string businessContact = "Business Contact";
        internal static string businessEmail = "Business Email";
        internal static string businessPhone = "Business Phone";
        internal static string fundName = "Fund Name *";
        internal static string strategy = "Strategy";
        internal static string vintageYear = "Vintage Year";
        internal static string strategyHeadquarter = "Strategy Headquarter";
        internal static string assetClass = "Asset Class";
        internal static string investmentStage = "Investment Stage";
        internal static string industryFocus = "Industry Focus";
        internal static string geographicFocus = "Geographic Focus";
        internal static string fundSizeM = "Fund Size (M)";
        internal static string currency = "Currency";

        // Initiate the By objects for elements
        internal static By fundManagerInputTxt = By.XPath(@"//input[@role='searchbox' and not(@placeholder)]");
        internal static By searchLoadingSpinnerIcon = By.XPath(@"//i[contains(@class,'spinner')]");
        internal static By loadingIcon = By.XPath(@"//div[contains(@class,'progress-spinner')]");
        internal static By fundManagerReturnOfResultsDropdown(string fundManager) => By.XPath(@"//ul[@role='listbox']//div[contains(.,'" + fundManager + "')]/ancestor::li");
        internal static By fundManagerAddNewLink = By.XPath(@"//span[.='Add New']");
        internal static By managerNameInputTxt = By.XPath(@"//input[@formcontrolname='manager_name']");
        internal static By fundNameInputTxt = By.XPath(@"//input[@formcontrolname='fund_name']");
        internal static By subAssetClassDropdown = By.XPath(@"//p-dropdown[@formcontrolname='sub_asset_class']");
        internal static By searchDropdownInputTxt = By.XPath(@"//input[contains(@class,'filter p-inputtext')]");
        internal static By dropdownSelect(string item) => By.XPath(@"//li[@aria-label='" + item + "']");
        internal static By assetClassInputTxt = By.XPath(@"//input[@formcontrolname='asset_class']");
        internal static By inceptionDateInputTxt = By.XPath(@"//span[contains(@class,'p-calendar')]/input");
        internal static By inceptionDateButton = By.XPath(@"//button[contains(@class,'datepicker')]");
        internal static By inceptionDateDayButton(string day) => By.XPath(@"//span[.='" + day + "']");
        internal static By latestActualValueInputTxt = By.XPath(@"//input[@formcontrolname='latest_actual_value']");
        internal static By businessStreetInputTxt = By.XPath(@"//input[@formcontrolname='business_street']");
        internal static By businessCityInputTxt = By.XPath(@"//input[@formcontrolname='business_city']");
        internal static By businessStateInputTxt = By.XPath(@"//input[@formcontrolname='business_state']");
        internal static By businessZipInputTxt = By.XPath(@"//input[@formcontrolname='business_zip']");
        internal static By businessCountryDropdown = By.XPath(@"//p-dropdown[@formcontrolname='business_country']");
        internal static By overlayDropdown = By.XPath(@"//div[contains(@class, 'overlayAnimation')]");
        internal static By cancelButton = By.XPath(@"//button[@label='Cancel']");
        internal static By saveButton = By.XPath(@"//button[@label='Save']");
        internal static By managerNameExistsErrorMessage = By.XPath(@"//div[contains(@class,'sm:col-6 ng-star-inserted')]/span");
        internal static By fundNameExistsErrorMessage = By.XPath(@"//div[contains(@class,'md:col-3 lg:col-2')]/span");

        /// Liquidity Section (By objects)
        internal static By redemptionFrequencyDropdown = By.XPath(@"//p-dropdown[@formcontrolname='data_redemption_frequency']");
        internal static By redemptionNoticePeriodInputTxt = By.XPath(@"//input[@formcontrolname='data_redemption_notice_days']");
        internal static By fieldCheckbox(string label) => By.XPath(@"//label[.='" + label + "']");
        internal static By hardLockMonthsInputTxt = By.XPath(@"//input[@formcontrolname='data_hard_lockup_months']");
        internal static By softLockMonthsInputTxt = By.XPath(@"//input[@formcontrolname='data_soft_lockup_months']");
        internal static By earlierRedemptionFeeInputTxt = By.XPath(@"//input[@formcontrolname='data_redemption_fee']");
        internal static By gateInputTxt = By.XPath(@"//input[@formcontrolname='data_redemption_gate']");
        internal static By gatePercInputNumber = By.XPath(@"//input[@formcontrolname='data_redemption_gate_percent']");
        internal static By percentOfNAVAvailableInputNumber = By.XPath(@"//input[@formcontrolname='data_percent_of_nav_available']");
        internal static By additionalNotesOnLiquidityInputTxt = By.XPath(@"//textarea[@formcontrolname='data_redemption_note']");

        /// Fees Section (By objects)
        internal static By managementFeeInputNumber = By.XPath(@"//input[@formcontrolname='data_management_fee']");
        internal static By managementFeePaidDropdown = By.XPath(@"//p-dropdown[@formcontrolname='data_management_fee_frequency']");
        internal static By performanceFeeInputNumber = By.XPath(@"//input[@formcontrolname='data_performance_fee']");
        internal static By catchUpDropdown = By.XPath(@"//p-dropdown[@formcontrolname='data_catch_up']");
        internal static By catchUpPercAgeIfSoftInputNumber = By.XPath(@"//input[@formcontrolname='data_catch_up_rate']");
        internal static By crystallizationEveryXYearsDropdown = By.XPath(@"//p-dropdown[@formcontrolname='data_crystalization_frequency']");
        internal static By hurdleFixedOrRelativeDropdown = By.XPath(@"//p-dropdown[@formcontrolname='data_hurdle_type']");
        internal static By hurdleRatePercInputTxt = By.XPath(@"//input[@formcontrolname='data_hurdle_rate']");
        internal static By hurdleBenchmarkSearchInputTxt = By.XPath(@"//label[.='Hurdle Benchmark']/preceding-sibling::p-autocomplete//input");
        internal static By hurdleBenchmarkReturnOfResults(string hurdleBenchmark) => By.XPath(@"//span[.='" + hurdleBenchmark + "']");
        internal static By benchMarkIDReadOnlyField = By.XPath(@"//label[.='Benchmark ID']/preceding-sibling::input");
        internal static By hurdleTypeDropdown = By.XPath(@"//p-dropdown[@formcontrolname='data_hurdle_hard_soft']");
        internal static By rampTypeDropdown = By.XPath(@"//p-dropdown[@formcontrolname='data_hurdle_ramp_type']");

        /// <summary>
        /// Private Fund (Add Manual)
        /// </summary>
        internal static By fundTypeDropdown = By.XPath(@"//div[@class='section-title' and contains(.,'Fund Type')]/following-sibling::div//p-dropdown");
        internal static By labelDropdown(string label) => By.XPath(@"//label[contains(.,'" + label + "')]/preceding-sibling::p-dropdown");
        internal static By labelInputTxt(string label) => By.XPath(@"//label[contains(.,'" + label + "')]/preceding-sibling::input");
        internal static By labelFundInputTxt(int number, string label) => By.XPath(@"//div[.=' Fund " + number + " ']/following-sibling::div[1]//label[.='" + label + "']/preceding-sibling::input");
        internal static By deleteFundNameButton(int number) => By.XPath(@"//div[.=' Fund " + number + " ']//span");
        internal static By addFundButton = By.XPath(@"//button[@label='Add Fund']");

        // Initiate the elements
        public IWebElement inputTxtFundManager => Driver.Browser.FindElement(fundManagerInputTxt);
        public IWebElement dropdownReturnOfResultsFundManager(int timeoutInSeconds, string fundManager)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(fundManagerReturnOfResultsDropdown(fundManager)));
        }
        public IWebElement addNewLinkFundManager(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(fundManagerAddNewLink));
        }
        public IWebElement inputTxtManagerName => Driver.Browser.FindElement(managerNameInputTxt);
        public IWebElement inputTxtFundName => Driver.Browser.FindElement(fundNameInputTxt);
        public IWebElement dropdownSubAssetClass => Driver.Browser.FindElement(subAssetClassDropdown);
        public IWebElement inputTxtSearchDropdown(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(d => d.FindElement(searchDropdownInputTxt));
        }
        public IWebElement selectDropdown(int timeoutInSeconds, string item)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(dropdownSelect(item)));
        }
        public IWebElement inputTxtAssetClass(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(d => d.FindElement(assetClassInputTxt));
        }
        public IWebElement inputTxtInceptionDate => Driver.Browser.FindElement(inceptionDateInputTxt);
        public IWebElement buttonInceptionDate(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(inceptionDateButton));
        }
        public IWebElement buttonInceptionDateDay(int timeoutInSeconds, string day)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(inceptionDateDayButton(day)));
        }
        public IWebElement inputTxtLatestActualValue => Driver.Browser.FindElement(latestActualValueInputTxt);
        public IWebElement inputTxtBusinessStreet => Driver.Browser.FindElement(businessStreetInputTxt);
        public IWebElement inputTxtBusinessCity => Driver.Browser.FindElement(businessCityInputTxt);
        public IWebElement inputTxtBusinessState => Driver.Browser.FindElement(businessStateInputTxt);
        public IWebElement inputTxtBusinessZIP => Driver.Browser.FindElement(businessZipInputTxt);
        public IWebElement dropdownBusinessCountry => Driver.Browser.FindElement(businessCountryDropdown);
        public IWebElement buttonCancel => Driver.Browser.FindElement(cancelButton);
        public IWebElement buttonSave => Driver.Browser.FindElement(saveButton);
        public IWebElement errorMessageManagerNameExists(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(managerNameExistsErrorMessage));
        }
        public IWebElement errorMessageFundNameExists(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(fundNameExistsErrorMessage));
        }

        /// Lquidity Section (IwebElemnt)
        public IWebElement dropdownRedemptionFrequency => Driver.Browser.FindElement(redemptionFrequencyDropdown);
        public IWebElement inputTxtRedemptionNoticePeriod => Driver.Browser.FindElement(redemptionNoticePeriodInputTxt);
        public IWebElement inputTxtHardLockMonths => Driver.Browser.FindElement(hardLockMonthsInputTxt);
        public IWebElement inputTxtSoftLockMonths => Driver.Browser.FindElement(softLockMonthsInputTxt);
        public IWebElement inputTxtEarlierRedemptionFee => Driver.Browser.FindElement(earlierRedemptionFeeInputTxt);
        public IWebElement inputTxtGate => Driver.Browser.FindElement(gateInputTxt);
        public IWebElement inputNumberGatePerc => Driver.Browser.FindElement(gatePercInputNumber);
        public IWebElement inputNumberPercentOfNAVAvailable => Driver.Browser.FindElement(percentOfNAVAvailableInputNumber);
        public IWebElement inputTxtAdditionalNotesOnLiquidity => Driver.Browser.FindElement(additionalNotesOnLiquidityInputTxt);

        /// Fees Section (IwebElemnt)
        public IWebElement inputNumberManagementFee => Driver.Browser.FindElement(managementFeeInputNumber);
        public IWebElement dropdownManagementFeePaid => Driver.Browser.FindElement(managementFeePaidDropdown);
        public IWebElement inputNumberPerformanceFee => Driver.Browser.FindElement(performanceFeeInputNumber);
        public IWebElement dropdownCatchUp => Driver.Browser.FindElement(catchUpDropdown);
        public IWebElement inputNumberCatchUpPercAgeIfSoft => Driver.Browser.FindElement(catchUpPercAgeIfSoftInputNumber);
        public IWebElement dropdownCrystallizationEveryXYears => Driver.Browser.FindElement(crystallizationEveryXYearsDropdown);
        public IWebElement checkboxField(string label) => Driver.Browser.FindElement(fieldCheckbox(label));
        public IWebElement dropdownHurdleFixedOrRelative => Driver.Browser.FindElement(hurdleFixedOrRelativeDropdown);
        public IWebElement inputTxtHurdleRatePerc => Driver.Browser.FindElement(hurdleRatePercInputTxt);
        public IWebElement inputTxtHurdleBenchmarkSearch => Driver.Browser.FindElement(hurdleBenchmarkSearchInputTxt);
        public IWebElement returnOfResultsHurdleBenchmark(int timeoutInSeconds, string hurdleBenchmark)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(hurdleBenchmarkReturnOfResults(hurdleBenchmark)));
        }
        public IWebElement readOnlyFieldBenchMarkID(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(benchMarkIDReadOnlyField));
        }
        public IWebElement dropdownHurdleType(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(hurdleTypeDropdown));
        }
        public IWebElement dropdownRampType(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(rampTypeDropdown));
        }

        /// Private Fund (Add Manual)
        public IWebElement dropdownFundType(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(fundTypeDropdown));
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
        public IWebElement inputLabelTxt(string label) => Driver.Browser.FindElement(labelInputTxt(label));
        public IWebElement inputFundLabelTxt(int number, string label) => Driver.Browser.FindElement(labelFundInputTxt(number, label));
        public IWebElement buttonFundNameDelete(int timeoutInSeconds, int number)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(deleteFundNameButton(number)));
        }
        public IWebElement buttonAddFund(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(addFundButton));
        }
    }

    internal sealed class AddEditFundAction : BasePage<AddEditFundAction, AddEditFundPage>
    {
        #region constructor
        private AddEditFundAction() { }
        #endregion

        #region Items action
        // Wait for loading Spinner icon to disappear
        public AddEditFundAction WaitForLoadingIconToDisappear(int timeoutInSeconds, By element)
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
        public AddEditFundAction WaitForElementVisible(int timeoutInSeconds, By element)
        {
            if (timeoutInSeconds > 0)
            {
                WebDriverWait wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
                wait.Until(ExpectedConditions.ElementIsVisible(element));
            }
            return this;
        }

        // Wait for all of elements visible (use for dropdown on-overlay visible)
        public AddEditFundAction WaitForAllElementsVisible(int timeoutInSeconds, By element)
        {
            if (timeoutInSeconds > 0)
            {
                WebDriverWait wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
                wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(element));
            }
            return this;
        }

        // Wait for element Invisible (can use for dropdown on-overlay Invisible)
        public AddEditFundAction WaitForElementInvisible(int timeoutInSeconds, By element)
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

        // scroll to element with JavaScript
        public AddEditFundAction ScrollIntoView(IWebElement iwebE)
        {
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].scrollIntoView(false);", iwebE);
            return this;
        }

        // Fund & Manager Information Section (Items action)
        public AddEditFundAction InputFundManagerToSearchFund(string fundManager)
        {
            this.Map.inputTxtFundManager.Clear();
            this.Map.inputTxtFundManager.SendKeys(fundManager);
            return this;
        }
        public AddEditFundAction ClickFundManagerDropdownReturnOfResults(int timeoutInSeconds, string fundManager)
        {
            Actions actions = new Actions(Driver.Browser);
            actions.MoveToElement(this.Map.dropdownReturnOfResultsFundManager(timeoutInSeconds, fundManager));
            actions.Perform();
            this.Map.dropdownReturnOfResultsFundManager(timeoutInSeconds, fundManager).Click();
            return this;
        }
        public AddEditFundAction ClickFundManagerAddNewLink(int timeoutInSeconds)
        {
            this.Map.addNewLinkFundManager(timeoutInSeconds).Click();
            return this;
        }
        public AddEditFundAction InputManagerName(string managerName)
        {
            this.Map.inputTxtManagerName.Clear();
            this.Map.inputTxtManagerName.SendKeys(managerName);
            return this;
        }
        public AddEditFundAction InputFundName(string fundName)
        {
            this.Map.inputTxtFundName.Clear();
            this.Map.inputTxtFundName.SendKeys(fundName);
            return this;
        }
        public AddEditFundAction ClickSubAssetClassDropdown()
        {
            this.Map.dropdownSubAssetClass.Click();
            return this;
        }
        public AddEditFundAction InputSearchDropdown(int timeoutInSeconds, string text)
        {
            this.Map.inputTxtSearchDropdown(timeoutInSeconds).Clear();
            this.Map.inputTxtSearchDropdown(timeoutInSeconds).SendKeys(text);
            return this;
        }
        public AddEditFundAction ClickSelectDropdown(int timeoutInSeconds, string item)
        {
            //this.Map.selectDropdown(timeoutInSeconds, item).Click();  --> element click intercepted
            ScrollIntoView(this.Map.selectDropdown(timeoutInSeconds, item));
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].click();", this.Map.selectDropdown(timeoutInSeconds, item));
            return this;
        }
        public AddEditFundAction InputAssetClass(int timeoutInSeconds, string assetClass)
        {
            this.Map.inputTxtAssetClass(timeoutInSeconds).SendKeys(assetClass);
            return this;
        }
        public AddEditFundAction InputInceptionDate(int timeoutInSeconds, string date)
        {
            this.Map.inputTxtInceptionDate.SendKeys(OpenQA.Selenium.Keys.Control + "a");
            this.Map.inputTxtInceptionDate.SendKeys(date);
            //this.Map.buttonInceptionDate(timeoutInSeconds).Click(); --> Element Click Intercepted Exception

            // Try with javascript if Element Click Intercepted Exception
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].click();", this.Map.buttonInceptionDate(timeoutInSeconds));
            WaitForElementInvisible(10, AddEditFundPage.overlayDropdown);
            return this;
        }
        //public AddEditFundAction InputInceptionDateForAddNewFund(int timeoutInSeconds, string day)
        //{
        //    this.Map.inputTxtInceptionDate.Click();
        //    this.Map.buttonInceptionDateDay(timeoutInSeconds, day).Click();
        //    WaitForElementInvisible(10, AddEditFundPage.overlayDropdown);
        //    return this;
        //}
        public AddEditFundAction InputLatestActualValue(string latestActualValue)
        {
            this.Map.inputTxtLatestActualValue.Clear();
            this.Map.inputTxtLatestActualValue.SendKeys(latestActualValue);
            return this;
        }
        public AddEditFundAction InputBusinessStreet(string text)
        {
            this.Map.inputTxtBusinessStreet.Clear();
            this.Map.inputTxtBusinessStreet.SendKeys(text);
            return this;
        }
        public AddEditFundAction InputBusinessCity(string text)
        {
            this.Map.inputTxtBusinessCity.Clear();
            this.Map.inputTxtBusinessCity.SendKeys(text);
            return this;
        }
        public AddEditFundAction InputBusinessState(string text)
        {
            this.Map.inputTxtBusinessState.Clear();
            this.Map.inputTxtBusinessState.SendKeys(text);
            return this;
        }
        public AddEditFundAction InputBusinessZIP(string text)
        {
            this.Map.inputTxtBusinessZIP.Clear();
            this.Map.inputTxtBusinessZIP.SendKeys(text);
            return this;
        }
        public AddEditFundAction ClickBusinessCountryDropdown()
        {
            this.Map.dropdownBusinessCountry.Click();
            return this;
        }
        public AddEditFundAction ClickCancelButton()
        {
            this.Map.buttonCancel.Click();
            return this;
        }
        public AddEditFundAction ClickSaveButton()
        {
            this.Map.buttonSave.Click();
            return this;
        }
        public string ErrorMessageManagerNameExistsGetText(int timeoutInSeconds)
        {
            return this.Map.errorMessageManagerNameExists(timeoutInSeconds).Text;
        }
        public string ErrorMessageFundNameExistsGetText(int timeoutInSeconds)
        {
            return this.Map.errorMessageFundNameExists(timeoutInSeconds).Text;
        }

        // Liquidity Section (Items action)
        public AddEditFundAction ClickRedemptionFrequencyDropdown()
        {
            this.Map.dropdownRedemptionFrequency.Click();
            return this;
        }
        public AddEditFundAction InputRedemptionNoticePeriod(string text)
        {
            this.Map.inputTxtRedemptionNoticePeriod.Clear();
            this.Map.inputTxtRedemptionNoticePeriod.SendKeys(text);
            return this;
        }
        public AddEditFundAction InputHardLockMonths(string text)
        {
            this.Map.inputTxtHardLockMonths.Clear();
            this.Map.inputTxtHardLockMonths.SendKeys(text);
            return this;
        }
        public AddEditFundAction InputSoftLockMonths(string text)
        {
            this.Map.inputTxtSoftLockMonths.Clear();
            this.Map.inputTxtSoftLockMonths.SendKeys(text);
            return this;
        }
        public AddEditFundAction InputEarlierRedemptionFee(string text)
        {
            this.Map.inputTxtEarlierRedemptionFee.Clear();
            this.Map.inputTxtEarlierRedemptionFee.SendKeys(text);
            return this;
        }
        public AddEditFundAction InputGate(string text)
        {
            this.Map.inputTxtGate.Clear();
            this.Map.inputTxtGate.SendKeys(text);
            return this;
        }
        public AddEditFundAction InputGatePerc(string text)
        {
            this.Map.inputNumberGatePerc.Clear();
            this.Map.inputNumberGatePerc.SendKeys(text);
            return this;
        }
        public AddEditFundAction InputPercentOfNAVAvailable(string text)
        {
            this.Map.inputNumberPercentOfNAVAvailable.Clear();
            this.Map.inputNumberPercentOfNAVAvailable.SendKeys(text);
            return this;
        }
        public AddEditFundAction InputAdditionalNotesOnLiquidity(string text)
        {
            this.Map.inputTxtAdditionalNotesOnLiquidity.Clear();
            this.Map.inputTxtAdditionalNotesOnLiquidity.SendKeys(text);
            return this;
        }

        // Fees Section (Items action)
        public AddEditFundAction InputManagementFee(string text)
        {
            this.Map.inputNumberManagementFee.Clear();
            this.Map.inputNumberManagementFee.SendKeys(text);
            return this;
        }
        public AddEditFundAction ClickManagementFeePaidDropdown()
        {
            this.Map.dropdownManagementFeePaid.Click();
            return this;
        }
        public AddEditFundAction InputPerformanceFee(string text)
        {
            this.Map.inputNumberPerformanceFee.Clear();
            this.Map.inputNumberPerformanceFee.SendKeys(text);
            return this;
        }
        public AddEditFundAction ClickCatchUpDropdown()
        {
            this.Map.dropdownCatchUp.Click();
            return this;
        }
        public AddEditFundAction InputCatchUpPercAgeIfSoft(string text)
        {
            this.Map.inputNumberCatchUpPercAgeIfSoft.Clear();
            this.Map.inputNumberCatchUpPercAgeIfSoft.SendKeys(text);
            return this;
        }
        public AddEditFundAction ClickCrystallizationEveryXYearsDropdown()
        {
            this.Map.dropdownCrystallizationEveryXYears.Click();
            return this;
        }
        public AddEditFundAction ClickFieldCheckbox(string label)
        {
            this.Map.checkboxField(label).Click();
            return this;
        }
        public AddEditFundAction ClickHurdleFixedOrRelativeDropdown()
        {
            this.Map.dropdownHurdleFixedOrRelative.Click();
            return this;
        }
        public AddEditFundAction InputHurdleRatePerc(string text)
        {
            this.Map.inputTxtHurdleRatePerc.Clear();
            this.Map.inputTxtHurdleRatePerc.SendKeys(text);
            return this;
        }
        public AddEditFundAction InputHurdleBenchmarkToSearch(string text)
        {
            this.Map.inputTxtHurdleBenchmarkSearch.Clear();
            this.Map.inputTxtHurdleBenchmarkSearch.SendKeys(text);
            return this;
        }
        public AddEditFundAction ClickHurdleBenchmarkReturnOfResults(int timeout, string hurdleBenchmark)
        {
            Actions actions = new Actions(Driver.Browser);
            actions.MoveToElement(this.Map.returnOfResultsHurdleBenchmark(timeout, hurdleBenchmark));
            actions.Perform();
            this.Map.returnOfResultsHurdleBenchmark(timeout, hurdleBenchmark).Click();
            return this;
        }
        public string BenchMarkIDReadOnlyFieldGetText(int timeoutInSeconds)
        {
            return this.Map.readOnlyFieldBenchMarkID(timeoutInSeconds).Text;
        }
        public AddEditFundAction ClickHurdleTypeDropdown(int timeoutInSeconds)
        {
            this.Map.dropdownHurdleType(timeoutInSeconds).Click();
            return this;
        }
        public AddEditFundAction ClickRampTypeDropdown(int timeoutInSeconds)
        {
            this.Map.dropdownRampType(timeoutInSeconds).Click();
            return this;
        }

        // Private Fund (Add Manual)
        public AddEditFundAction ClickFundTypeDropdown(int timeoutInSeconds)
        {
            this.Map.dropdownFundType(timeoutInSeconds).Click();
            return this;
        }
        public AddEditFundAction ClickLabelDropdown(int timeoutInSeconds, string label)
        {
            this.Map.dropdownLabel(timeoutInSeconds, label).Click();
            return this;
        }
        public AddEditFundAction SelectItemInDropdown(int timeoutInSeconds, string item)
        {
            this.Map.selectItemDropdown(timeoutInSeconds, item).Click();
            return this;
        }
        public AddEditFundAction InputTxtLabel(string label, string txt)
        {
            this.Map.inputLabelTxt(label).Clear();
            this.Map.inputLabelTxt(label).SendKeys(txt);
            return this;
        }
        public AddEditFundAction InputTxtFundLabel(int fundNumber, string label, string txt)
        {
            this.Map.inputFundLabelTxt(fundNumber, label).Clear();
            this.Map.inputFundLabelTxt(fundNumber, label).SendKeys(txt);
            return this;
        }
        public AddEditFundAction ClickDeleteFundNameButton(int timeoutInSeconds, int number)
        {
            Actions actions = new Actions(Driver.Browser);
            actions.MoveToElement(this.Map.buttonFundNameDelete(timeoutInSeconds, number));
            actions.Perform();
            this.Map.buttonFundNameDelete(timeoutInSeconds, number).Click();
            return this;
        }
        public AddEditFundAction ClickAddFundButton(int timeoutInSeconds)
        {
            Actions actions = new Actions(Driver.Browser);
            actions.MoveToElement(this.Map.buttonAddFund(timeoutInSeconds));
            actions.Perform();
            this.Map.buttonAddFund(timeoutInSeconds).Click();
            return this;
        }
        #endregion

        #region Built-in actions
        // Fund & Manager Information Section (Built-in actions)
        public AddEditFundAction ClickAndSearchSubAssetClassDropDown(int timeoutInSeconds, string subAssetClass)
        {
            ClickSubAssetClassDropdown();
            InputSearchDropdown(timeoutInSeconds, subAssetClass);
            ClickSelectDropdown(timeoutInSeconds, subAssetClass);
            return this;
        }

        public AddEditFundAction ClickAndSelectSubAssetClassDropDown(int timeoutInSeconds, string subAssetClass)
        {
            ClickSubAssetClassDropdown();
            WaitForAllElementsVisible(10, AddEditFundPage.overlayDropdown);
            WaitForAllElementsVisible(10, AddEditFundPage.dropdownSelect(subAssetClass));
            ClickSelectDropdown(timeoutInSeconds, subAssetClass);
            WaitForElementInvisible(10, AddEditFundPage.overlayDropdown);
            return this;
        }

        public AddEditFundAction ClickAndSearchBusinessCountryDropDown(int timeoutInSeconds, string businessCountry)
        {
            ClickBusinessCountryDropdown();
            InputSearchDropdown(timeoutInSeconds, businessCountry);
            ClickSelectDropdown(timeoutInSeconds, businessCountry);
            return this;
        }

        public AddEditFundAction ClickAndSelectBusinessCountryDropDown(int timeoutInSeconds, string businessCountry)
        {
            ClickBusinessCountryDropdown();
            WaitForAllElementsVisible(10, AddEditFundPage.overlayDropdown);
            WaitForAllElementsVisible(10, AddEditFundPage.dropdownSelect(businessCountry));
            ClickSelectDropdown(timeoutInSeconds, businessCountry);
            WaitForElementInvisible(10, AddEditFundPage.overlayDropdown);
            return this;
        }

        // Liquidity Section (Built-in actions)
        public AddEditFundAction ClickAndSelectRedemptionFrequencyDropDown(int timeoutInSeconds, string redemptionFrequency)
        {
            ClickRedemptionFrequencyDropdown();
            WaitForAllElementsVisible(10, AddEditFundPage.overlayDropdown);
            WaitForAllElementsVisible(10, AddEditFundPage.dropdownSelect(redemptionFrequency));
            ClickSelectDropdown(timeoutInSeconds, redemptionFrequency);
            WaitForElementInvisible(10, AddEditFundPage.overlayDropdown);
            return this;
        }

        // Fees Section (Built-in actions)
        public AddEditFundAction ClickAndSelectManagementFeePaidDropDown(int timeoutInSeconds, string managementFeePaid)
        {
            ClickManagementFeePaidDropdown();
            WaitForAllElementsVisible(10, AddEditFundPage.overlayDropdown);
            WaitForAllElementsVisible(10, AddEditFundPage.dropdownSelect(managementFeePaid));
            ClickSelectDropdown(timeoutInSeconds, managementFeePaid);
            WaitForElementInvisible(10, AddEditFundPage.overlayDropdown);
            return this;
        }
        public AddEditFundAction ClickAndSelectCatchUpDropDown(int timeoutInSeconds, string catchUp)
        {
            ClickCatchUpDropdown();
            WaitForAllElementsVisible(10, AddEditFundPage.overlayDropdown);
            WaitForAllElementsVisible(10, AddEditFundPage.dropdownSelect(catchUp));
            ClickSelectDropdown(timeoutInSeconds, catchUp);
            WaitForElementInvisible(10, AddEditFundPage.overlayDropdown);
            return this;
        }
        public AddEditFundAction ClickAndSelectCrystallizationEveryXYearsDropDown(int timeoutInSeconds, string crystalizationFrequency)
        {
            ClickCrystallizationEveryXYearsDropdown();
            WaitForAllElementsVisible(10, AddEditFundPage.overlayDropdown);
            WaitForAllElementsVisible(10, AddEditFundPage.dropdownSelect(crystalizationFrequency));
            ClickSelectDropdown(timeoutInSeconds, crystalizationFrequency);
            WaitForElementInvisible(10, AddEditFundPage.overlayDropdown);
            return this;
        }
        public AddEditFundAction ClickAndSelectHurdleFixedOrRelativeDropDown(int timeoutInSeconds, string hurdleFixedOrRelative)
        {
            ClickHurdleFixedOrRelativeDropdown();
            WaitForAllElementsVisible(10, AddEditFundPage.overlayDropdown);
            WaitForAllElementsVisible(10, AddEditFundPage.dropdownSelect(hurdleFixedOrRelative));
            ClickSelectDropdown(timeoutInSeconds, hurdleFixedOrRelative);
            WaitForElementInvisible(10, AddEditFundPage.overlayDropdown);
            return this;
        }
        public AddEditFundAction ClickAndSelectHurdleTypeDropDown(int timeoutInSeconds, string hurdleType)
        {
            ClickHurdleTypeDropdown(timeoutInSeconds);
            WaitForAllElementsVisible(10, AddEditFundPage.overlayDropdown);
            WaitForAllElementsVisible(10, AddEditFundPage.dropdownSelect(hurdleType));
            ClickSelectDropdown(timeoutInSeconds, hurdleType);
            WaitForElementInvisible(10, AddEditFundPage.overlayDropdown);
            return this;
        }
        public AddEditFundAction ClickAndSelectRampTypeDropDown(int timeoutInSeconds, string rampType)
        {
            ClickRampTypeDropdown(timeoutInSeconds);
            WaitForAllElementsVisible(10, AddEditFundPage.overlayDropdown);
            WaitForAllElementsVisible(10, AddEditFundPage.dropdownSelect(rampType));
            ClickSelectDropdown(timeoutInSeconds, rampType);
            WaitForElementInvisible(10, AddEditFundPage.overlayDropdown);
            return this;
        }
        public AddEditFundAction ClickToUnCheckTheCheckbox(string label)
        {
            bool isChecked = false;
            isChecked = Driver.Browser.FindElement(AddEditFundPage.fieldCheckbox(label)).GetAttribute("class").Contains("active");
            if (isChecked)
            {
                ClickFieldCheckbox(label);
            }
            return this;
        }
        public AddEditFundAction ClickToCheckTheCheckbox(string label)
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
        public AddEditFundAction ClickAndSelectItemInDropdown(int timeoutInSeconds, string label, string item)
        {
            ScrollIntoView(this.Map.dropdownLabel(timeoutInSeconds, label));
            ClickLabelDropdown(timeoutInSeconds, label);
            WaitForAllElementsVisible(10, AddEditFundPage.overlayDropdown);
            WaitForAllElementsVisible(10, AddEditFundPage.dropdownSelect(item));
            SelectItemInDropdown(timeoutInSeconds, item);
            WaitForElementInvisible(10, AddEditFundPage.overlayDropdown);
            return this;
        }
        #endregion
    }
}
