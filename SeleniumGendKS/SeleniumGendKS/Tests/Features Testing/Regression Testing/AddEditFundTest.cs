using AventStack.ExtentReports;
using NUnit.Framework;
using SeleniumGendKS.Core.BaseTestCase;
using SeleniumGendKS.Pages;
using System;
using System.Windows.Forms;

namespace SeleniumGendKS.Tests.Features_Testing.Regression_Testing
{
    [TestFixture, Order(3)]
    internal class AddEditFundTest : BaseTestCase
    {
        // Variables declare
        [Obsolete]
        readonly ExtentReports rep = ExtReportgetInstance();
        private string? summaryTC;

        [Test, Category("Regression Testing")]
        public void TC001_Add_FundManual_AddFundManagerExists()
        {
            #region Variables declare
            string urlInstance = LoginPage.url;
            const string managerNameExistsMessage = "Manager name already exists";
            const string fundNameExistsMessage = "Fund name already exists in the manager";
            string? managerNameManual = null;
            string? fundNameManual = null;
            string? fundNameManual2 = null;
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Add FundManual Test - TC001");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Log into the application
                LoginAction.Instance.LoginSiteNoGodaddy(urlInstance);

                // Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
                if (urlInstance.Contains("sandbox"))
                {
                    managerNameManual = "QA Test 05";
                    fundNameManual = "Main QA Test 05";
                    fundNameManual2 = "Child 01 of QA Test 05";
                }

                if (urlInstance.Contains("conceptia"))
                {
                    managerNameManual = "QA Test 01";
                    fundNameManual = "Main of QA Test 01";
                    fundNameManual2 = "Dinh Fund 01";
                }

                // Workaround --> when click cancel button does not work. (Input a text to 'Search Fund' field)
                FundDashboardAction.Instance.InputNameToSearchFund(10, "abcef").InputNameToSearchFund(10, "");

                // Click on plus (+) sign icon (to go to the Add New Fund page)
                FundDashboardAction.Instance.ClickAddNewFundIcon(10);

                // Add a new Fund
                AddEditFundAction.Instance.InputFundManagerToSearchFund(managerNameManual); // "QA Test"

                // Check if loading Spinner Icon in search textbox is shown
                System.Threading.Thread.Sleep(500);
                if (AddEditFundAction.Instance.IsElementPresent(AddEditFundPage.searchLoadingSpinnerIcon))
                { AddEditFundAction.Instance.WaitForLoadingIconToDisappear(20, AddEditFundPage.searchLoadingSpinnerIcon); }

                /// Add a new Fund                          
                AddEditFundAction.Instance.ClickFundManagerAddNewLink(10)
                                          .InputManagerName(managerNameManual)
                                          .InputFundName(fundNameManual)
                                          .ClickAndSelectSubAssetClassDropDown(10, "US Growth Equity")
                                          .InputInceptionDate(10, "03/29/2022")
                                          .InputLatestActualValue("QA auto Latest Actual Value")
                                          .InputBusinessStreet("QA auto Business Street")
                                          .InputBusinessCity("QA auto Business City")
                                          .InputBusinessState("QA auto Business State")
                                          .InputBusinessZIP("QA auto Business ZIP")
                                          .ClickAndSelectBusinessCountryDropDown(10, "Japan")
                                          .ClickSaveButton();

                // Check if loading Spinner Icon in search textbox is shown
                System.Threading.Thread.Sleep(500);
                if (AddEditFundAction.Instance.IsElementPresent(AddEditFundPage.searchLoadingSpinnerIcon))
                { AddEditFundAction.Instance.WaitForLoadingIconToDisappear(20, AddEditFundPage.searchLoadingSpinnerIcon); }

                // Verify the error message fund exists is shown correctly after clicking on save button
                verifyPoint = managerNameExistsMessage == AddEditFundAction.Instance.ErrorMessageManagerNameExistsGetText(10);
                verifyPoints.Add(summaryTC = "Verify the error message (Manager Name Exists) is shown correctly: " + managerNameExistsMessage + " ", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = fundNameExistsMessage == AddEditFundAction.Instance.ErrorMessageFundNameExistsGetText(10);
                verifyPoints.Add(summaryTC = "Verify the error message (Fund Name Exists) is shown correctly: " + fundNameExistsMessage + " ", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click on Cancel button (Back to the previous page)
                AddEditFundAction.Instance.ClickCancelButton()
                                          .WaitForElementVisible(10, FundDashboardPage.addNewFundIcon);

                // Click on plus (+) sign icon (to go to the Add New Fund page)
                FundDashboardAction.Instance.ClickAddNewFundIcon(10);

                // Search another Fund Manager is exists
                AddEditFundAction.Instance.InputFundManagerToSearchFund(managerNameManual)
                                          .WaitForLoadingIconToDisappear(20, AddEditFundPage.searchLoadingSpinnerIcon);

                AddEditFundAction.Instance.ClickFundManagerDropdownReturnOfResults(10, managerNameManual);

                AddEditFundAction.Instance.InputFundName(fundNameManual2)
                                          .ClickSaveButton()
                                          .WaitForLoadingIconToDisappear(20, AddEditFundPage.loadingIcon);

                verifyPoint = fundNameExistsMessage == AddEditFundAction.Instance.ErrorMessageFundNameExistsGetText(10);
                verifyPoints.Add(summaryTC = "Verify user is unable to add a fund name that is existing (after searching an existing Fund Manager successfully): " + fundNameExistsMessage + " ", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
            }
            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Warning
                ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong! Please check console log.");
            }
            #endregion
        }

        [Test, Category("Regression Testing")]
        public void TC002_Edit_FundManual_UpdateFundSuccessfully()
        {
            #region Variables declare
            string urlInstance = LoginPage.url;
            const string toastMessageUpdateSuccess = "Update fund successfully.";
            string? managerNameManual = null;
            string? fundNameManual = null;
            string? fundType = null;
            string? inceptionDate = null;
            string? fundStatus = null;
            string? latestActualValue = null;
            string? redemptionNotedays = null;
            string? hardLockmonths = null;
            string? softLockmonths = null;
            string? earlyRedemptionFee = null;
            string? gate = null;
            string? gatePerc = null;
            string? percentofNAVAvailable = null;
            string? sidepocket = null;
            string? additionalNotesOnLiquidity = null;
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Edit FundManual Test - TC002");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Log into the application
                LoginAction.Instance.LoginSiteNoGodaddy(urlInstance); // LoginSite(urlInstance);

                // Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
                if (urlInstance.Contains("sandbox"))
                {
                    managerNameManual = "QA Test 05";
                    fundNameManual = "Main QA Test 05";

                    // Fund Info
                    fundType = "";
                    inceptionDate = "03/29/2022";
                    fundStatus = "";
                    latestActualValue = "QA auto update Latest Actual Value";
                    redemptionNotedays = "60";
                    hardLockmonths = "QA auto update Hard Lock Months";
                    softLockmonths = "QA auto update Soft Lock Months";
                    earlyRedemptionFee = "QA auto update Earlier Redemption Fee";
                    gate = "No";
                    gatePerc = "33.3";
                    percentofNAVAvailable = "44.03";
                    sidepocket = "Yes";
                    additionalNotesOnLiquidity = "QA auto update Additional Notes On Liquidity";
                }

                if (urlInstance.Contains("conceptia"))
                {
                    managerNameManual = "QA Test 01";
                    fundNameManual = "Main of QA Test 01";

                    // Fund Info
                    fundType = "";
                    inceptionDate = @"03/29/2022";
                    fundStatus = "";
                    latestActualValue = "QA auto update Latest Actual Value";
                    redemptionNotedays = "60";
                    hardLockmonths = "QA auto update Hard Lock Months";
                    softLockmonths = "QA auto update Soft Lock Months";
                    earlyRedemptionFee = "QA auto update Earlier Redemption Fee";
                    gate = "No";
                    gatePerc = "33.3";
                    percentofNAVAvailable = "44.03";
                    sidepocket = "Yes";
                    additionalNotesOnLiquidity = "QA auto update Additional Notes On Liquidity";
                }

                // Search a Manual Fund
                FundDashboardAction.Instance.InputNameToSearchFund(10, managerNameManual)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(managerNameManual))
                                            .ClickFundNameReturnOfResults(10, managerNameManual)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(fundNameManual))
                                            .ClickFundNameReturnOfResults(10, fundNameManual)
                                            .WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);

                // Click on pencil icon (to go to the edit-Fund page)
                FundDashboardAction.Instance.ClickEditFundPencilIcon(10);

                // Edit Fund (Add new Fund Manager already exists)
                AddEditFundAction.Instance.InputFundName(fundNameManual + "_Update")
                                          .ClickAndSelectSubAssetClassDropDown(10, "US Buyout")
                                          .InputInceptionDate(10, inceptionDate)
                                          .InputLatestActualValue(latestActualValue)
                                          .InputBusinessStreet("QA auto update Business Street")
                                          .InputBusinessCity("QA auto update Business City")
                                          .InputBusinessState("QA auto update Business State")
                                          .InputBusinessZIP("QA auto update Business ZIP")
                                          .ClickAndSelectBusinessCountryDropDown(10, "United Kingdom")
                                          // Liquidity (section)
                                          .ClickAndSelectRedemptionFrequencyDropDown(10, "Quarterly")
                                          .InputRedemptionNoticePeriod(redemptionNotedays)
                                          .ClickToCheckTheCheckbox("Hard Lockup") // Checkbox
                                          .InputHardLockMonths(hardLockmonths)
                                          .InputSoftLockMonths(softLockmonths)
                                          .InputEarlierRedemptionFee(earlyRedemptionFee)
                                          .InputGate(gate)
                                          .InputGatePerc(gatePerc)
                                          .InputPercentOfNAVAvailable(percentofNAVAvailable)
                                          .ClickToCheckTheCheckbox("Sidepocket?") // Checkbox
                                          .InputAdditionalNotesOnLiquidity(additionalNotesOnLiquidity)
                                          // Fees (section)
                                          .InputManagementFee("55.05")
                                          .ClickAndSelectManagementFeePaidDropDown(10, "Quarterly")
                                          .InputPerformanceFee("66.06")
                                          .ClickToCheckTheCheckbox("High Water Mark") // Checkbox
                                          .ClickAndSelectCatchUpDropDown(10, "Yes")
                                          .InputCatchUpPercAgeIfSoft("77.07")
                                          .ClickAndSelectCrystallizationEveryXYearsDropDown(10, "3")
                                          .ClickToCheckTheCheckbox("Hurdle Status") // Checkbox
                                          .ClickAndSelectHurdleFixedOrRelativeDropDown(10, "Relative") // or Fixed
                                                                                                       //.InputHurdleRatePerc("88.08") --> Only Enable when Hurdle = Fixed
                                          .InputHurdleBenchmarkToSearch("msci");
                if (LoginAction.Instance.IsElementPresent(AddEditFundPage.searchLoadingSpinnerIcon))
                {
                    AddEditFundAction.Instance.WaitForLoadingIconToDisappear(20, AddEditFundPage.searchLoadingSpinnerIcon);
                }                     
                AddEditFundAction.Instance.ClickHurdleBenchmarkReturnOfResults(10, "MSCI Emerging Latin America Net TR Index (LCL)") // BenchMarkID = 892000 (old name: MSCI Emerging Latin America Net Total Return Local Index)
                                          .ClickAndSelectHurdleTypeDropDown(10, "Soft")
                                          .ClickAndSelectRampTypeDropDown(10, "Performance Dependent")
                                          .ClickSaveButton();

                // Verify all informations from Fund Info
                verifyPoint = toastMessageUpdateSuccess == FundDashboardAction.Instance.toastMessageAlertGetText(10, toastMessageUpdateSuccess);
                verifyPoints.Add(summaryTC = "Verify Fund Info - toast message (update successfully) is shown correctly after updating: '" + toastMessageUpdateSuccess + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Wait for toast message is disappeared
                AddEditFundAction.Instance.WaitForElementInvisible(10, FundDashboardPage.toastMessage(toastMessageUpdateSuccess));

                verifyPoint = fundNameManual + "_Update" == FundDashboardAction.Instance.FundNameTitleGetText(10);
                verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Name (Manual) is shown correctly after updating: '" + fundNameManual + "_Update'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = fundType == FundDashboardAction.Instance.FundInfoGetText(10, "Fund Type");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Type (Manual) is shown correctly after updating: '" + fundType + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = managerNameManual == FundDashboardAction.Instance.FundInfoGetText(10, "Manager Name");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Manager Name (Manual) is shown correctly after updating: '" + managerNameManual + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = inceptionDate == FundDashboardAction.Instance.FundInfoGetText(10, "Inception Date");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Inception Date (Manual) is shown correctly after updating: '" + inceptionDate + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = fundStatus == FundDashboardAction.Instance.FundInfoGetText(10, "Fund Status");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Status (Manual) is shown correctly after updating: '" + fundStatus + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = latestActualValue == FundDashboardAction.Instance.FundInfoGetText(10, "Latest Actual Value");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Latest Actual Value (Manual) is shown correctly after updating: '" + latestActualValue + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = redemptionNotedays == FundDashboardAction.Instance.FundInfoGetText(10, "Redemption Note (days)");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Redemption Note (days) (Manual) is shown correctly after updating: '" + redemptionNotedays + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = hardLockmonths == FundDashboardAction.Instance.FundInfoGetText(10, "Hard Lock (months)");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Hard Lock (months) (Manual) is shown correctly after updating: '" + hardLockmonths + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = softLockmonths == FundDashboardAction.Instance.FundInfoGetText(10, "Soft Lock (months)");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Soft Lock (months) (Manual) is shown correctly after updating: '" + softLockmonths + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = earlyRedemptionFee == FundDashboardAction.Instance.FundInfoGetText(10, "Early Redemption Fee");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Early Redemption Fee (Manual) is shown correctly after updating: '" + earlyRedemptionFee + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = gate == FundDashboardAction.Instance.FundInfoGetText(10, "Gate");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Gate (Manual) is shown correctly after updating: '" + gate + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = gatePerc == FundDashboardAction.Instance.FundInfoGetText(10, "Gate %");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Gate % (Manual) is shown correctly after updating: '" + gatePerc + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = percentofNAVAvailable == FundDashboardAction.Instance.FundInfoGetText(10, "Percent of NAV Available");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Percent of NAV Available (Manual) is shown correctly after updating: '" + percentofNAVAvailable + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = sidepocket == FundDashboardAction.Instance.FundInfoGetText(10, "Sidepocket?");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Sidepocket? (Manual) is shown correctly after updating: '" + sidepocket + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = additionalNotesOnLiquidity == FundDashboardAction.Instance.FundInfoGetText(10, "Additional Notes on Liquidity");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Additional Notes on Liquidity (Manual) is shown correctly after updating: '" + additionalNotesOnLiquidity + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Re-Update the Fund Name (back to the original Fund Name)
                // Click on pencil icon (to go to the edit-Fund page)
                FundDashboardAction.Instance.ClickEditFundPencilIcon(10);
                System.Threading.Thread.Sleep(1500);

                // Edit Fund (Add new Fund Manager already exists)
                AddEditFundAction.Instance.InputFundName(fundNameManual)
                                          .ClickSaveButton();

                // Wait for toast message is appeared
                AddEditFundAction.Instance.WaitForElementVisible(10, FundDashboardPage.toastMessage(toastMessageUpdateSuccess));
            }

            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Warning
                ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong! Please check console log.");
            }
            #endregion
        }

        [Test, Category("Regression Testing")]
        public void TC003_Edit_FundManual_AddFundManagerExists()
        {
            #region Variables declare
            string urlInstance = LoginPage.url;
            const string managerNameExistsMessage = "Manager name already exists";
            string? managerNameManual = null;
            string? fundNameManual = null;
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Edit FundManual Test - TC003");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Log into the application
                LoginAction.Instance.LoginSiteNoGodaddy(urlInstance); // LoginSite(urlInstance)

                // Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
                if (urlInstance.Contains("sandbox"))
                {
                    managerNameManual = "QA Test 05";
                    fundNameManual = "Main QA Test 05";
                }

                if (urlInstance.Contains("conceptia"))
                {
                    managerNameManual = "QA Test 01";
                    fundNameManual = "Main of QA Test 01";
                }

                // Search a Manual Fund
                FundDashboardAction.Instance.InputNameToSearchFund(10, managerNameManual)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(managerNameManual))
                                            .ClickFundNameReturnOfResults(10, managerNameManual)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(fundNameManual))
                                            .ClickFundNameReturnOfResults(10, fundNameManual)
                                            .WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);

                // Click on pencil icon (to go to the edit-Fund page)
                FundDashboardAction.Instance.ClickEditFundPencilIcon(10);

                // Edit Fund (Add new Fund Manager already exists)
                AddEditFundAction.Instance.InputFundManagerToSearchFund(managerNameManual)
                                          .WaitForLoadingIconToDisappear(20, AddEditFundPage.searchLoadingSpinnerIcon)
                                          .ClickFundManagerAddNewLink(10)
                                          .InputManagerName(managerNameManual)
                                          .ClickSaveButton();
                                          //.WaitForLoadingIconToDisappear(20, AddEditFundPage.loadingIcon);

                // Verify the error message fund exists is shown correctly after clicking on save button
                verifyPoint = managerNameExistsMessage == AddEditFundAction.Instance.ErrorMessageManagerNameExistsGetText(10);
                verifyPoints.Add(summaryTC = "Verify the error message (Manager Name Exists) is shown correctly: " + managerNameExistsMessage + " ", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
            }
            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Warning
                ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong! Please check console log.");
            }
            #endregion
        }

        [Test, Category("Regression Testing")]
        public void TC004_Add_FundManual_AddNewFund()
        {
            #region Variables declare
            string urlInstance = LoginPage.url;
            const string toastMessageAddSuccess = "Add new fund successfully.";
            string? managerNameManual = null;
            string? fundNameManual = null;
            string? inceptionDate = null;
            //string? inceptionDateDay = null;
            string? latestActualValue = null;
            string? fundType = null;
            string? fundStatus = null;
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Add FundManual Test - TC004");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Log into the application
                LoginAction.Instance.LoginSiteNoGodaddy(urlInstance); //LoginSite(urlInstance)

                // Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
                if (urlInstance.Contains("sandbox"))
                {
                    managerNameManual = "QA_Auto_Manager" + @"_" + DateTime.Now.ToString().Replace("/", "_").Replace(":", "_").Replace(" ", "_");
                    fundNameManual = "QA_Auto_Fund" + @"_" + DateTime.Now.ToString().Replace("/", "_").Replace(":", "_").Replace(" ", "_");
                    //inceptionDate = DateTime.Now.Date.ToString("MM/dd/yyyy");
                    string date = DateTime.Now.Date.ToString("MM/dd/yyyy");
                    inceptionDate = date.Replace("-", "/");
                    //inceptionDateDay = DateTime.Now.Day.ToString(); //.ToString("dd");
                    latestActualValue = "QA auto Latest Actual Value";
                    fundType = "";
                    fundStatus = "";

                    // Click on plus (+) sign icon (to go to the Add New Fund page)
                    FundDashboardAction.Instance.ClickAddNewFundIcon(10);

                    // Add a new Fund
                    AddEditFundAction.Instance.InputFundManagerToSearchFund("QA Test")
                                              .WaitForLoadingIconToDisappear(20, AddEditFundPage.searchLoadingSpinnerIcon)
                                              .ClickFundManagerAddNewLink(10)
                                              .InputManagerName(managerNameManual)
                                              .InputFundName(fundNameManual)
                                              .ClickAndSelectSubAssetClassDropDown(10, "US Growth Equity")
                                              .InputInceptionDate(10, inceptionDate) // .InputInceptionDateForAddNewFund(10, inceptionDateDay)
                                              .InputLatestActualValue(latestActualValue)
                                              .InputBusinessStreet("QA auto Business Street")
                                              .InputBusinessCity("QA auto Business City")
                                              .InputBusinessState("QA auto Business State")
                                              .InputBusinessZIP("QA auto Business ZIP")
                                              .ClickAndSelectBusinessCountryDropDown(10, "Japan")
                                              .ClickSaveButton()
                                              .WaitForLoadingIconToDisappear(20, AddEditFundPage.loadingIcon);

                    // Verify all informations from Fund Info
                    verifyPoint = toastMessageAddSuccess == FundDashboardAction.Instance.toastMessageAlertGetText(10, toastMessageAddSuccess);
                    verifyPoints.Add(summaryTC = "Verify Fund Info - toast message (Add successfully) is shown correctly after adding: '" + toastMessageAddSuccess + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    // Wait for toast message is disappeared
                    AddEditFundAction.Instance.WaitForElementInvisible(10, FundDashboardPage.toastMessage(toastMessageAddSuccess));

                    verifyPoint = fundNameManual == FundDashboardAction.Instance.FundNameTitleGetText(10);
                    verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Name (Manual) is shown correctly after adding: '" + fundNameManual + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = fundType == FundDashboardAction.Instance.FundInfoGetText(10, "Fund Type");
                    verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Type (Manual) is shown correctly after adding: '" + fundType + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = managerNameManual == FundDashboardAction.Instance.FundInfoGetText(10, "Manager Name");
                    verifyPoints.Add(summaryTC = "Verify Fund Info - Manager Name (Manual) is shown correctly after adding: '" + managerNameManual + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = inceptionDate == FundDashboardAction.Instance.FundInfoGetText(10, "Inception Date");
                    verifyPoints.Add(summaryTC = "Verify Fund Info - Inception Date (Manual) is shown correctly after adding: '" + inceptionDate + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = fundStatus == FundDashboardAction.Instance.FundInfoGetText(10, "Fund Status");
                    verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Status (Manual) is shown correctly after adding: '" + fundStatus + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = latestActualValue == FundDashboardAction.Instance.FundInfoGetText(10, "Latest Actual Value");
                    verifyPoints.Add(summaryTC = "Verify Fund Info - Latest Actual Value (Manual) is shown correctly after adding: '" + latestActualValue + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    // Click on plus (+) sign icon (to go to the Add New Fund page)
                    FundDashboardAction.Instance.ClickAddNewFundIcon(10);

                    // Add a new Fund in an existing Fund Manager
                    AddEditFundAction.Instance.InputFundManagerToSearchFund(managerNameManual)
                                              .WaitForLoadingIconToDisappear(20, AddEditFundPage.searchLoadingSpinnerIcon)
                                              .ClickFundManagerDropdownReturnOfResults(10, managerNameManual)
                                              .InputFundName(fundNameManual + @"_2")
                                              .ClickAndSelectSubAssetClassDropDown(10, "US Growth Equity")
                                              .InputInceptionDate(10, inceptionDate)
                                              .InputLatestActualValue(latestActualValue)
                                              .InputBusinessStreet("QA auto Business Street")
                                              .InputBusinessCity("QA auto Business City")
                                              .InputBusinessState("QA auto Business State")
                                              .InputBusinessZIP("QA auto Business ZIP")
                                              .ClickAndSelectBusinessCountryDropDown(10, "Japan")
                                              .ClickSaveButton()
                                              .WaitForLoadingIconToDisappear(20, AddEditFundPage.loadingIcon);

                    // Verify all informations from Fund Info
                    verifyPoint = toastMessageAddSuccess == FundDashboardAction.Instance.toastMessageAlertGetText(10, toastMessageAddSuccess);
                    verifyPoints.Add(summaryTC = "Verify Fund Info - toast message (Add Fund successfully in an existing Manager) is shown correctly after adding: '" + toastMessageAddSuccess + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    // Wait for toast message is disappeared
                    AddEditFundAction.Instance.WaitForElementInvisible(10, FundDashboardPage.toastMessage(toastMessageAddSuccess));

                    verifyPoint = fundNameManual + @"_2" == FundDashboardAction.Instance.FundNameTitleGetText(10);
                    verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Name (Manual) is shown correctly after adding fund in an existing Manager: '" + fundNameManual + @"_2" + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = fundType == FundDashboardAction.Instance.FundInfoGetText(10, "Fund Type");
                    verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Type (Manual) is shown correctly after adding fund in an existing Manager: '" + fundType + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = managerNameManual == FundDashboardAction.Instance.FundInfoGetText(10, "Manager Name");
                    verifyPoints.Add(summaryTC = "Verify Fund Info - Manager Name (Manual) is shown correctly after adding fund in an existing Manager: '" + managerNameManual + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = inceptionDate == FundDashboardAction.Instance.FundInfoGetText(10, "Inception Date");
                    verifyPoints.Add(summaryTC = "Verify Fund Info - Inception Date (Manual) is shown correctly after adding fund in an existing Manager: '" + inceptionDate + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = fundStatus == FundDashboardAction.Instance.FundInfoGetText(10, "Fund Status");
                    verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Status (Manual) is shown correctly after adding fund in an existing Manager: '" + fundStatus + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = latestActualValue == FundDashboardAction.Instance.FundInfoGetText(10, "Latest Actual Value");
                    verifyPoints.Add(summaryTC = "Verify Fund Info - Latest Actual Value (Manual) is shown correctly after adding fund in an existing Manager: '" + latestActualValue + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);
                }

                if (urlInstance.Contains("conceptia"))
                {
                    Console.WriteLine(summaryTC = "Notes: TC004 Add new Fund is only add new Fund on Sandbox Site!!!");
                    test.Log(Status.Info, summaryTC);
                }
            }
            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Warning
                ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong! Please check console log.");
            }
            #endregion
        }

        [Test, Category("Regression Testing")]
        public void TC005_Add_PrivateFund_AddFirmExists()
        {
            #region Variables declare
            string urlInstance = LoginPage.url;
            string? fundManagerName = null;
            string? firm = null;
            string? fundName1 = null, fundName2 = null, fundName3 = null;
            string message = "Error: Data on following fields are already existed. Please try again.";
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Add Private Fund Test - TC005");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Log into the application
                LoginAction.Instance.LoginSiteNoGodaddy();

                // Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
                if (urlInstance.Contains("sandbox")) 
                {
                    fundManagerName = "PriMan F03";
                    firm = "PriMan F03";
                    fundName1 = "Fund 1 of PriMan F03"; fundName2 = "Fund 2 of PriMan F03"; fundName3 = "Fund 3 of PriMan F03";
                }
                if (urlInstance.Contains("conceptia")) 
                {
                    fundManagerName = "Priman Sta01";
                    firm = "Firm Sta01";
                    fundName1 = "Fund 1 of Firm Sta01"; fundName2 = "Fund 2 of Firm Sta01"; fundName3 = "Fund 3 of Firm Sta01";
                }

                // Workaround --> when click cancel button does not work. (Input a text to 'Search Fund' field)
                FundDashboardAction.Instance.InputNameToSearchFund(10, "abcef").InputNameToSearchFund(10, "");

                // Click on plus (+) sign icon (to go to the Add New Fund page)
                FundDashboardAction.Instance.ClickAddNewFundIcon(10);

                // Add a new Private Fund
                AddEditFundAction.Instance.ClickFundTypeDropdown(10)
                                          .WaitForAllElementsVisible(10, AddEditFundPage.overlayDropdown)
                                          .WaitForAllElementsVisible(10, AddEditFundPage.dropdownSelect("Private Fund"))
                                          .SelectItemInDropdown(10, "Private Fund")
                                          .WaitForElementInvisible(10, AddEditFundPage.overlayDropdown)
                                          .InputTxtLabel(AddEditFundPage.fundNameManager, fundManagerName)
                                          .InputTxtLabel(AddEditFundPage.firm, firm)
                                          .ClickAndSelectItemInDropdown(10, AddEditFundPage.subAssetClass, "FAD Real Estate")
                                          .InputInceptionDate(10, "09/13/2022")
                                          .InputTxtLabel(AddEditFundPage.latestActualValue, "301")
                                          .InputTxtLabel(AddEditFundPage.businessCity, "302")
                                          .InputTxtLabel(AddEditFundPage.businessState, "303")
                                          .ClickAndSelectItemInDropdown(10, AddEditFundPage.businessCountry, "United Kingdom")
                                          .InputTxtLabel(AddEditFundPage.businessStreet, "304")
                                          .InputTxtLabel(AddEditFundPage.businessZip, "305")
                                          .InputTxtLabel(AddEditFundPage.businessContact, "306")
                                          .InputTxtLabel(AddEditFundPage.businessEmail, "307")
                                          .InputTxtLabel(AddEditFundPage.businessPhone, "308")

                                          // Add Fund Name (fund child of Manager)
                                          .InputTxtFundLabel(1, AddEditFundPage.fundName, fundName1)
                                          .InputTxtFundLabel(1, AddEditFundPage.strategy, "3101")
                                          .InputTxtFundLabel(1, AddEditFundPage.vintageYear, "3102")
                                          .InputTxtFundLabel(1, AddEditFundPage.assetClass, "3103")
                                          .InputTxtFundLabel(1, AddEditFundPage.investmentStage, "3104")
                                          .InputTxtFundLabel(1, AddEditFundPage.industryFocus, "3105")
                                          .InputTxtFundLabel(1, AddEditFundPage.geographicFocus, "3106")
                                          .InputTxtFundLabel(1, AddEditFundPage.fundSizeM, "3107")
                                          .InputTxtFundLabel(1, AddEditFundPage.currency, "USD")
                                          .ClickAddFundButton(10)
                                          .InputTxtFundLabel(2, AddEditFundPage.fundName, fundName2)
                                          .InputTxtFundLabel(2, AddEditFundPage.strategy, "3201")
                                          .InputTxtFundLabel(2, AddEditFundPage.vintageYear, "3202")
                                          .InputTxtFundLabel(2, AddEditFundPage.assetClass, "3203")
                                          .InputTxtFundLabel(2, AddEditFundPage.investmentStage, "3204")
                                          .InputTxtFundLabel(2, AddEditFundPage.industryFocus, "3205")
                                          .InputTxtFundLabel(2, AddEditFundPage.geographicFocus, "3206")
                                          .InputTxtFundLabel(2, AddEditFundPage.fundSizeM, "3207")
                                          .InputTxtFundLabel(2, AddEditFundPage.currency, "USD")
                                          .ClickAddFundButton(10)
                                          .InputTxtFundLabel(3, AddEditFundPage.fundName, fundName3)
                                          .InputTxtFundLabel(3, AddEditFundPage.strategy, "3301")
                                          .InputTxtFundLabel(3, AddEditFundPage.vintageYear, "3302")
                                          .InputTxtFundLabel(3, AddEditFundPage.assetClass, "3303")
                                          .InputTxtFundLabel(3, AddEditFundPage.investmentStage, "3304")
                                          .InputTxtFundLabel(3, AddEditFundPage.industryFocus, "3305")
                                          .InputTxtFundLabel(3, AddEditFundPage.geographicFocus, "3306")
                                          .InputTxtFundLabel(3, AddEditFundPage.fundSizeM, "3307")
                                          .InputTxtFundLabel(3, AddEditFundPage.currency, "USD")
                                          .ClickSaveButton();
                                          //.WaitForLoadingIconToDisappear(20, AddEditFundPage.loadingIcon);

                // Verify toast message is shown when add a firm Exists
                verifyPoint = message == FundDashboardAction.Instance.toastMessageAlertGetText(10, message);
                verifyPoints.Add(summaryTC = "Verify toast message (firm Exists) is shown correctly after clicking Save button: '" + message + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Wait for toast message is disappeared
                AddEditFundAction.Instance.WaitForElementInvisible(10, FundDashboardPage.toastMessage(message));

                // Update Firm and Fund Name (fund child of Manager)
                AddEditFundAction.Instance.InputTxtLabel(AddEditFundPage.fundNameManager, "QA_Auto_PrivFund_01")
                                          .InputTxtLabel(AddEditFundPage.firm, "QA_Auto_PrivFund_01")
                                          .InputTxtFundLabel(1, AddEditFundPage.fundName, fundName1)
                                          .InputTxtFundLabel(2, AddEditFundPage.fundName, fundName1)
                                          .ClickSaveButton();
                                          //.WaitForLoadingIconToDisappear(20, AddEditFundPage.loadingIcon);

                // Verify toast message is shown when add FundNames (at Fund List) have same name
                message = "These fields cannot be the same. Please try again.";
                verifyPoint = message == FundDashboardAction.Instance.toastMessageAlertGetText(10, message);
                verifyPoints.Add(summaryTC = "Verify toast message (same Fund Name) is shown correctly after clicking Save button: '" + message + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Wait for toast message is disappeared
                AddEditFundAction.Instance.WaitForElementInvisible(10, FundDashboardPage.toastMessage(message));
            }
            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Warning
                ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong! Please check console log.");
            }
            #endregion
        }

        [Test, Category("Regression Testing")]
        public void TC006_Add_PrivateFund_AddNewFund()
        {
            #region Variables declare
            string urlInstance = LoginPage.url;
            // Manager info
            string fundManagerName;
            string firm;
            string subAssetClass;
            string inceptionDate;
            string latestActualValue;
            string businessCity;
            string businessState;
            string businessCountry;
            string businessStreet;
            string businessZIP;
            string businessContact;
            string contactEmail;
            string contactPhone;
            
            // Fund info
            string fundName1, fundName2, fundName3;
            string strategy1, strategy2, strategy3;
            string yearFirmFounded, vintageYear2, vintageYear3;
            string strategyHeadquarter1, strategyHeadquarter2, strategyHeadquarter3;
            string assetClass1, assetClass2, assetClass3;
            string investmentStage1, investmentStage2, investmentStage3;
            string industryFocus1, industryFocus2, industryFocus3;
            string geographicFocus1, geographicFocus2, geographicFocus3;
            string fundSizeM1, fundSizeM2, fundSizeM3;
            string currency1, currency2, currency3;
            string message = "Add new fund successfully.";
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Add Private Fund Test - TC006");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Log into the application
                LoginAction.Instance.LoginSiteNoGodaddy();

                // Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
                if (urlInstance.Contains("sandbox")) 
                {
                    // Manager Info
                    fundManagerName = "QA_Auto_ManagerPriv" + @"_" + DateTime.Now.ToString().Replace("/", "_").Replace(":", "_").Replace(" ", "_");
                    firm = "QA_Auto_FirmPriv" + @"_" + DateTime.Now.ToString().Replace("/", "_").Replace(":", "_").Replace(" ", "_");
                    subAssetClass = "Natural Resources";
                    string date = DateTime.Now.Date.ToString("MM/dd/yyyy");
                    inceptionDate = date.Replace("-", "/");
                    latestActualValue = "9600000";
                    businessCity = "QA_auto_BC";
                    businessState = "QA_auto_BState";
                    businessCountry = "Japan";
                    businessStreet = "QA_auto_BStreet";
                    businessZIP = "QA_auto_ZIP";
                    businessContact = "QA_auto_BContact";
                    contactEmail = "QA_auto_abc@email.com";
                    contactPhone = "010203040506";

                    // Fund Info
                    fundName1 = "Fund 1 of QA_Auto_FirmPriv"; fundName2 = "Fund 2 of QA_Auto_FirmPriv"; fundName3 = "Fund 3 of QA_Auto_FirmPriv";
                    strategy1 = "3101"; strategy2 = "3201"; strategy3 = "3301";
                    yearFirmFounded = "3102"; vintageYear2 = "3202"; vintageYear3 = "3302";
                    strategyHeadquarter1 = "3103"; strategyHeadquarter2 = "3203"; strategyHeadquarter3 = "3303";
                    assetClass1 = "3104"; assetClass2 = "3204"; assetClass3 = "3304";
                    investmentStage1 = "3105"; investmentStage2 = "3205"; investmentStage3 = "3305";
                    industryFocus1 = "3106"; industryFocus2 = "3206"; industryFocus3 = "3306";
                    geographicFocus1 = "3107"; geographicFocus2 = "3207"; geographicFocus3 = "3307";
                    fundSizeM1 = "3108"; fundSizeM2 = "3208"; fundSizeM3 = "3308";
                    currency1 = "USD"; currency2 = "USD"; currency3 = "USD";

                    // Workaround --> when click cancel button does not work. (Input a text to 'Search Fund' field)
                    FundDashboardAction.Instance.InputNameToSearchFund(10, "abcef").InputNameToSearchFund(10, "");

                    // Click on plus (+) sign icon (to go to the Add New Fund page)
                    FundDashboardAction.Instance.ClickAddNewFundIcon(10);

                    // Add a new Private Fund
                    AddEditFundAction.Instance.ClickFundTypeDropdown(10)
                                              .WaitForAllElementsVisible(10, AddEditFundPage.overlayDropdown)
                                              .WaitForAllElementsVisible(10, AddEditFundPage.dropdownSelect("Private Fund"))
                                              .SelectItemInDropdown(10, "Private Fund")
                                              .WaitForElementInvisible(10, AddEditFundPage.overlayDropdown)
                                              .InputTxtLabel(AddEditFundPage.fundNameManager, fundManagerName)
                                              .InputTxtLabel(AddEditFundPage.firm, firm)
                                              .ClickAndSelectItemInDropdown(10, AddEditFundPage.subAssetClass, "FAD Real Estate")
                                              .InputInceptionDate(10, inceptionDate)
                                              .InputTxtLabel(AddEditFundPage.latestActualValue, latestActualValue)
                                              .InputTxtLabel(AddEditFundPage.businessCity, businessCity)
                                              .InputTxtLabel(AddEditFundPage.businessState, businessState)
                                              .ClickAndSelectItemInDropdown(10, AddEditFundPage.businessCountry, businessCountry)
                                              .InputTxtLabel(AddEditFundPage.businessStreet, businessStreet)
                                              .InputTxtLabel(AddEditFundPage.businessZip, businessZIP)
                                              .InputTxtLabel(AddEditFundPage.businessContact, businessContact)
                                              .InputTxtLabel(AddEditFundPage.businessEmail, contactEmail)
                                              .InputTxtLabel(AddEditFundPage.businessPhone, contactPhone)

                                              // Add Fund Name (fund child of Manager)
                                              .InputTxtFundLabel(1, AddEditFundPage.fundName, fundName1)
                                              .InputTxtFundLabel(1, AddEditFundPage.strategy, strategy1)
                                              .InputTxtFundLabel(1, AddEditFundPage.vintageYear, yearFirmFounded)
                                              .InputTxtFundLabel(1, AddEditFundPage.strategyHeadquarter, strategyHeadquarter1)
                                              .InputTxtFundLabel(1, AddEditFundPage.assetClass, assetClass1)
                                              .InputTxtFundLabel(1, AddEditFundPage.investmentStage, investmentStage1)
                                              .InputTxtFundLabel(1, AddEditFundPage.industryFocus, industryFocus1)
                                              .InputTxtFundLabel(1, AddEditFundPage.geographicFocus, geographicFocus1)
                                              .InputTxtFundLabel(1, AddEditFundPage.fundSizeM, fundSizeM1)
                                              .InputTxtFundLabel(1, AddEditFundPage.currency, currency1)
                                              //.InputTxtFundLabel(1, AddEditFundPage.cambridgeVintageYear, yearFirmFounded)

                                              .ClickAddFundButton(10)
                                              .InputTxtFundLabel(2, AddEditFundPage.fundName, fundName2)
                                              .InputTxtFundLabel(2, AddEditFundPage.strategy, strategy2)
                                              .InputTxtFundLabel(2, AddEditFundPage.vintageYear, vintageYear2)
                                              .InputTxtFundLabel(2, AddEditFundPage.strategyHeadquarter, strategyHeadquarter2)
                                              .InputTxtFundLabel(2, AddEditFundPage.assetClass, assetClass2)
                                              .InputTxtFundLabel(2, AddEditFundPage.investmentStage, investmentStage2)
                                              .InputTxtFundLabel(2, AddEditFundPage.industryFocus, industryFocus2)
                                              .InputTxtFundLabel(2, AddEditFundPage.geographicFocus, geographicFocus2)
                                              .InputTxtFundLabel(2, AddEditFundPage.fundSizeM, fundSizeM2)
                                              .InputTxtFundLabel(2, AddEditFundPage.currency, currency2)
                                              //.InputTxtFundLabel(2, AddEditFundPage.cambridgeVintageYear, vintageYear2)
                                              .ClickAddFundButton(10)
                                              .InputTxtFundLabel(3, AddEditFundPage.fundName, fundName3)
                                              .InputTxtFundLabel(3, AddEditFundPage.strategy, strategy3)
                                              .InputTxtFundLabel(3, AddEditFundPage.vintageYear, vintageYear3)
                                              .InputTxtFundLabel(3, AddEditFundPage.strategyHeadquarter, strategyHeadquarter3)
                                              .InputTxtFundLabel(3, AddEditFundPage.assetClass, assetClass3)
                                              .InputTxtFundLabel(3, AddEditFundPage.investmentStage, investmentStage3)
                                              .InputTxtFundLabel(3, AddEditFundPage.industryFocus, industryFocus3)
                                              .InputTxtFundLabel(3, AddEditFundPage.geographicFocus, geographicFocus3)
                                              .InputTxtFundLabel(3, AddEditFundPage.fundSizeM, fundSizeM3)
                                              .InputTxtFundLabel(3, AddEditFundPage.currency, currency3)
                                              //.InputTxtFundLabel(3, AddEditFundPage.cambridgeVintageYear, vintageYear3)
                                              .ClickSaveButton();
                                              //.WaitForLoadingIconToDisappear(20, AddEditFundPage.loadingIcon);

                    // Verify toast message (Add successfully) Private Fund is shown correctly after adding
                    verifyPoint = message == FundDashboardAction.Instance.toastMessageAlertGetText(10, message);
                    verifyPoints.Add(summaryTC = "Verify toast message (Add successfully) Private Fund is shown correctly after adding: '" + message + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    // Wait for toast message is disappeared
                    AddEditFundAction.Instance.WaitForElementInvisible(10, FundDashboardPage.toastMessage(message));
                                              //.WaitForLoadingIconToDisappear(20, AddEditFundPage.loadingIcon);

                    #region Verify all informations from Manager Info (tab)
                    // Click on Manager Info tab
                    FundDashboardAction.Instance.ClickManagerInfoTab(10);

                    // Verify all informations from Manager Info (tab)
                    verifyPoint = firm == FundDashboardAction.Instance.FundNameTitleGetText(10); // fundManagerName
                    verifyPoints.Add(summaryTC = "Verify Manager Info - Firm (Cambridge Fund - Manual) is shown correctly after searching: '" + firm + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = yearFirmFounded == FundDashboardAction.Instance.FundInfoGetText(10, FundDashboardPage.yearFirmFounded);
                    verifyPoints.Add(summaryTC = "Verify Manager Info - Year Firm Founded (Cambridge Fund - Manual) is shown correctly after searching: '" + yearFirmFounded + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = strategyHeadquarter1 == FundDashboardAction.Instance.FundInfoGetText(10, FundDashboardPage.strategyHeadquarter);
                    verifyPoints.Add(summaryTC = "Verify Manager Info - Strategy Headquarter (Cambridge Fund - Manual) is shown correctly after searching: '" + strategyHeadquarter1 + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = businessContact == FundDashboardAction.Instance.FundInfoGetText(10, FundDashboardPage.businessContact);
                    verifyPoints.Add(summaryTC = "Verify Manager Info - Business Contact (Cambridge Fund - Manual) is shown correctly after searching: '" + businessContact + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = contactPhone == FundDashboardAction.Instance.FundInfoGetText(10, FundDashboardPage.contactPhone);
                    verifyPoints.Add(summaryTC = "Verify Manager Info - Contact Phone (Cambridge Fund - Manual) is shown correctly after searching: '" + contactPhone + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = contactEmail == FundDashboardAction.Instance.FundInfoGetText(10, FundDashboardPage.contactEmail);
                    verifyPoints.Add(summaryTC = "Verify Manager Info - Contact Email (Cambridge Fund - Manual) is shown correctly after searching: '" + contactEmail + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);
                    #endregion

                    #region Verify all informations from Fund List (tab)
                    // Click on Fund List tab
                    FundDashboardAction.Instance.ClickFundListTab(10)
                                                .WaitForElementVisible(10, FundDashboardPage.expandAllBtn)
                                                .ClickExpandAllButton(10); // --> CLick Expand All button

                    // Verify all informations from Fund List (tab)
                    verifyPoint = strategy1 == FundDashboardAction.Instance.ListStrategiesGetText(10, 1);
                    verifyPoints.Add(summaryTC = "Verify Fund List - Strategy 1 (Cambridge Fund - Manual) is shown correctly after searching: '" + strategy1 + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = strategy2 == FundDashboardAction.Instance.ListStrategiesGetText(10, 2);
                    verifyPoints.Add(summaryTC = "Verify Fund List - Strategy 2 (Cambridge Fund - Manual) is shown correctly after searching: '" + strategy2 + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = strategy3 == FundDashboardAction.Instance.ListStrategiesGetText(10, 3);
                    verifyPoints.Add(summaryTC = "Verify Fund List - Strategy 3 (Cambridge Fund - Manual) is shown correctly after searching: '" + strategy3 + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = fundName1 == FundDashboardAction.Instance.ListFundsNameGetText(10, 1, 1);
                    verifyPoints.Add(summaryTC = "Verify Fund List - Fund name 1 (Cambridge Fund - Manual) is shown correctly after searching: '" + fundName1 + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = fundName2 == FundDashboardAction.Instance.ListFundsNameGetText(10, 2, 1);
                    verifyPoints.Add(summaryTC = "Verify Fund List - Fund name 2 (Cambridge Fund - Manual) is shown correctly after searching: '" + fundName2 + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = fundName3 == FundDashboardAction.Instance.ListFundsNameGetText(10, 3, 1);
                    verifyPoints.Add(summaryTC = "Verify Fund List - Fund name 3 (Cambridge Fund - Manual) is shown correctly after searching: '" + fundName3 + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = yearFirmFounded == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.vintageYear)
                                  && vintageYear2 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 2, 1, FundDashboardPage.vintageYear)
                                  && vintageYear3 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 3, 1, FundDashboardPage.vintageYear);
                    verifyPoints.Add(summaryTC = "Verify Fund List - Vintage Year (Cambridge Fund) is shown correctly after searching: VY1='" + yearFirmFounded + "', VY2='" + vintageYear2 + "', VY3='" + vintageYear3 + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = assetClass1 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.assetClass)
                               && assetClass2 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 2, 1, FundDashboardPage.assetClass)
                               && assetClass3 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 3, 1, FundDashboardPage.assetClass);
                    verifyPoints.Add(summaryTC = "Verify Fund List - Asset Class (Cambridge Fund) is shown correctly after searching: AC1='" + assetClass1 + "', AC2='" + assetClass2 + "', AC3='" + assetClass3 + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = investmentStage1 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.investmentStage)
                               && investmentStage2 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 2, 1, FundDashboardPage.investmentStage)
                               && investmentStage3 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 3, 1, FundDashboardPage.investmentStage);
                    verifyPoints.Add(summaryTC = "Verify Fund List - Investment Stage (Cambridge Fund) is shown correctly after searching: IS1='" + investmentStage1 + "', IS2='" + investmentStage2 + "', IS3='" + investmentStage3 + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = industryFocus1 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.industryFocus)
                               && industryFocus2 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 2, 1, FundDashboardPage.industryFocus)
                               && industryFocus3 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 3, 1, FundDashboardPage.industryFocus);
                    verifyPoints.Add(summaryTC = "Verify Fund List - Industry Focus (Cambridge Fund) is shown correctly after searching: IF1='" + industryFocus1 + "', IF2='" + industryFocus2 + "', IF3='" + industryFocus3 + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = geographicFocus1 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.geographicFocus)
                               && geographicFocus2 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 2, 1, FundDashboardPage.geographicFocus)
                               && geographicFocus3 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 3, 1, FundDashboardPage.geographicFocus);
                    verifyPoints.Add(summaryTC = "Verify Fund List - Geographic Focus (Cambridge Fund) is shown correctly after searching: GF1='" + geographicFocus1 + "', GF2='" + geographicFocus2 + "', GF3='" + geographicFocus3 + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    verifyPoint = fundSizeM1 + " " + currency1 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.fundSizeM)
                               && fundSizeM2 + " " + currency2 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 2, 1, FundDashboardPage.fundSizeM)
                               && fundSizeM3 + " " + currency3 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 3, 1, FundDashboardPage.fundSizeM);
                    verifyPoints.Add(summaryTC = "Verify Fund List - Fund Size(M) (Cambridge Fund) is shown correctly after searching: FS1='" + fundSizeM1 + " " + currency1 + "', FS2='" + fundSizeM2 + " " + currency2 + "', FS3='" + fundSizeM3 + " " + currency3 + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    //verifyPoint = currency1 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.currency)
                    //           && currency2 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 2, 1, FundDashboardPage.currency)
                    //           && currency3 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 3, 1, FundDashboardPage.currency);
                    //verifyPoints.Add(summaryTC = "Verify Fund List - Currency (Cambridge Fund) is shown correctly after searching: Cur1='" + currency1 + "', Cur2='" + currency2 + "', Cur3='" + currency3 + "'", verifyPoint);
                    //ExtReportResult(verifyPoint, summaryTC);
                    #endregion
                }

                if (urlInstance.Contains("conceptia"))
                {
                    Console.WriteLine(summaryTC = "Notes: TC006 Add new Private Fund is only add new Fund on Sandbox Site!!!");
                    test.Log(Status.Info, summaryTC);
                }

            }
            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Warning
                ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong! Please check console log.");
            }
            #endregion
        }

        [Test, Category("Regression Testing")]
        public void TC007_Edit_PrivateFund()
        {
            #region Variables declare
            string urlInstance = LoginPage.url;
            const string sourceIcon = "M";
            string? inputSearch = null;
            // Manager Info
            string? fundManagerName = null, fundManagerNameUpdate = null;
            string? firm = null;
            string? subAssetClass = null, subAssetClassUpdate = null, inceptionDate = null, 
                    latestActualValue = null, latestActualValueUpdate = null, 
                    businessCity = null, businessCityUpdate = null, businessState = null, businessStateUpdate = null,
                    businessCountry = null, businessCountryUpdate = null, businessStreet = null, businessStreetUpdate = null,
                    businessZip = null, businessZipUpdate = null, businessContact = null, businessContactUpdate = null,
                    businessEmail = null, businessEmailUpdate = null, businessPhone = null, businessPhoneUpdate = null;

            // Fund Info
            string? fundName1 = null, fundName1Update = null, fundName2 = null, fundName3 = null, fundName4 = null;
            string? strategy1 = null, strategy1Update = null, strategy2 = null, strategy3 = null, strategy4 = null;
            string? vintageYear1 = null, vintageYear1Update = null, vintageYear2 = null, vintageYear3 = null, vintageYear4 = null;
            string? strategyHeadquarter1 = null, strategyHeadquarter1Update = null, strategyHeadquarter2 = null, strategyHeadquarter3 = null, strategyHeadquarter4 = null;
            string? assetClass1 = null, assetClass1Update = null, assetClass2 = null, assetClass3 = null, assetClass4 = null;
            string? investmentStage1 = null, investmentStage1Update = null, investmentStage2 = null, investmentStage3 = null, investmentStage4 = null;
            string? industryFocus1 = null, industryFocus1Update = null, industryFocus2 = null, industryFocus3 = null, industryFocus4 = null;
            string? geographicFocus1 = null, geographicFocus1Update = null, geographicFocus2 = null, geographicFocus3 = null, geographicFocus4 = null;
            int? fundSizeM1 = null, fundSizeM1Update = null, fundSizeM2 = null, fundSizeM3 = null, fundSizeM4 = null;
            string? currency1 = null, currency1Update = null, currency2 = null, currency3 = null, currency4 = null;
            string message = "These fields cannot be the same. Please try again.";
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Edit Private Fund Test - TC007");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Log into the application
                LoginAction.Instance.LoginSiteNoGodaddy();

                // Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
                if (urlInstance.Contains("sandbox"))
                {  
                    // Manager Info 
                    inputSearch = "Firmofpriman";
                    firm = "Firmofpriman qatest 1"; // (id=285)
                    fundManagerName = "Priman qatest 01"; fundManagerNameUpdate = fundManagerName + "_Update";
                    subAssetClass = "US Buyout"; subAssetClassUpdate = "FAD Real Estate";
                    latestActualValue = "LAVA 01"; latestActualValueUpdate = latestActualValue + "_Update";
                    inceptionDate = "09/23/2022"; businessCity = "BCIT 01"; businessCityUpdate = businessCity + "_Update"; businessState = "BSTA 01"; businessStateUpdate = businessState + "_Update";
                    businessCountry = "United Kingdom"; businessCountryUpdate = "Japan"; businessStreet = "BSTR 01"; businessStreetUpdate = businessStreet + "_Update"; 
                    businessZip = "BZIP 01"; businessZipUpdate = businessZip + "_Update"; businessContact = "BCON 01"; businessContactUpdate = businessContact + "_Update";
                    businessEmail = "BEMA01@test.com"; businessEmailUpdate = businessEmail + "_Update"; businessPhone = "123456789"; businessPhoneUpdate = "987654321";

                    // Fund Info
                    fundName1 = "F1 of Firmofpriman qatest 1"; fundName2 = "F2 of Firmofpriman qatest 1"; fundName3 = "F3 of Firmofpriman qatest 1"; fundName4 = "F4 of Firmofpriman qatest 1";
                    fundName1Update = fundName1 + "_Update";
                    strategy1 = "STRA 101"; strategy1Update = strategy1 + "_Update"; strategy2 = "STRA 201"; strategy3 = "STRA 301"; strategy4 = "STRA 401";
                    vintageYear1 = "2020"; vintageYear1Update = "2030"; vintageYear2 = "2021"; vintageYear3 = "2022"; vintageYear4 = "2023";
                    strategyHeadquarter1 = "SHQA 101"; strategyHeadquarter1Update = strategyHeadquarter1 + "_Update"; strategyHeadquarter2 = "SHQA 201"; strategyHeadquarter3 = "SHQA 301"; strategyHeadquarter4 = "SHQA 401";
                    assetClass1 = "AC 101"; assetClass1Update = assetClass1 + "_Update"; assetClass2 = "AC 201"; assetClass3 = "AC 301"; assetClass4 = "AC 401";
                    investmentStage1 = "IS 101"; investmentStage1Update = investmentStage1 + "_Update"; investmentStage2 = "IS 201"; investmentStage3 = "IS 301"; investmentStage4 = "IS 401";
                    industryFocus1 = "IF 101"; industryFocus1Update = industryFocus1 + "_Update"; industryFocus2 = "IF 201"; industryFocus3 = "IF 301"; industryFocus4 = "IF 401";
                    geographicFocus1 = "GF 101"; geographicFocus1Update = geographicFocus1 + "_Update"; geographicFocus2 = "GF 201"; geographicFocus3 = "GF 301"; geographicFocus4 = "GF 401";
                    fundSizeM1 = 101; fundSizeM1Update = 1011; fundSizeM2 = 201; fundSizeM3 = 301; fundSizeM4 = 401;
                    currency1 = "USD"; currency1Update = "VND"; currency2 = "USD"; currency3 = "USD"; currency4 = "USD";
                }
                if (urlInstance.Contains("conceptia"))
                {
                    // Manager Info
                    inputSearch = "priman";
                    firm = "Priman Sta01";
                    fundManagerName = "Priman Sta02"; fundManagerNameUpdate = fundManagerName + "_Update";
                    subAssetClass = "FAD Real Estate"; subAssetClassUpdate = "US Buyout";
                    latestActualValue = "LAVA 01"; latestActualValueUpdate = latestActualValue + "_Update";
                    inceptionDate = "09/13/2022"; businessCity = "BCIT 01"; businessCityUpdate = businessCity + "_Update"; businessState = "BSTA 01"; businessStateUpdate = businessState + "_Update";
                    businessCountry = "United Kingdom"; businessCountryUpdate = "Japan"; businessStreet = "BSTR 01"; businessStreetUpdate = businessStreet + "_Update";
                    businessZip = "BZIP 01"; businessZipUpdate = businessZip + "_Update"; businessContact = "BCON 01"; businessContactUpdate = businessContact + "_Update";
                    businessEmail = "BEMA01@test.com"; businessEmailUpdate = businessEmail + "_Update"; businessPhone = "123456789"; businessPhoneUpdate = "987654321";

                    // Fund Info
                    fundName1 = "Fund 1 of Firm Sta01"; fundName2 = "Fund 2 of Firm Sta01"; fundName3 = "Fund 3 of Firm Sta01"; fundName4 = "Fund 4 of Firm Sta01"; 
                    fundName1Update = fundName1 + "_Update";
                    strategy1 = "STRA 101"; strategy1Update = strategy1 + "_Update"; strategy2 = "STRA 201"; strategy3 = "STRA 301"; strategy4 = "STRA 401";
                    vintageYear1 = "2020"; vintageYear1Update = "2030"; vintageYear2 = "2021"; vintageYear3 = "2022"; vintageYear4 = "2023";
                    strategyHeadquarter1 = "SHQA 101"; strategyHeadquarter1Update = strategyHeadquarter1 + "_Update"; strategyHeadquarter2 = "SHQA 201"; strategyHeadquarter3 = "SHQA 301"; strategyHeadquarter4 = "SHQA 401";
                    assetClass1 = "AC 101"; assetClass1Update = assetClass1 + "_Update"; assetClass2 = "AC 201"; assetClass3 = "AC 301"; assetClass4 = "AC 401";
                    investmentStage1 = "IS 101"; investmentStage1Update = investmentStage1 + "_Update"; investmentStage2 = "IS 201"; investmentStage3 = "IS 301"; investmentStage4 = "IS 401";
                    industryFocus1 = "IF 101"; industryFocus1Update = industryFocus1 + "_Update"; industryFocus2 = "IF 201"; industryFocus3 = "IF 301"; industryFocus4 = "IF 401";
                    geographicFocus1 = "GF 101"; geographicFocus1Update = geographicFocus1 + "_Update"; geographicFocus2 = "GF 201"; geographicFocus3 = "GF 301"; geographicFocus4 = "GF 401";
                    fundSizeM1 = 101; fundSizeM1Update = 1011; fundSizeM2 = 201; fundSizeM3 = 301; fundSizeM4 = 401;
                    currency1 = "USD"; currency1Update = "VND"; currency2 = "USD"; currency3 = "USD"; currency4 = "USD";
                }

                // Search a Cambridge Fund (KS-455 the Firm should be searched in the search bar)
                FundDashboardAction.Instance.InputNameToSearchFund(10, inputSearch)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResultsWithItemSource(firm, sourceIcon))
                                            .ClickFundNameReturnOfResults(10, firm, sourceIcon);
                                            //.WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);

                // Verify all informations from Manager Info (tab)
                verifyPoint = firm == FundDashboardAction.Instance.FundNameTitleGetText(10);
                verifyPoints.Add(summaryTC = "Verify Manager Info - firm (Cambridge Fund - Manual) is shown correctly after searching: '" + firm + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click on pencil icon (to go to the edit-Fund page)
                FundDashboardAction.Instance.ClickEditFundPencilIcon(10)
                                            //.WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon)
                                            .WaitForElementVisible(10, AddEditFundPage.labelFundInputTxt(1, "Fund Name *"));

                // Edit a Private Fund (add FundNames at Fund List have same name)
                AddEditFundAction.Instance.ClickAddFundButton(10)
                                          .InputTxtFundLabel(4, AddEditFundPage.fundName, fundName1) // Add an existing Fund Name (fund child of Manager)
                                          .ClickSaveButton();

                // Verify toast message is shown when add FundNames (at Fund List) have same name
                System.Threading.Thread.Sleep(1000);
                verifyPoint = message == FundDashboardAction.Instance.toastMessageAlertGetText(10, message);
                verifyPoints.Add(summaryTC = "Verify toast message (same Fund Name) is shown correctly after clicking Save button: '" + message + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Wait for toast message is disappeared
                AddEditFundAction.Instance.WaitForElementInvisible(10, FundDashboardPage.toastMessage(message));

                // Edit a Private Fund Successfully
                AddEditFundAction.Instance.InputTxtFundLabel(4, AddEditFundPage.fundName, fundName4)
                                          .InputTxtFundLabel(4, AddEditFundPage.strategy, strategy4)
                                          .InputTxtFundLabel(4, AddEditFundPage.vintageYear, vintageYear4)
                                          .InputTxtFundLabel(4, AddEditFundPage.strategyHeadquarter, strategyHeadquarter4)
                                          .InputTxtFundLabel(4, AddEditFundPage.assetClass, assetClass4)
                                          .InputTxtFundLabel(4, AddEditFundPage.investmentStage, investmentStage4)
                                          .InputTxtFundLabel(4, AddEditFundPage.industryFocus, industryFocus4)
                                          .InputTxtFundLabel(4, AddEditFundPage.geographicFocus, geographicFocus4)
                                          .InputTxtFundLabel(4, AddEditFundPage.fundSizeM, fundSizeM4.ToString())
                                          .InputTxtFundLabel(4, AddEditFundPage.currency, currency4)
                                          //.InputTxtFundLabel(4, AddEditFundPage.cambridgeVintageYear, vintageYear4)
                                          .InputTxtFundLabel(1, AddEditFundPage.fundName, fundName1Update)
                                          .InputTxtFundLabel(1, AddEditFundPage.strategy, strategy1Update)
                                          .InputTxtFundLabel(1, AddEditFundPage.vintageYear, vintageYear1Update)
                                          .InputTxtFundLabel(1, AddEditFundPage.strategyHeadquarter, strategyHeadquarter1Update)
                                          .InputTxtFundLabel(1, AddEditFundPage.assetClass, assetClass1Update)
                                          .InputTxtFundLabel(1, AddEditFundPage.investmentStage, investmentStage1Update)
                                          .InputTxtFundLabel(1, AddEditFundPage.industryFocus, industryFocus1Update)
                                          .InputTxtFundLabel(1, AddEditFundPage.geographicFocus, geographicFocus1Update)
                                          .InputTxtFundLabel(1, AddEditFundPage.fundSizeM, (fundSizeM1Update).ToString())
                                          .InputTxtFundLabel(1, AddEditFundPage.currency, currency1Update)
                                          //.InputTxtFundLabel(1, AddEditFundPage.cambridgeVintageYear, vintageYear1 + "_Update")
                                          .InputTxtLabel(AddEditFundPage.fundNameManager, fundManagerNameUpdate)
                                          .ClickAndSelectItemInDropdown(10, AddEditFundPage.subAssetClass, subAssetClassUpdate)
                                          .InputInceptionDate(10, "09/24/2022")
                                          .InputTxtLabel(AddEditFundPage.latestActualValue, latestActualValueUpdate)
                                          .InputTxtLabel(AddEditFundPage.businessCity, businessCityUpdate)
                                          .InputTxtLabel(AddEditFundPage.businessState, businessStateUpdate)
                                          .ClickAndSelectItemInDropdown(10, AddEditFundPage.businessCountry, businessCountryUpdate)
                                          .InputTxtLabel(AddEditFundPage.businessStreet, businessStreetUpdate)
                                          .InputTxtLabel(AddEditFundPage.businessZip, businessZipUpdate)
                                          .InputTxtLabel(AddEditFundPage.businessContact, businessContactUpdate)
                                          .InputTxtLabel(AddEditFundPage.businessEmail, businessEmailUpdate)
                                          .InputTxtLabel(AddEditFundPage.businessPhone, businessPhoneUpdate)
                                          .ClickSaveButton();

                // Verify toast message (Edit successfully) Private Fund is shown correctly after editing
                message = "Update fund successfully.";
                verifyPoint = message == FundDashboardAction.Instance.toastMessageAlertGetText(20, message);
                verifyPoints.Add(summaryTC = "Verify toast message (edit successfully) Private Fund is shown correctly after editing: '" + message + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Wait for toast message is disappeared
                AddEditFundAction.Instance.WaitForElementInvisible(10, FundDashboardPage.toastMessage(message));

                // Check if Loading (Spinner) Icon is displayed then wait for it to disappear
                if (AddEditFundAction.Instance.IsElementPresent(AddEditFundPage.loadingIcon))
                {
                    AddEditFundAction.Instance.WaitForLoadingIconToDisappear(20, AddEditFundPage.loadingIcon);
                }

                #region Verify all informations from Manager Info (tab)
                // Click on Manager Info tab
                FundDashboardAction.Instance.ClickManagerInfoTab(10);

                // Verify all informations from Manager Info (tab)
                verifyPoint = vintageYear2 == FundDashboardAction.Instance.FundInfoGetText(10, FundDashboardPage.yearFirmFounded); // vintageYear1Update
                verifyPoints.Add(summaryTC = "Verify Manager Info - Year Firm Founded (Cambridge Fund - Manual) is shown correctly after editing: '" + vintageYear2 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = strategyHeadquarter2 == FundDashboardAction.Instance.FundInfoGetText(10, FundDashboardPage.strategyHeadquarter); // strategyHeadquarter1Update
                verifyPoints.Add(summaryTC = "Verify Manager Info - Strategy Headquarter (Cambridge Fund - Manual) is shown correctly after editing: '" + strategyHeadquarter2 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = businessContactUpdate == FundDashboardAction.Instance.FundInfoGetText(10, FundDashboardPage.businessContact);
                verifyPoints.Add(summaryTC = "Verify Manager Info - Business Contact (Cambridge Fund - Manual) is shown correctly after editing: '" + businessContactUpdate + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = businessPhoneUpdate == FundDashboardAction.Instance.FundInfoGetText(10, FundDashboardPage.contactPhone);
                verifyPoints.Add(summaryTC = "Verify Manager Info - Business Phone (Cambridge Fund - Manual) is shown correctly after editing: '" + businessPhoneUpdate + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = businessEmailUpdate == FundDashboardAction.Instance.FundInfoGetText(10, FundDashboardPage.contactEmail);
                verifyPoints.Add(summaryTC = "Verify Manager Info - Business Email (Cambridge Fund - Manual) is shown correctly after editing: '" + businessEmailUpdate + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                #endregion

                #region Verify all informations from Fund List (tab)
                // Click on Fund List tab
                FundDashboardAction.Instance.ClickFundListTab(10)
                                            .WaitForElementVisible(10, FundDashboardPage.expandAllBtn)
                                            .ClickExpandAllButton(10); // --> CLick Expand All button


                // Verify all informations from Fund List (tab)
                verifyPoint = strategy1Update == FundDashboardAction.Instance.ListStrategiesGetText(10, 1);
                verifyPoints.Add(summaryTC = "Verify Fund List - Strategy 1 (Cambridge Fund - Manual) is shown correctly after searching: '" + strategy1Update + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = strategy2 == FundDashboardAction.Instance.ListStrategiesGetText(10, 2);
                verifyPoints.Add(summaryTC = "Verify Fund List - Strategy 2 (Cambridge Fund - Manual) is shown correctly after searching: '" + strategy2 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = strategy3 == FundDashboardAction.Instance.ListStrategiesGetText(10, 3);
                verifyPoints.Add(summaryTC = "Verify Fund List - Strategy 3 (Cambridge Fund - Manual) is shown correctly after searching: '" + strategy3 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = strategy4 == FundDashboardAction.Instance.ListStrategiesGetText(10, 4);
                verifyPoints.Add(summaryTC = "Verify Fund List - Strategy 4 (Cambridge Fund - Manual) is shown correctly after searching: '" + strategy4 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = fundName1Update == FundDashboardAction.Instance.ListFundsNameGetText(10, 1, 1);
                verifyPoints.Add(summaryTC = "Verify Fund List - Fund name 1 (Cambridge Fund - Manual) is shown correctly after searching: '" + fundName1Update + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = fundName2 == FundDashboardAction.Instance.ListFundsNameGetText(10, 2, 1);
                verifyPoints.Add(summaryTC = "Verify Fund List - Fund name 2 (Cambridge Fund - Manual) is shown correctly after searching: '" + fundName2 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = fundName3 == FundDashboardAction.Instance.ListFundsNameGetText(10, 3, 1);
                verifyPoints.Add(summaryTC = "Verify Fund List - Fund name 3 (Cambridge Fund - Manual) is shown correctly after searching: '" + fundName3 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = fundName4 == FundDashboardAction.Instance.ListFundsNameGetText(10, 4, 1);
                verifyPoints.Add(summaryTC = "Verify Fund List - Fund name 4 (Cambridge Fund - Manual) is shown correctly after searching: '" + fundName4 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = vintageYear1Update == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.vintageYear)
                           && vintageYear2 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 2, 1, FundDashboardPage.vintageYear)
                           && vintageYear3 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 3, 1, FundDashboardPage.vintageYear)
                           && vintageYear4 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 4, 1, FundDashboardPage.vintageYear);
                verifyPoints.Add(summaryTC = "Verify Fund List - Vintage Year (Cambridge Fund) is shown correctly after searching: VY1='" + vintageYear1Update + "', VY2='" + vintageYear2 + "', VY3='" + vintageYear3 + "', VY4='" + vintageYear4 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = assetClass1Update == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.assetClass)
                           && assetClass2 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 2, 1, FundDashboardPage.assetClass)
                           && assetClass3 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 3, 1, FundDashboardPage.assetClass)
                           && assetClass4 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 4, 1, FundDashboardPage.assetClass);
                verifyPoints.Add(summaryTC = "Verify Fund List - Asset Class (Cambridge Fund) is shown correctly after searching: AC1='" + assetClass1Update + "', AC2='" + assetClass2 + "', AC3='" + assetClass3 + "', AC4='" + assetClass4 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = investmentStage1Update == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.investmentStage)
                           && investmentStage2 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 2, 1, FundDashboardPage.investmentStage)
                           && investmentStage3 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 3, 1, FundDashboardPage.investmentStage)
                           && investmentStage4 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 4, 1, FundDashboardPage.investmentStage);
                verifyPoints.Add(summaryTC = "Verify Fund List - Investment Stage (Cambridge Fund) is shown correctly after searching: IS1='" + investmentStage1Update + "', IS2='" + investmentStage2 + "', IS3='" + investmentStage3 + "', IS4='" + investmentStage4 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = industryFocus1Update == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.industryFocus)
                           && industryFocus2 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 2, 1, FundDashboardPage.industryFocus)
                           && industryFocus3 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 3, 1, FundDashboardPage.industryFocus)
                           && industryFocus4 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 4, 1, FundDashboardPage.industryFocus);
                verifyPoints.Add(summaryTC = "Verify Fund List - Industry Focus (Cambridge Fund) is shown correctly after searching: IF1='" + industryFocus1Update + "', IF2='" + industryFocus2 + "', IF3='" + industryFocus3 + "', IF4='" + industryFocus4 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = geographicFocus1Update == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.geographicFocus)
                           && geographicFocus2 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 2, 1, FundDashboardPage.geographicFocus)
                           && geographicFocus3 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 3, 1, FundDashboardPage.geographicFocus)
                           && geographicFocus4 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 4, 1, FundDashboardPage.geographicFocus);
                verifyPoints.Add(summaryTC = "Verify Fund List - Geographic Focus (Cambridge Fund) is shown correctly after searching: GF1='" + geographicFocus1Update + "', GF2='" + geographicFocus2 + "', GF3='" + geographicFocus3 + "', GF4='" + geographicFocus4 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = fundSizeM1Update + " " + currency1Update == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.fundSizeM)
                           && fundSizeM2 + " " + currency2 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 2, 1, FundDashboardPage.fundSizeM)
                           && fundSizeM3 + " " + currency3 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 3, 1, FundDashboardPage.fundSizeM)
                           && fundSizeM4 + " " + currency4 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 4, 1, FundDashboardPage.fundSizeM);
                verifyPoints.Add(summaryTC = "Verify Fund List - Fund Size(M) (Cambridge Fund) is shown correctly after searching: FS1='" + fundSizeM1Update + " " + currency1Update + "', FS2='" + fundSizeM2 + " " + currency2 + "', FS3='" + fundSizeM3 + " " + currency3 + "', FS4='" + fundSizeM4 + " " + currency4 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                #endregion

                #region Re-Update the Fund Manager Name (back to the original Fund Manager Name)
                // Click on pencil icon (to go to the edit-Fund page)
                FundDashboardAction.Instance.ClickEditFundPencilIcon(10)
                                            //.WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon)
                                            .WaitForElementVisible(10, AddEditFundPage.labelFundInputTxt(1, "Fund Name *"));

                // Edit a Private Fund Successfully
                AddEditFundAction.Instance.InputTxtLabel(AddEditFundPage.fundNameManager, fundManagerName)
                                          .ClickAndSelectItemInDropdown(10, AddEditFundPage.subAssetClass, subAssetClass)
                                          .InputInceptionDate(10, inceptionDate)
                                          .InputTxtLabel(AddEditFundPage.latestActualValue, latestActualValue)
                                          .InputTxtLabel(AddEditFundPage.businessCity, businessCity)
                                          .InputTxtLabel(AddEditFundPage.businessState, businessState)
                                          .ClickAndSelectItemInDropdown(10, AddEditFundPage.businessCountry, businessCountry)
                                          .InputTxtLabel(AddEditFundPage.businessStreet, businessStreet)
                                          .InputTxtLabel(AddEditFundPage.businessZip, businessZip)
                                          .InputTxtLabel(AddEditFundPage.businessContact, businessContact)
                                          .InputTxtLabel(AddEditFundPage.businessEmail, businessEmail)
                                          .InputTxtLabel(AddEditFundPage.businessPhone, businessPhone)
                                          .ClickDeleteFundNameButton(10, 1) // Delete Fund Name
                                          .ClickDeleteFundNameButton(10, 1) // Delete Fund Name
                                          .ClickDeleteFundNameButton(10, 1) // Delete Fund Name
                                          .InputTxtFundLabel(1, AddEditFundPage.fundName, fundName1)
                                          .InputTxtFundLabel(1, AddEditFundPage.strategy, strategy1)
                                          .InputTxtFundLabel(1, AddEditFundPage.vintageYear, vintageYear1)
                                          .InputTxtFundLabel(1, AddEditFundPage.strategyHeadquarter, strategyHeadquarter1)
                                          .InputTxtFundLabel(1, AddEditFundPage.assetClass, assetClass1)
                                          .InputTxtFundLabel(1, AddEditFundPage.investmentStage, investmentStage1)
                                          .InputTxtFundLabel(1, AddEditFundPage.industryFocus, industryFocus1)
                                          .InputTxtFundLabel(1, AddEditFundPage.geographicFocus, geographicFocus1)
                                          .InputTxtFundLabel(1, AddEditFundPage.fundSizeM, fundSizeM1.ToString())
                                          .InputTxtFundLabel(1, AddEditFundPage.currency, currency1)
                                          //.InputTxtFundLabel(1, AddEditFundPage.cambridgeVintageYear, vintageYear1)
                                          .ClickAddFundButton(10)
                                          .InputTxtFundLabel(2, AddEditFundPage.fundName, fundName2)
                                          .InputTxtFundLabel(2, AddEditFundPage.strategy, strategy2)
                                          .InputTxtFundLabel(2, AddEditFundPage.vintageYear, vintageYear2)
                                          .InputTxtFundLabel(2, AddEditFundPage.strategyHeadquarter, strategyHeadquarter2)
                                          .InputTxtFundLabel(2, AddEditFundPage.assetClass, assetClass2)
                                          .InputTxtFundLabel(2, AddEditFundPage.investmentStage, investmentStage2)
                                          .InputTxtFundLabel(2, AddEditFundPage.industryFocus, industryFocus2)
                                          .InputTxtFundLabel(2, AddEditFundPage.geographicFocus, geographicFocus2)
                                          .InputTxtFundLabel(2, AddEditFundPage.fundSizeM, fundSizeM2.ToString())
                                          .InputTxtFundLabel(2, AddEditFundPage.currency, currency2)
                                          //.InputTxtFundLabel(2, AddEditFundPage.cambridgeVintageYear, vintageYear2)
                                          .ClickAddFundButton(10)
                                          .InputTxtFundLabel(3, AddEditFundPage.fundName, fundName3)
                                          .InputTxtFundLabel(3, AddEditFundPage.strategy, strategy3)
                                          .InputTxtFundLabel(3, AddEditFundPage.vintageYear, vintageYear3)
                                          .InputTxtFundLabel(3, AddEditFundPage.strategyHeadquarter, strategyHeadquarter3)
                                          .InputTxtFundLabel(3, AddEditFundPage.assetClass, assetClass3)
                                          .InputTxtFundLabel(3, AddEditFundPage.investmentStage, investmentStage3)
                                          .InputTxtFundLabel(3, AddEditFundPage.industryFocus, industryFocus3)
                                          .InputTxtFundLabel(3, AddEditFundPage.geographicFocus, geographicFocus3)
                                          .InputTxtFundLabel(3, AddEditFundPage.fundSizeM, fundSizeM3.ToString())
                                          .InputTxtFundLabel(3, AddEditFundPage.currency, currency3)
                                          //.InputTxtFundLabel(3, AddEditFundPage.cambridgeVintageYear, vintageYear3)
                                          .ClickSaveButton();

                // Check if Loading (Spinner) Icon is displayed then wait for it to disappear
                if (AddEditFundAction.Instance.IsElementPresent(AddEditFundPage.loadingIcon))
                {
                    AddEditFundAction.Instance.WaitForLoadingIconToDisappear(20, AddEditFundPage.loadingIcon);
                }

                // Wait for toastMessage 'Update fund successfully.' to disappear
                AddEditFundAction.Instance.WaitForElementInvisible(10, FundDashboardPage.toastMessage(message));

                // Check if Loading (Spinner) Icon is displayed then wait for it to disappear
                if (AddEditFundAction.Instance.IsElementPresent(AddEditFundPage.loadingIcon))
                {
                    AddEditFundAction.Instance.WaitForLoadingIconToDisappear(20, AddEditFundPage.loadingIcon);
                }
                #endregion
            }
            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Warning
                ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong! Please check console log.");
            }
            #endregion
        }
    }
}