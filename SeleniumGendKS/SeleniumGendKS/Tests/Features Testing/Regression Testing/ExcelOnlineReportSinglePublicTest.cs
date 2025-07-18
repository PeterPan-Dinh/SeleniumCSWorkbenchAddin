using AventStack.ExtentReports;
using NUnit.Framework;
using OpenQA.Selenium;
using SeleniumGendKS.Core.BaseTestCase;
using SeleniumGendKS.Core.ExcelComparision;
using SeleniumGendKS.Core.Selenium;
using SeleniumGendKS.Pages;
using System.Diagnostics;
using System.Reflection;
using System.Threading;
using static System.Collections.Specialized.BitVector32;

namespace SeleniumGendKS.Tests.Features_Testing.Regression_Testing
{
    [TestFixture, Order(6)]
    internal class ExcelOnlineReportSinglePublicTest : BaseTestCase
    {
        // Variables declare
        [Obsolete]
        readonly ExtentReports rep = ExtReportgetInstance();
        private string? summaryTC;
        private string downloadedfileName = LoginPage.downloadedBook;
        ///Determine the folder which Excel files were saved
        private string ExcelFileDownloadedPath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads";
        private string ExcelFileBaselinePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Tests\Documents\Excel files\SinglePublicReport_Baseline\");
        private string? sheetName;
        private const string bookSavedTitleLoadDone = "//*[@data-unique-id='DocumentTitleSaveStatus']";

        [Test, Category("Regression Testing")]
        public void TC001_ExcelOnline_SinglePublicReport_Albourne_NoUserInput()
        {
            #region Variables declare
            string urlInstance = LoginPage.url;
            const string sourceIcon = "A";
            const string dateSection = "Date Section";
            const string customRiskBenchmarkModelling = "Custom Risk Benchmark Modelling";
            const string feeModelSection = "Fee Model Section";
            const string liquiditySection = "Liquidity Section";
            const string msgFundDoesNotHaveData = "This fund does not have Return data. Select another data source or upload the fund data.";
            const string messageCRBM = "This Benchmark does not cover the time range of the report. Select another one.";
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Single Public Report (Fund Albourne - No UserInput) - TC001");
            try
            {
                // Delete Manifest, WorkBook and ExcelCompareDiff file 
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", downloadedfileName);

                // Login Excel Online and then add (manifest.xml) FAD Add-in
                LoginAction.Instance.LoginExcelOnlineSiteAndAddInNoGodaddy(); // LoginExcelOnlineSiteAndAddIn();

                // Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
                string? fileNameBaseline = null;
                if (urlInstance.Contains("sandbox"))
                {
                    fileNameBaseline = @"Book_Melvin_Sandbox.xlsx";
                }
                if (urlInstance.Contains("conceptia"))
                {
                    fileNameBaseline = @"Book_Melvin.xlsx";
                }

                #region Verify alert 'This fund does not have Return data...'
                // Search a Fund - Source = Albourne (that does not have data)
                string managerNameAlbourne = "AQR Capital Management, LLC";
                string fundNameAlbourne = "AQR Absolute Return Credit Strategy"; // AQR Alternative Trends Strategy
                FundDashboardAction.Instance.InputNameToSearchFund(10, "aqr Capital Management", managerNameAlbourne, fundNameAlbourne, sourceIcon);

                // Click to expand User Input (section)
                FADAddInAction.Instance.ClickUserInputSection()
                                       .ClickUserInputSubSection(10, dateSection)
                                       .ClickUserInputSubSection(10, customRiskBenchmarkModelling)
                                       .ClickUserInputSubSection(10, feeModelSection);

                // Verify a red notification is shown at bottom next to the Run button if the fund does not have return data
                verifyPoint = msgFundDoesNotHaveData == FADAddInAction.Instance.ErrorInvalidMessageGetText(10);
                verifyPoints.Add(summaryTC = "A red notification is shown at bottom next to the Run button if the fund does not have return data:\n<br/> '" + msgFundDoesNotHaveData + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                #endregion

                #region Verify red toast error message 'This Benchmark does not cover ...'
                // Search a Fund - Source = Albourne
                managerNameAlbourne = "Citadel Advisors LLC";
                fundNameAlbourne = "Citadel Multi Strategy Funds";
                FundDashboardAction.Instance.InputNameToSearchFund(10, "citadel", managerNameAlbourne, fundNameAlbourne, sourceIcon);

                // Click to expand User Input (section)
                FADAddInAction.Instance.ClickUserInputSubSection(10, dateSection)
                                       .ClickUserInputSubSection(10, customRiskBenchmarkModelling)
                                       .CheckIfExistingCRBMThenDeleteAll(10, sourceIcon)
                                       .ClickUserInputSubSection(10, feeModelSection);

                // Check if A red notification is shown then waiting for that is disappeared
                if (FADAddInAction.Instance.IsElementPresent(FADAddInPage.errorInvalidMessage))
                {
                    FADAddInAction.Instance.WaitForLoadingIconToDisappear(10, FADAddInPage.errorInvalidMessage);
                }

                // Add a crbm that have start date > start date of Fund
                FADAddInAction.Instance.ClickCRBMAddButton(20)
                                       .InputTxtNameCRBMRow(10, 1, "NASDAQ OMX China Technology Net Index (USD)") // start date = 2011
                                       .WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("NASDAQ OMX China Technology Net Index (USD)")) // Old name: S&P U.S. High Yield Corporate Bond Index Total Return
                                       .ClickNameCRBMReturnOfResults(10, "NASDAQ OMX China Technology Net Index (USD)");

                // Verify the error message in red toast is shown
                /// KS-238 The Start Date of the CRBM must be earlier or equal to the start date of the fund (fund return)
                verifyPoint = messageCRBM == FADAddInAction.Instance.ErrorMessageNameCRBMGetTextRow(10, 1);
                verifyPoints.Add(summaryTC = "Verify the error message in red toast is shown: " + messageCRBM + " ", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                #endregion

                #region Run Report for Melvin
                // Search a Fund - Source = Albourne
                managerNameAlbourne = "Melvin Capital Management LP";
                fundNameAlbourne = "Melvin Capital Fund";
                FundDashboardAction.Instance.InputNameToSearchFund(10, "melvin", managerNameAlbourne, fundNameAlbourne, sourceIcon);

                // Click to expand User Input (section)
                FADAddInAction.Instance.ClickUserInputSubSection(10, dateSection)
                                       .ClickUserInputSubSection(10, customRiskBenchmarkModelling)
                                       .CheckIfExistingCRBMThenDeleteAll(10, sourceIcon)
                                       .ClickUserInputSubSection(10, feeModelSection)
                                       .ClickUserInputSubSection(10, liquiditySection);

                // Check if A red notification is shown then waiting for that is disappeared
                if (FADAddInAction.Instance.IsElementPresent(FADAddInPage.errorInvalidMessage))
                {
                    FADAddInAction.Instance.WaitForLoadingIconToDisappear(10, FADAddInPage.errorInvalidMessage);
                }

                // Click on Hurdle Status checkbox to uncheck if it was already checked
                FADAddInAction.Instance.ClickToUnCheckTheCheckbox("Hurdle Status")
                                       /// Liquidity Section 
                                       .ClickAndSelectItemInDropdown(10, "Lockup", "Soft")
                                       .InputTxtLabelField("Lockup Length (months)", "12")
                                       .ClickAndSelectItemInDropdown(10, "Liquidity Frequency", "Semi-Annually")
                                       .InputTxtLabelField("Investor Gate", "100")
                                       .ClickAndSelectItemInDropdown(10, "Sidepocket Probability", "50%")
                                       .InputTxtLabelField("Max % of Sidepocket Permitted", "50"); Thread.Sleep(500);
                FADAddInAction.Instance.WaitForElementInvisible(10, FADAddInPage.errorInvalidMessageContent("All the liquidity parameters need to be populated except for Investor Gate"));

                // Click on Run button
                FADAddInAction.Instance.ClickRunButton(10)
                                       .WaitForLoadingIconToDisappear(60, FundDashboardPage.loadingIcon)
                                       .WaitforExcelWebRenderGraphDataAtLastSheet(60); Thread.Sleep(5000); //.WaitforExcelWebRenderDone2(60);

                // Check if the status "Saving.../Saved" is displayed then wait for it shows
                if (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone))) 
                {
                    FADAddInAction.Instance.WaitForElementVisible(10, By.XPath(bookSavedTitleLoadDone));
                }

                // Check if the status "Saving.../Saved" is displayed then refresh page, relogin re-enter data
                int timeout = 0;
                while(FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == true && timeout < 6)
                {
                    if (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == true && timeout == 4)
                    {
                        Driver.Browser.Navigate().Refresh();
                        LoginAction.Instance.LoginExcelOnlineSiteAfterRefresh();

                        // re enter data to run report
                        // Search a Fund - Source = Albourne
                        managerNameAlbourne = "Melvin Capital Management LP";
                        fundNameAlbourne = "Melvin Capital Fund";
                        FundDashboardAction.Instance.InputNameToSearchFund(10, "melvin", managerNameAlbourne, fundNameAlbourne, sourceIcon);

                        // Click to expand User Input (section)
                        FADAddInAction.Instance.ClickUserInputSection();
                        FADAddInAction.Instance.ClickUserInputSubSection(10, dateSection);
                        FADAddInAction.Instance.ClickUserInputSubSection(10, customRiskBenchmarkModelling);
                        FADAddInAction.Instance.CheckIfExistingCRBMThenDeleteAll(10, sourceIcon);
                        FADAddInAction.Instance.ClickUserInputSubSection(10, feeModelSection);
                        FADAddInAction.Instance.ClickUserInputSubSection(10, liquiditySection);

                        // Click on Hurdle Status checkbox to uncheck if it was already checked
                        FADAddInAction.Instance.ClickToUnCheckTheCheckbox("Hurdle Status")
                                               /// Liquidity Section 
                                               .ClickAndSelectItemInDropdown(10, "Lockup", "Soft")
                                               .InputTxtLabelField("Lockup Length (months)", "12")
                                               .ClickAndSelectItemInDropdown(10, "Liquidity Frequency", "Semi-Annually")
                                               .InputTxtLabelField("Investor Gate", "100")
                                               .ClickAndSelectItemInDropdown(10, "Sidepocket Probability", "50%")
                                               .InputTxtLabelField("Max % of Sidepocket Permitted", "50"); Thread.Sleep(500);
                        FADAddInAction.Instance.WaitForElementInvisible(10, FADAddInPage.errorInvalidMessageContent("All the liquidity parameters need to be populated except for Investor Gate"));

                        // Click on Run button
                        FADAddInAction.Instance.ClickRunButton(10)
                                               .WaitForLoadingIconToDisappear(60, FundDashboardPage.loadingIcon)
                                               .WaitforExcelWebRenderDone2(60);
                        break;
                    }
                    if (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == false)
                    {
                        Thread.Sleep(11000);

                        // Check if the dialog "Invalid sheet" is displayed then click on OK button
                        if (FADAddInAction.Instance.IsElementPresent(NavigationPage.okInvalidSheetButtonExcelOnline))
                        {
                            NavigationAction.Instance.ClicKOkButtonInvalidSheetExcelOnline(10); Thread.Sleep(1000);
                        }
                        Thread.Sleep(1000);
                        break;
                    }
                    if (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == true && timeout == 5)
                        {
                            Console.WriteLine("timeout - Excel WebRender fail to load done!!!");
                            break;
                        }
                    Thread.Sleep(1000);
                    timeout ++;
                }
                
                // Download Excel Online Report
                NavigationAction.Instance.DownloadExcelOnlineReport(10);

                // Verify Excel file was downloaded successfully
                verifyPoint = LoginAction.Instance.CheckFileDownloadIsComplete(9, ExcelFileDownloadedPath, downloadedfileName);
                verifyPoints.Add(summaryTC = "Verify the '" + downloadedfileName + "' file is downloaded successful", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Paste value for Downloaded File
                System.Threading.Thread.Sleep(2000);
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "Raw Data", "B2:F500", 2, "#,##0.00"); // (if 6 digits --> 6, "#,######0.000000")
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B16:C18", 2, "#,##0.00"); // 'Workbench'-Auto Score 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B21:J29", 2, "#,##0.00"); // 'Workbench'-Manager Alpha 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B32:K49", 2, "#,##0.00"); // 'Workbench'-Metrics Over Selected Time Frames 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "C53:D75", 2, "#,##0.00"); // 'Workbench'-Annualized Return 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "C78:D83", 2, "#,##0.00"); // 'Workbench'-Annualized Return 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B86:M94", 2, "#,##0.00"); // 'Workbench'-Net Monthly Return 2 digits

                //Verify content of the downloaded file by doing compare with the baseline excel files.
                sheetName = "Raw Data";
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline, sheetName, fileNameBaseline);
                verifyPoints.Add(summaryTC = "Verify the information of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundNameAlbourne +")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Verify Excel format of the downloaded file by doing compare with the baseline excel files.
                sheetName = "WorkBench";
                verifyPoint = ExcelComparision.IsTheFormatSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline, sheetName, fileNameBaseline);
                verifyPoints.Add(summaryTC = "Verify Excel Format of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundNameAlbourne + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Verify the information of Single Public Report Sheet 'WorkBench'
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline, sheetName, fileNameBaseline);
                verifyPoints.Add(summaryTC = "Verify the information of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundNameAlbourne + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                #endregion

                //Closing Excel Application Process
                foreach (var application in Process.GetProcessesByName("EXCEL"))
                {
                    application.Kill(); System.Threading.Thread.Sleep(1000);
                }

                // Delete Manifest and WorkBook file
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", downloadedfileName);
            }
            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                //Closing Excel Application Process
                foreach (var application in Process.GetProcessesByName("EXCEL"))
                {
                    application.Kill(); System.Threading.Thread.Sleep(1000);
                }
                
                // Delete Manifest, WorkBook and ExcelCompareDiff file 
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", downloadedfileName);

                // Warning
                ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong! Please check console log.");
            }
            #endregion
        }


        [Test, Category("Regression Testing")]
        public void TC002_ExcelOnline_SinglePublicReport_Albourne_UserInput()
        {
            #region Variables declare
            string urlInstance = LoginPage.url;
            const string sourceIcon = "A";
            const string dateSection = "Date Section";
            const string customRiskBenchmarkModelling = "Custom Risk Benchmark Modelling";
            const string feeModelSection = "Fee Model Section";
            const string liquiditySection = "Liquidity Section";
            const string messageCRBM = "This Benchmark does not cover the time range of the report. Select another one.";
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Single Public Report (Fund Albourne - UserInput) - TC002");
            try
            {
                // Delete Manifest, WorkBook and ExcelCompareDiff file 
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", downloadedfileName);

                // Login Excel Online and then add (manifest.xml) FAD Add-in
                LoginAction.Instance.LoginExcelOnlineSiteAndAddInNoGodaddy();

                // Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
                string? fileNameBaseline = null;
                if (urlInstance.Contains("sandbox"))
                {
                    fileNameBaseline = @"Book_Melvin_Input_Sandbox.xlsx";
                }
                if (urlInstance.Contains("conceptia"))
                {
                    fileNameBaseline = @"Book_Melvin_Input.xlsx";
                }

                // Search a Fund - Source = Albourne
                string managerNameAlbourne = "Melvin Capital Management LP";
                string fundNameAlbourne = "Melvin Capital Fund";
                FundDashboardAction.Instance.InputNameToSearchFund(30, "melvin", managerNameAlbourne, fundNameAlbourne, sourceIcon);

                // Click to expand User Input (section)
                FADAddInAction.Instance.ClickUserInputSection()
                                       .ClickUserInputSubSection(10, dateSection)
                                       .ClickUserInputSubSection(10, customRiskBenchmarkModelling)
                                       .ClickUserInputSubSection(10, feeModelSection)
                                       .ClickUserInputSubSection(10, liquiditySection);

                #region Input Data for (User Input)
                // User Input
                FADAddInAction.Instance.InputTxtLabelField("Tracking Error", "11");

                // Custom Risk Benchmark Modelling
                // Add the 1st CRBM
                FADAddInAction.Instance.ClickCRBMAddButton(10)
                                       .InputTxtNameCRBMRow(10, 1, "s&p")
                                       .WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("S&P 500 Index"))
                                       .ClickNameCRBMReturnOfResults(10, "S&P 500 Index")
                                       .InputNumberBetaCRBMRow(10, 1, "1.5")
                                       .InputNumberGrossExposureCRBMRow(10, 1, "-2.5");

                /*// Add the 2nd CRBM
                //FADAddInAction.Instance.PageDownToScrollDownPage();  // --> Apply when run auton on Remote QA-PC 
                FADAddInAction.Instance.ClickCRBMAddButton(10)
                                       .InputTxtNameCRBMRow(10, 2, "S&P US")
                                       .WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("S&P US High Yield Corporate Bond TR Index")) // Old name: S&P U.S. High Yield Corporate Bond Index Total Return
                                       .ClickNameCRBMReturnOfResults(10, "S&P US High Yield Corporate Bond TR Index");

                // Verify the error message in red toast is shown
                /// KS-238 The Start Date of the CRBM must be earlier or equal to the start date of the fund (fund return)
                verifyPoint = messageCRBM == FADAddInAction.Instance.ErrorMessageNameCRBMGetTextRow(10, 2);
                verifyPoints.Add(summaryTC = "Verify the error message in red toast is shown: " + messageCRBM + " ", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click delete icon of the 2nd CRBM
                FADAddInAction.Instance.ClickCRBMDeleteButton(10, 2)
                                       .WaitForElementInvisible(10, FADAddInPage.nameCRBMInputTxtRow(2)); */

                // Check if the red error message is shown then verify it
                string msg = "All the liquidity parameters need to be populated except for Investor Gate";
                if (FADAddInAction.Instance.IsElementPresent(FADAddInPage.errorInvalidMessageContent(msg)))
                {
                    verifyPoint = msg == FADAddInAction.Instance.ErrorInvalidMessageGetText(10);
                    verifyPoints.Add(summaryTC = "Verify the error message in red text is shown: " + msg + " ", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);
                }

                // Add the 2nd CRBM (another valid)
                FADAddInAction.Instance.ClickCRBMAddButton(10)
                                       .InputTxtNameCRBMRow(10, 2, "msci")
                                       .WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("MSCI World Net TR Index (LCL)")) // old: MSCI Europe Growth Net TR Index (LCL) (factset)
                                       .ClickNameCRBMReturnOfResults(10, "MSCI World Net TR Index (LCL)")
                                       .InputNumberBetaCRBMRow(10, 2, "-3.5")
                                       .InputNumberGrossExposureCRBMRow(10, 2, "4.5");

                // Fee Model Section
                FADAddInAction.Instance.InputTxtLabelField("Management Fee", "5")
                                       .ClickAndSelectItemInDropdown(10, "Management Fee Paid", "Quarterly")
                                       .InputTxtLabelField("Performance Fee", "21")
                                       .ClickToCheckTheCheckbox("High Water Mark")
                                       .ClickAndSelectItemInDropdown(10, "Catch Up", "Yes")
                                       .InputTxtLabelField("Catch Up %-age (if Soft)", "11") // Only enable when "Catch Up" = Yes
                                       .ClickAndSelectItemInDropdown(10, "Crystallization Every X Years", "10")
                /// Hurdle Status
                                       .ClickToCheckTheCheckbox("Hurdle Status")
                                       .ClickAndSelectItemInDropdown(10, "Hurdle Fixed or Relative", "Relative") // Only enable when Hurdle Status is checked (checkbox)
                                                                                                                 //InputTxtLabelField("Hurdle Rate (%)", "12"); // Only enable when Hurdle Fixed or Relative = Fixed
                                       .InputTxtLabelField("Fixed Percentage on Top of Index Return (%)", "40")
                                       //.InputTxtSearchLabelField("Hurdle Benchmark", "msci") // (new: KS-834/838) (old: Only enable when Hurdle Fixed or Relative = Relative)
                                       .ClickAddButtonInXTable(10, "hurdle-section")
                                       .InputTxtNameBenchmarkXTableRow(10, "hurdle-section", "1", "MSCI World Net TR Index (LCL)")
                                       .WaitForElementVisible(30, FADAddInPage.nameCRBMReturnOfResults("MSCI World Net TR Index (LCL)")) // old: MSCI Emerging Markets Latin America ex Brazil Net TR Index (USD)
                                       .ClickNameCRBMReturnOfResults(10, "MSCI World Net TR Index (LCL)")
                                       .InputTxtExposureBenchmarkXTableRow(10, "hurdle-section", "1", "55.7")
                                       .ClickAndSelectItemInDropdown(10, "Hurdle Type", "Soft Hurdle")
                                       .ClickAndSelectItemInDropdown(10, "Ramp Type", "Performance Dependent")
                /// Liquidity Section 
                                       .ClickAndSelectItemInDropdown(10, "Lockup", "Soft")
                                       .InputTxtLabelField("Lockup Length (months)", "3")
                                       .ClickAndSelectItemInDropdown(10, "Liquidity Frequency", "Daily")
                                       .InputTxtLabelField("Investor Gate", "40")
                                       .ClickAndSelectItemInDropdown(10, "Sidepocket Probability", "50%")
                                       .InputTxtLabelField("Max % of Sidepocket Permitted", "30"); Thread.Sleep(500);

                // wait for the red error message to disappear
                FADAddInAction.Instance.WaitForElementInvisible(10, FADAddInPage.errorInvalidMessage);
                #endregion

                // Click on Run button
                FADAddInAction.Instance.ClickRunButton(10)
                                       .WaitForLoadingIconToDisappear(60, FundDashboardPage.loadingIcon)
                                       .WaitforExcelWebRenderGraphDataAtLastSheet(60); Thread.Sleep(5000); //.WaitforExcelWebRenderDone2(60);

                // Check if the status "Saving.../Saved" is displayed then refresh page, relogin re-enter data
                int timeout = 0;
                while (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == true && timeout < 6)
                {
                    if (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == true && timeout == 4)
                    {
                        Driver.Browser.Navigate().Refresh();
                        LoginAction.Instance.LoginExcelOnlineSiteAfterRefresh();

                        // re enter data to run report
                        // Search a Fund - Source = Albourne
                        managerNameAlbourne = "Melvin Capital Management LP";
                        fundNameAlbourne = "Melvin Capital Fund";
                        FundDashboardAction.Instance.InputNameToSearchFund(10, "melvin", managerNameAlbourne, fundNameAlbourne, sourceIcon);

                        // Click to expand User Input (section)
                        FADAddInAction.Instance.ClickUserInputSection();
                        FADAddInAction.Instance.ClickUserInputSubSection(10, dateSection);
                        FADAddInAction.Instance.ClickUserInputSubSection(10, customRiskBenchmarkModelling);
                        FADAddInAction.Instance.CheckIfExistingCRBMThenDeleteAll(10, sourceIcon);
                        FADAddInAction.Instance.ClickUserInputSubSection(10, feeModelSection);
                        FADAddInAction.Instance.ClickUserInputSubSection(10, liquiditySection);

                        #region Input Data for (User Input)
                        // User Input
                        FADAddInAction.Instance.InputTxtLabelField("Tracking Error", "11");

                        // Custom Risk Benchmark Modelling
                        // Add the 1st CRBM
                        FADAddInAction.Instance.ClickCRBMAddButton(10)
                                               .InputTxtNameCRBMRow(10, 1, "s&p")
                                               .WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("S&P 500 Index"))
                                               .ClickNameCRBMReturnOfResults(10, "S&P 500 Index")
                                               .InputNumberBetaCRBMRow(10, 1, "1.5")
                                               .InputNumberGrossExposureCRBMRow(10, 1, "-2.5");

                        // Add the 2nd CRBM (another valid)
                        FADAddInAction.Instance.ClickCRBMAddButton(10)
                                               .InputTxtNameCRBMRow(10, 2, "msci")
                                               .WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("MSCI World Net TR Index (LCL)")) // old: MSCI Europe Growth Net TR Index (LCL) (factset)
                                               .ClickNameCRBMReturnOfResults(10, "MSCI World Net TR Index (LCL)")
                                               .InputNumberBetaCRBMRow(10, 2, "-3.5")
                                               .InputNumberGrossExposureCRBMRow(10, 2, "4.5");

                        // Fee Model Section
                        FADAddInAction.Instance.InputTxtLabelField("Management Fee", "5")
                                               .ClickAndSelectItemInDropdown(10, "Management Fee Paid", "Quarterly")
                                               .InputTxtLabelField("Performance Fee", "21")
                                               .ClickToCheckTheCheckbox("High Water Mark")
                                               .ClickAndSelectItemInDropdown(10, "Catch Up", "Yes")
                                               .InputTxtLabelField("Catch Up %-age (if Soft)", "11") // Only enable when "Catch Up" = Yes
                                               .ClickAndSelectItemInDropdown(10, "Crystallization Every X Years", "10")
                                               /// Hurdle Status
                                               .ClickToCheckTheCheckbox("Hurdle Status")
                                               .ClickAndSelectItemInDropdown(10, "Hurdle Fixed or Relative", "Relative") // Only enable when Hurdle Status is checked (checkbox)
                                                                                                                         //InputTxtLabelField("Hurdle Rate (%)", "12"); // Only enable when Hurdle Fixed or Relative = Fixed
                                               .InputTxtLabelField("Fixed Percentage on Top of Index Return (%)", "40")
                                               //.InputTxtSearchLabelField("Hurdle Benchmark", "msci") // (new: KS-834/838) (old: Only enable when Hurdle Fixed or Relative = Relative)
                                               .ClickAddButtonInXTable(10, "hurdle-section")
                                               .InputTxtNameBenchmarkXTableRow(10, "hurdle-section", "1", "MSCI World Net TR Index (LCL)")
                                               .WaitForElementVisible(30, FADAddInPage.nameCRBMReturnOfResults("MSCI World Net TR Index (LCL)")) // old: MSCI Emerging Markets Latin America ex Brazil Net TR Index (USD)
                                               .ClickNameCRBMReturnOfResults(10, "MSCI World Net TR Index (LCL)")
                                               .InputTxtExposureBenchmarkXTableRow(10, "hurdle-section", "1", "55.7")
                                               .ClickAndSelectItemInDropdown(10, "Hurdle Type", "Soft Hurdle")
                                               .ClickAndSelectItemInDropdown(10, "Ramp Type", "Performance Dependent")
                                               /// Liquidity Section 
                                               .ClickAndSelectItemInDropdown(10, "Lockup", "Soft")
                                               .InputTxtLabelField("Lockup Length (months)", "3")
                                               .ClickAndSelectItemInDropdown(10, "Liquidity Frequency", "Daily")
                                               .InputTxtLabelField("Investor Gate", "40")
                                               .ClickAndSelectItemInDropdown(10, "Sidepocket Probability", "50%")
                                               .InputTxtLabelField("Max % of Sidepocket Permitted", "30"); Thread.Sleep(500);

                        // wait for the red error message to disappear
                        FADAddInAction.Instance.WaitForElementInvisible(10, FADAddInPage.errorInvalidMessage);
                        #endregion

                        // Click on Run button
                        FADAddInAction.Instance.ClickRunButton(10)
                                               .WaitForLoadingIconToDisappear(60, FundDashboardPage.loadingIcon)
                                               .WaitforExcelWebRenderDone2(60);
                        break;
                    }
                    if (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == false)
                    {
                        Thread.Sleep(11000);

                        // Check if the dialog "Invalid sheet" is displayed then click on OK button
                        if (FADAddInAction.Instance.IsElementPresent(NavigationPage.okInvalidSheetButtonExcelOnline))
                        {
                            NavigationAction.Instance.ClicKOkButtonInvalidSheetExcelOnline(10); Thread.Sleep(1000);
                        }
                        Thread.Sleep(1000);
                        break;
                    }
                    if (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == true && timeout == 5)
                    {
                        Console.WriteLine("timeout - Excel WebRender fail to load done!!!");
                        break;
                    }
                    Thread.Sleep(1000);
                    timeout++;
                }

                // Download Excel Online Report
                NavigationAction.Instance.DownloadExcelOnlineReport(10);

                // Verify Excel file was downloaded successfully
                verifyPoint = LoginAction.Instance.CheckFileDownloadIsComplete(9, ExcelFileDownloadedPath, downloadedfileName);
                verifyPoints.Add(summaryTC = "Verify the '" + downloadedfileName + "' file is downloaded successful", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Paste value for Downloaded File
                System.Threading.Thread.Sleep(2000);
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "Raw Data", "B2:H500", 2, "#,##0.00"); // 'Raw Data' 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B18:C20", 2, "#,##0.00"); // 'Workbench'-Auto Score 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B23:J31", 2, "#,##0.00"); // 'Workbench'-Manager Alpha 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B34:K51", 2, "#,##0.00"); // 'Workbench'-Metrics Over Selected Time Frames 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "C55:F77", 2, "#,##0.00"); // 'Workbench'-Annualized Return 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "C80:F85", 2, "#,##0.00"); // 'Workbench'-Annualized Return 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B88:M96", 2, "#,##0.00"); // 'Workbench'-Net Monthly Return 2 digits

                // Verify content of the downloaded file by doing compare with the baseline excel files
                sheetName = "Raw Data";
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline, sheetName, fileNameBaseline);
                verifyPoints.Add(summaryTC = "Verify the information of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundNameAlbourne + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Verify the information of Single Public Report Sheet 'WorkBench'
                sheetName = "WorkBench";
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline, sheetName, fileNameBaseline);
                verifyPoints.Add(summaryTC = "Verify the information of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundNameAlbourne + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                #region Run Report for Oasis Japan Strategic Fund
                /* No need to run report for Oasis Japan Strategic Fund
                // Delete WorkBook files
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "Book");

                // (Oasis Japan Strategic Fund) Single Public Report
                managerNameAlbourne = "Oasis Management (Hong Kong)";
                fundNameAlbourne = "Oasis Japan Strategic Fund";

                // Switch to the FAD Add-in frame (to interact with Workbench for Excel Online)
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameIdFADAddInCurrentInstance(LoginPage.instanceName));

                // Search a Fund - Source = Albourne
                FundDashboardAction.Instance.InputNameToSearchFund(10, "oasis", managerNameAlbourne, fundNameAlbourne);

                // Click to expand User Input (section)
                FADAddInAction.Instance.ClickUserInputSubSection(10, dateSection)
                                       .ClickUserInputSubSection(10, customRiskBenchmarkModelling)
                                       .ClickUserInputSubSection(10, feeModelSection);

                #region Input Data for (User Input)
                // User Input
                FADAddInAction.Instance.InputTxtLabelField("Tracking Error", "9");

                // Custom Risk Benchmark Modelling
                // Delete 2 CRBM previous
                FADAddInAction.Instance.ClickCRBMDeleteButton(10, 1).ClickCRBMDeleteButton(10, 1);

                // Add the 1st CRBM
                FADAddInAction.Instance.ClickCRBMAddButton(10)
                                       .InputTxtNameCRBMRow(10, 1, "s&p")
                                       .WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("S&P 500 Industrials (Sector) Net Total Return Index"))
                                       .ClickNameCRBMReturnOfResults(10, "S&P 500 Industrials (Sector) Net Total Return Index")
                                       .InputNumberBetaCRBMRow(10, 1, "-0.5")
                                       .InputNumberGrossExposureCRBMRow(10, 1, "-1.5");
                // Fee Model Section
                FADAddInAction.Instance.InputTxtLabelField("Management Fee", "3.25")
                                       .ClickAndSelectItemInDropdown(10, "Management Fee Paid", "Monthly")
                                       .InputTxtLabelField("Performance Fee", "14")
                                       .ClickToCheckTheCheckbox("High Water Mark")
                                       .ClickAndSelectItemInDropdown(10, "Catch Up", "No")
                                       //.InputTxtLabelField("Catch Up %-age (if Soft)", "11") // Only enable when "Catch Up" = Yes
                                       .ClickAndSelectItemInDropdown(10, "Crystallization Every X Years", "9")
                /// Hurdle Status
                                       .ClickToCheckTheCheckbox("Hurdle Status")
                                       .ClickAndSelectItemInDropdown(10, "Hurdle Fixed or Relative", "Fixed") // Only enable when Hurdle Status is checked (checkbox)
                                       .InputTxtLabelField("Hurdle Rate (%)", "11") // Only enable when Hurdle Fixed or Relative = Fixed
                                       //.InputTxtSearchLabelField("Hurdle Benchmark", "msci") // Only enable when Hurdle Fixed or Relative = Relative
                                       //.WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("MSCI EM LATIN AMERICA ex BRAZIL Net Return USD Index"))
                                       //.ClickNameCRBMReturnOfResults(10, "MSCI EM LATIN AMERICA ex BRAZIL Net Return USD Index")
                .ClickAndSelectItemInDropdown(10, "Hurdle Type", "Hard Hurdle")
                .ClickAndSelectItemInDropdown(10, "Ramp Type", "NAV Dependent");
                #endregion

                // Click on Run button
                FADAddInAction.Instance.ClickRunButton(10)
                                       .WaitForLoadingIconToDisappear(60, FundDashboardPage.loadingIcon)
                                       .WaitforExcelWebRenderDone(60);

                // Download Excel Online Report
                NavigationAction.Instance.DownloadExcelOnlineReport(10);

                // Verify Excel file was downloaded successfully
                verifyPoint = LoginAction.Instance.CheckFileDownloadIsComplete(9, ExcelFileDownloadedPath, "Book.xlsx");
                verifyPoints.Add(summaryTC = "Verify the Book.xlsx (" + fundNameAlbourne + ") file is downloaded successful", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Verify content of the downloaded file by doing compare with the baseline excel files
                fileNameBaseline = @"Book_OasisJapan_Input.xlsx";
                sheetName = "Raw Data";
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + @"Book.xlsx", ExcelFileBaselinePath + fileNameBaseline, sheetName);
                verifyPoints.Add(summaryTC = "Verify the information of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundNameAlbourne + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                sheetName = "WorkBench";
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + @"Book.xlsx", ExcelFileBaselinePath + fileNameBaseline, sheetName);
                verifyPoints.Add(summaryTC = "Verify the information of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundNameAlbourne + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                */
                #endregion

                //Closing Excel Application Process
                foreach (var application in Process.GetProcessesByName("EXCEL"))
                {
                    application.Kill(); System.Threading.Thread.Sleep(1000);
                }

                // Delete Manifest and WorkBook file
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", downloadedfileName);
            }
            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                //Closing Excel Application Process
                foreach (var application in Process.GetProcessesByName("EXCEL"))
                {
                    application.Kill(); System.Threading.Thread.Sleep(1000);
                }

                // Delete Manifest and WorkBook file
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", downloadedfileName);

                // Warning
                ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong! Please check console log.");
            }
            #endregion
        }


        [Test, Category("Regression Testing")]
        public void TC003_ExcelOnline_SinglePublicReport_Manual_NoUserInput()
        {
            #region Variables declare
            string urlInstance = LoginPage.url;
            const string sourceIcon = "M";
            const string dateSection = "Date Section";
            const string customRiskBenchmarkModelling = "Custom Risk Benchmark Modelling";
            const string feeModelSection = "Fee Model Section";
            const string liquiditySection = "Liquidity Section";
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Single Public Report (Fund Manual - No UserInput) - TC003");
            try
            {
                // Delete Manifest, WorkBook and ExcelCompareDiff file 
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", downloadedfileName);

                // Login Excel Online and then add (manifest.xml) FAD Add-in
                LoginAction.Instance.LoginExcelOnlineSiteAndAddInNoGodaddy();

                // Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
                string? fileNameBaseline = null;
                string? fileNameBaseline2 = null;
                string? instance = null;
                if (urlInstance.Contains("sandbox"))
                {
                    instance = "sandbox";
                    fileNameBaseline = @"Book_PreludeManual_Sandbox.xlsx";
                    fileNameBaseline2 = @"Book_GoldenChinaManual_Sandbox.xlsx";
                }
                if (urlInstance.Contains("conceptia"))
                {
                    instance = "conceptia";
                    fileNameBaseline = @"Book_PreludeManual.xlsx";
                    fileNameBaseline2 = @"Book_GoldenChinaManual.xlsx";
                }

                #region (Prelude Capital - Prelude Structured Alternatives Fund) Single Public Report
                // Search a Fund - Source = Albourne
                string managerNameManual = "Prelude Capital";
                string fundNameManual = "Prelude Structured Alternatives Fund";
                FundDashboardAction.Instance.InputNameToSearchFund(10, "prelude", managerNameManual, fundNameManual, sourceIcon);

                // Click to expand User Input (section)
                FADAddInAction.Instance.ClickUserInputSection()
                                       .ClickUserInputSubSection(10, dateSection)
                                       .ClickUserInputSubSection(10, customRiskBenchmarkModelling)
                                       .ClickUserInputSubSection(10, feeModelSection)
                                       .ClickUserInputSubSection(10, liquiditySection);

                // Check if A red notification is shown then waiting for that is disappeared
                if (FADAddInAction.Instance.IsElementPresent(FADAddInPage.errorInvalidMessage))
                {
                    FADAddInAction.Instance.WaitForLoadingIconToDisappear(10, FADAddInPage.errorInvalidMessage);
                }

                /// Liquidity Section 
                FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, "Lockup", "None")
                                       .InputTxtLabelField("Lockup Length (months)", "12")
                                       .ClickAndSelectItemInDropdown(10, "Liquidity Frequency", "Monthly")
                                       .InputTxtLabelField("Investor Gate", "")
                                       .ClickAndSelectItemInDropdown(10, "Sidepocket Probability", "50%")
                                       .InputTxtLabelField("Max % of Sidepocket Permitted", "50"); Thread.Sleep(500);
                FADAddInAction.Instance.WaitForElementInvisible(10, FADAddInPage.errorInvalidMessageContent("All the liquidity parameters need to be populated except for Investor Gate"));

                // Click on Run button
                FADAddInAction.Instance.ClickRunButton(10); System.Windows.Forms.SendKeys.SendWait(@"{ENTER}");
                FADAddInAction.Instance.ClickRunButton(10)
                                       .WaitForLoadingIconToDisappear(60, FundDashboardPage.loadingIcon)
                                       .WaitforExcelWebRenderGraphDataAtLastSheet(60); Thread.Sleep(5000); //.WaitforExcelWebRenderDone2(60);

                // Check if the status "Saving.../Saved" is displayed then refresh page, relogin re-enter data
                int timeout = 0;
                while (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == true && timeout < 6)
                {
                    if (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == true && timeout == 4)
                    {
                        Driver.Browser.Navigate().Refresh();
                        LoginAction.Instance.LoginExcelOnlineSiteAfterRefresh();

                        // re enter data to run report
                        // Search a Fund - Source = Albourne
                        managerNameManual = "Prelude Capital";
                        fundNameManual = "Prelude Structured Alternatives Fund";
                        FundDashboardAction.Instance.InputNameToSearchFund(10, "prelude", managerNameManual, fundNameManual, sourceIcon);

                        // Click to expand User Input (section)
                        FADAddInAction.Instance.ClickUserInputSection();
                        FADAddInAction.Instance.ClickUserInputSubSection(10, dateSection);
                        FADAddInAction.Instance.ClickUserInputSubSection(10, customRiskBenchmarkModelling);
                        FADAddInAction.Instance.CheckIfExistingCRBMThenDeleteAll(10, sourceIcon);
                        FADAddInAction.Instance.ClickUserInputSubSection(10, feeModelSection);
                        FADAddInAction.Instance.ClickUserInputSubSection(10, liquiditySection);

                        /// Liquidity Section 
                        FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, "Lockup", "None")
                                               .InputTxtLabelField("Lockup Length (months)", "12")
                                               .ClickAndSelectItemInDropdown(10, "Liquidity Frequency", "Monthly")
                                               .InputTxtLabelField("Investor Gate", "")
                                               .ClickAndSelectItemInDropdown(10, "Sidepocket Probability", "50%")
                                               .InputTxtLabelField("Max % of Sidepocket Permitted", "50"); Thread.Sleep(500);
                        FADAddInAction.Instance.WaitForElementInvisible(10, FADAddInPage.errorInvalidMessageContent("All the liquidity parameters need to be populated except for Investor Gate"));

                        // Click on Run button
                        FADAddInAction.Instance.ClickRunButton(10)
                                               .WaitForLoadingIconToDisappear(60, FundDashboardPage.loadingIcon)
                                               .WaitforExcelWebRenderDone2(60);
                        break;
                    }
                    if (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == false)
                    {
                        Thread.Sleep(11000);

                        // Check if the dialog "Invalid sheet" is displayed then click on OK button
                        if (FADAddInAction.Instance.IsElementPresent(NavigationPage.okInvalidSheetButtonExcelOnline))
                        {
                            NavigationAction.Instance.ClicKOkButtonInvalidSheetExcelOnline(10); Thread.Sleep(1000);
                        }
                        Thread.Sleep(1000);
                        break;
                    }
                    if (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == true && timeout == 5)
                    {
                        Console.WriteLine("timeout - Excel WebRender fail to load done!!!");
                        break;
                    }
                    Thread.Sleep(1000);
                    timeout++;
                }

                // Download Excel Online Report
                NavigationAction.Instance.DownloadExcelOnlineReport(10);

                // Verify Excel file was downloaded successfully
                verifyPoint = LoginAction.Instance.CheckFileDownloadIsComplete(9, ExcelFileDownloadedPath, downloadedfileName);
                verifyPoints.Add(summaryTC = "Verify the '" + downloadedfileName + "' file is downloaded successful", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Paste value for the downloaded file
                System.Threading.Thread.Sleep(2000);
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "Raw Data", "B2:F500", 6, "#,######0.000000"); // 'Raw Data' 6 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B16:C18", 2, "#,##0.00"); // 'Workbench'-Auto Score 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B21:J32", 2, "#,##0.00"); // 'Workbench'-Manager Alpha 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B35:K52", 2, "#,##0.00"); // 'Workbench'-Metrics Over Selected Time Frames 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "C56:D80", 2, "#,##0.00"); // 'Workbench'-Annualized Return 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "C83:D88", 2, "#,##0.00"); // 'Workbench'-Annualized Return 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B91:M102", 2, "#,##0.00"); // 'Workbench'-Net Monthly Return 2 digits

                //Verify content of the downloaded file by doing compare with the baseline excel files.
                sheetName = "Raw Data";
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline, sheetName, fileNameBaseline);
                verifyPoints.Add(summaryTC = "Verify the information of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundNameManual + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                //// Verify Excel format of the downloaded file by doing compare with the baseline excel files.
                //sheetName = "WorkBench";
                //verifyPoint = ExcelComparision.IsTheFormatSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline, sheetName, fileNameBaseline);
                //verifyPoints.Add(summaryTC = "Verify Excel Format of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundNameManual + ")", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                // Verify the information of Single Public Report Sheet 'WorkBench'
                sheetName = "WorkBench";
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline, sheetName, fileNameBaseline);
                verifyPoints.Add(summaryTC = "Verify the information of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundNameManual + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Delete WorkBook file
                LoginAction.Instance.DeleteAllFilesWorkBookExcel();
                #endregion

                #region (Greenwoods - Golden China Fund) Single Public Report
                /*
                managerNameManual = "Greenwoods";
                fundNameManual = "Golden China Fund";

                // Switch to the FAD Add-in frame (to interact with Workbench for Excel Online)
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameIdFADAddInCurrentInstance(instance));

                // Search a Fund - Source = Albourne
                FADAddInAction.Instance.HomeToScrollUpPage();
                FundDashboardAction.Instance.InputNameToSearchFund(10, "greenwoods", managerNameManual, fundNameManual, sourceIcon);

                // Click to expand User Input (section)
                FADAddInAction.Instance//.WaitForLoadingIconToDisappear(10, FADAddInPage.MessageFundDoesNotHaveReturnData)
                                       .ClickUserInputSubSection(10, dateSection)
                                       .ClickUserInputSubSection(10, customRiskBenchmarkModelling)
                                       .ClickUserInputSubSection(10, feeModelSection)
                                       .ClickUserInputSubSection(10, liquiditySection);

                // Check if A red notification is shown then waiting for that is disappeared
                System.Threading.Thread.Sleep(1000);
                if (FADAddInAction.Instance.IsElementPresent(FADAddInPage.errorInvalidMessage))
                {
                    FADAddInAction.Instance.WaitForLoadingIconToDisappear(10, FADAddInPage.errorInvalidMessage);
                }

                // Click on Run button
                FADAddInAction.Instance.ClickRunButton(10); System.Windows.Forms.SendKeys.SendWait(@"{ENTER}");
                FADAddInAction.Instance.ClickRunButton(10)
                                       .WaitForLoadingIconToDisappear(60, FundDashboardPage.loadingIcon)
                                       .WaitforExcelWebRenderDone2(60); Thread.Sleep(1000);

                // Download Excel Online Report
                NavigationAction.Instance.DownloadExcelOnlineReport(10);

                // Verify Excel file was downloaded successfully
                verifyPoint = LoginAction.Instance.CheckFileDownloadIsComplete(9, ExcelFileDownloadedPath, downloadedfileName);
                verifyPoints.Add(summaryTC = "Verify the '" + downloadedfileName + "' (" + fundNameManual + ") file is downloaded successful", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Paste value for the downloaded file
                System.Threading.Thread.Sleep(2000);
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "Raw Data", "B2:E500", 6, "#,######0.000000"); // 'Raw Data' 6 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B16:C18", 2, "#,##0.00"); // 'Workbench'-Auto Score 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B21:J38", 2, "#,##0.00"); // 'Workbench'-Manager Alpha 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B41:K58", 2, "#,##0.00"); // 'Workbench'-Metrics Over Selected Time Frames 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "C62:D86", 2, "#,##0.00"); // 'Workbench'-Annualized Return 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "C89:D94", 2, "#,##0.00"); // 'Workbench'-Annualized Return 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B97:M114", 2, "#,##0.00"); // 'Workbench'-Net Monthly Return 2 digits

                //Verify content of the downloaded file by doing compare with the baseline excel files.
                sheetName = "Raw Data";
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline2, sheetName, fileNameBaseline2);
                verifyPoints.Add(summaryTC = "Verify the information of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundNameManual + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                //// Verify Excel format of the downloaded file by doing compare with the baseline excel files.
                //sheetName = "WorkBench";
                //verifyPoint = ExcelComparision.IsTheFormatSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline2, sheetName, fileNameBaseline2);
                //verifyPoints.Add(summaryTC = "Verify Excel Format of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundNameManual + ")", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                // Verify the information of Single Public Report Sheet 'WorkBench'
                sheetName = "WorkBench";
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline2, sheetName, fileNameBaseline2);
                verifyPoints.Add(summaryTC = "Verify the information of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundNameManual + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC); */
                #endregion

                //Closing Excel Application Process
                foreach (var application in Process.GetProcessesByName("EXCEL"))
                {
                    application.Kill(); System.Threading.Thread.Sleep(1000);
                }

                // Delete Manifest and WorkBook file
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", downloadedfileName);
            }
            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                //Closing Excel Application Process
                foreach (var application in Process.GetProcessesByName("EXCEL"))
                {
                    application.Kill(); System.Threading.Thread.Sleep(1000);
                }

                // Delete Manifest, WorkBook and ExcelCompareDiff file 
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", downloadedfileName);

                // Warning
                ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong! Please check console log.");
            }
            #endregion
        }


        [Test, Category("Regression Testing")]
        public void TC004_ExcelOnline_SinglePublicReport_Manual_UserInput()
        {
            #region Variables declare
            string urlInstance = LoginPage.url;
            const string sourceIcon = "M";
            const string dateSection = "Date Section";
            const string customRiskBenchmarkModelling = "Custom Risk Benchmark Modelling";
            const string feeModelSection = "Fee Model Section";
            const string liquiditySection = "Liquidity Section";
            const string messageCRBM = "This Benchmark does not cover the time range of the report. Select another one.";
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Single Public Report (Fund Manual - UserInput) - TC004");
            try
            {
                // Delete Manifest, WorkBook and ExcelCompareDiff file 
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", downloadedfileName);

                // Login Excel Online and then add (manifest.xml) FAD Add-in
                LoginAction.Instance.LoginExcelOnlineSiteAndAddInNoGodaddy();

                // Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
                string? fileNameBaseline = null;
                string? fileNameBaseline2 = null;
                string? instance = null;
                if (urlInstance.Contains("sandbox"))
                {
                    instance = "sandbox";
                    fileNameBaseline = @"Book_PreludeManual_Input_Sandbox.xlsx";
                    fileNameBaseline2 = @"Book_GoldenChinaManual_Input_Sandbox.xlsx";
                }
                if (urlInstance.Contains("conceptia"))
                {
                    instance = "conceptia";
                    fileNameBaseline = @"Book_PreludeManual_Input.xlsx";
                    fileNameBaseline2 = @"Book_GoldenChinaManual_Input.xlsx";
                }

                #region (Prelude Capital - Prelude Structured Alternatives Fund) Single Public Report
                // Search a Fund - Source = Albourne
                string managerNameManual = "Prelude Capital";
                string fundNameManual = "Prelude Structured Alternatives Fund";
                FundDashboardAction.Instance.InputNameToSearchFund(10, "prelude", managerNameManual, fundNameManual, sourceIcon);

                // Click to expand User Input (section)
                FADAddInAction.Instance.ClickUserInputSection()
                                       .ClickUserInputSubSection(10, dateSection)
                                       .ClickUserInputSubSection(10, customRiskBenchmarkModelling)
                                       .ClickUserInputSubSection(10, feeModelSection)
                                       .ClickUserInputSubSection(10, liquiditySection);

                #region Input Data for (User Input)
                // User Input
                FADAddInAction.Instance.InputTxtLabelField("Tracking Error", "11");

                // Custom Risk Benchmark Modelling
                // Add the 1st CRBM
                FADAddInAction.Instance.ClickCRBMAddButton(10)
                                       .InputTxtNameCRBMRow(10, 1, "s&p")
                                       .WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("S&P 500 Index"))
                                       .ClickNameCRBMReturnOfResults(10, "S&P 500 Index")
                                       .InputNumberBetaCRBMRow(10, 1, "1.5")
                                       .InputNumberGrossExposureCRBMRow(10, 1, "2.5");

                /*// Add the 2nd CRBM
                //FADAddInAction.Instance.PageDownToScrollDownPage(); // --> Apply when run auton on Remote QA-PC 
                FADAddInAction.Instance.ClickCRBMAddButton(10)
                                       .InputTxtNameCRBMRow(10, 2, "S&P US")
                                       .WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("S&P US High Yield Corporate Bond TR Index"))
                                       .ClickNameCRBMReturnOfResults(10, "S&P US High Yield Corporate Bond TR Index");

                // Verify the error message in red toast is shown
                /// KS-238 The Start Date of the CRBM must be earlier or equal to the start date of the fund (fund return)
                verifyPoint = messageCRBM == FADAddInAction.Instance.ErrorMessageNameCRBMGetTextRow(10, 2);
                verifyPoints.Add(summaryTC = "Verify the error message in red toast is shown: " + messageCRBM + " ", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click delete icon of the 2nd CRBM
                FADAddInAction.Instance.ClickCRBMDeleteButton(10, 2)
                                       .WaitForElementInvisible(10, FADAddInPage.nameCRBMInputTxtRow(2)); */

                // Add the 2nd CRBM (another valid)
                FADAddInAction.Instance.ClickCRBMAddButton(10)
                                       .InputTxtNameCRBMRow(10, 2, "msci")
                                       .WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("MSCI Europe Growth Net TR Index (LCL)"))
                                       .ClickNameCRBMReturnOfResults(10, "MSCI Europe Growth Net TR Index (LCL)")
                                       .InputNumberBetaCRBMRow(10, 2, "3.5")
                                       .InputNumberGrossExposureCRBMRow(10, 2, "4.5");

                // Fee Model Section
                FADAddInAction.Instance.InputTxtLabelField("Management Fee", "5");
                FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, "Management Fee Paid", "Quarterly");
                FADAddInAction.Instance.InputTxtLabelField("Performance Fee", "21");
                FADAddInAction.Instance.ClickToCheckTheCheckbox("High Water Mark");
                FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, "Catch Up", "Yes");
                FADAddInAction.Instance.InputTxtLabelField("Catch Up %-age (if Soft)", "11"); // Only enable when "Catch Up" = Yes
                FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, "Crystallization Every X Years", "10");
                /// Hurdle Status
                FADAddInAction.Instance.ClickToCheckTheCheckbox("Hurdle Status");
                FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, "Hurdle Fixed or Relative", "Relative"); // Only enable when Hurdle Status is checked (checkbox)
                                                                                                                  //InputTxtLabelField("Hurdle Rate (%)", "12"); // Only enable when Hurdle Fixed or Relative = Fixed
                FADAddInAction.Instance.InputTxtLabelField("Fixed Percentage on Top of Index Return (%)", "40");
                FADAddInAction.Instance.ClickAddButtonInXTable(10, "hurdle-section");
                FADAddInAction.Instance.InputTxtNameBenchmarkXTableRow(10, "hurdle-section", "1", "msci");
                //.InputTxtSearchLabelField("Hurdle Benchmark", "msci") (new: KS-834/838) (old: Only enable when Hurdle Fixed or Relative = Relative)
                FADAddInAction.Instance.WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("MSCI Emerging Markets Latin America ex Brazil Net TR Index (USD)"));
                FADAddInAction.Instance.ClickNameCRBMReturnOfResults(10, "MSCI Emerging Markets Latin America ex Brazil Net TR Index (USD)");
                FADAddInAction.Instance.InputTxtExposureBenchmarkXTableRow(10, "hurdle-section", "1", "55.7");
                FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, "Hurdle Type", "Soft Hurdle");
                FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, "Ramp Type", "Performance Dependent");
                /// Liquidity Section 
                FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, "Lockup", "None");
                FADAddInAction.Instance.InputTxtLabelField("Lockup Length (months)", "12");
                FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, "Liquidity Frequency", "Monthly");
                FADAddInAction.Instance.InputTxtLabelField("Investor Gate", "90.55");
                FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, "Sidepocket Probability", "50%");
                FADAddInAction.Instance.InputTxtLabelField("Max % of Sidepocket Permitted", "50"); Thread.Sleep(500);

                // wait for the red error message to disappear
                FADAddInAction.Instance.WaitForElementInvisible(10, FADAddInPage.errorInvalidMessage);
                #endregion

                // Click on Run button
                FADAddInAction.Instance.ClickRunButton(10); System.Windows.Forms.SendKeys.SendWait(@"{ENTER}");
                FADAddInAction.Instance.ClickRunButton(10)
                                       .WaitForLoadingIconToDisappear(60, FundDashboardPage.loadingIcon)
                                       .WaitforExcelWebRenderGraphDataAtLastSheet(60); Thread.Sleep(5000); //.WaitforExcelWebRenderDone2(60);

                // Check if the status "Saving.../Saved" is displayed then refresh page, relogin re-enter data
                int timeout = 0;
                while (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == true && timeout < 6)
                {
                    if (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == true && timeout == 4)
                    {
                        Driver.Browser.Navigate().Refresh();
                        LoginAction.Instance.LoginExcelOnlineSiteAfterRefresh();

                        // re enter data to run report
                        // Search a Fund - Source = Albourne
                        managerNameManual = "Prelude Capital";
                        fundNameManual = "Prelude Structured Alternatives Fund";
                        FundDashboardAction.Instance.InputNameToSearchFund(10, "prelude", managerNameManual, fundNameManual, sourceIcon);

                        // Click to expand User Input (section)
                        FADAddInAction.Instance.ClickUserInputSection();
                        FADAddInAction.Instance.ClickUserInputSubSection(10, dateSection);
                        FADAddInAction.Instance.ClickUserInputSubSection(10, customRiskBenchmarkModelling);
                        FADAddInAction.Instance.CheckIfExistingCRBMThenDeleteAll(10, sourceIcon);
                        FADAddInAction.Instance.ClickUserInputSubSection(10, feeModelSection);
                        FADAddInAction.Instance.ClickUserInputSubSection(10, liquiditySection);

                        #region Input Data for (User Input)
                        // User Input
                        FADAddInAction.Instance.InputTxtLabelField("Tracking Error", "11");

                        // Custom Risk Benchmark Modelling
                        // Add the 1st CRBM
                        FADAddInAction.Instance.ClickCRBMAddButton(10)
                                               .InputTxtNameCRBMRow(10, 1, "s&p")
                                               .WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("S&P 500 Index"))
                                               .ClickNameCRBMReturnOfResults(10, "S&P 500 Index")
                                               .InputNumberBetaCRBMRow(10, 1, "1.5")
                                               .InputNumberGrossExposureCRBMRow(10, 1, "2.5");

                        // Add the 2nd CRBM (another valid)
                        FADAddInAction.Instance.ClickCRBMAddButton(10)
                                               .InputTxtNameCRBMRow(10, 2, "msci")
                                               .WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("MSCI Europe Growth Net TR Index (LCL)"))
                                               .ClickNameCRBMReturnOfResults(10, "MSCI Europe Growth Net TR Index (LCL)")
                                               .InputNumberBetaCRBMRow(10, 2, "3.5")
                                               .InputNumberGrossExposureCRBMRow(10, 2, "4.5");

                        // Fee Model Section
                        FADAddInAction.Instance.InputTxtLabelField("Management Fee", "5");
                        FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, "Management Fee Paid", "Quarterly");
                        FADAddInAction.Instance.InputTxtLabelField("Performance Fee", "21");
                        FADAddInAction.Instance.ClickToCheckTheCheckbox("High Water Mark");
                        FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, "Catch Up", "Yes");
                        FADAddInAction.Instance.InputTxtLabelField("Catch Up %-age (if Soft)", "11"); // Only enable when "Catch Up" = Yes
                        FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, "Crystallization Every X Years", "10");
                        /// Hurdle Status
                        FADAddInAction.Instance.ClickToCheckTheCheckbox("Hurdle Status");
                        FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, "Hurdle Fixed or Relative", "Relative"); // Only enable when Hurdle Status is checked (checkbox)
                                                                                                                          //InputTxtLabelField("Hurdle Rate (%)", "12"); // Only enable when Hurdle Fixed or Relative = Fixed
                        FADAddInAction.Instance.InputTxtLabelField("Fixed Percentage on Top of Index Return (%)", "40");
                        FADAddInAction.Instance.ClickAddButtonInXTable(10, "hurdle-section");
                        FADAddInAction.Instance.InputTxtNameBenchmarkXTableRow(10, "hurdle-section", "1", "msci");
                        //.InputTxtSearchLabelField("Hurdle Benchmark", "msci") (new: KS-834/838) (old: Only enable when Hurdle Fixed or Relative = Relative)
                        FADAddInAction.Instance.WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("MSCI Emerging Markets Latin America ex Brazil Net TR Index (USD)"));
                        FADAddInAction.Instance.ClickNameCRBMReturnOfResults(10, "MSCI Emerging Markets Latin America ex Brazil Net TR Index (USD)");
                        FADAddInAction.Instance.InputTxtExposureBenchmarkXTableRow(10, "hurdle-section", "1", "55.7");
                        FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, "Hurdle Type", "Soft Hurdle");
                        FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, "Ramp Type", "Performance Dependent");
                        /// Liquidity Section 
                        FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, "Lockup", "None");
                        FADAddInAction.Instance.InputTxtLabelField("Lockup Length (months)", "12");
                        FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, "Liquidity Frequency", "Monthly");
                        FADAddInAction.Instance.InputTxtLabelField("Investor Gate", "90.55");
                        FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, "Sidepocket Probability", "50%");
                        FADAddInAction.Instance.InputTxtLabelField("Max % of Sidepocket Permitted", "50"); Thread.Sleep(500);

                        // wait for the red error message to disappear
                        FADAddInAction.Instance.WaitForElementInvisible(10, FADAddInPage.errorInvalidMessage);
                        #endregion

                        // Click on Run button
                        FADAddInAction.Instance.ClickRunButton(10)
                                               .WaitForLoadingIconToDisappear(60, FundDashboardPage.loadingIcon)
                                               .WaitforExcelWebRenderDone2(60);
                        break;
                    }
                    if (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == false)
                    {
                        Thread.Sleep(11000);

                        // Check if the dialog "Invalid sheet" is displayed then click on OK button
                        if (FADAddInAction.Instance.IsElementPresent(NavigationPage.okInvalidSheetButtonExcelOnline))
                        {
                            NavigationAction.Instance.ClicKOkButtonInvalidSheetExcelOnline(10); Thread.Sleep(1000);
                        }
                        Thread.Sleep(1000);
                        break;
                    }
                    if (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == true && timeout == 5)
                    {
                        Console.WriteLine("timeout - Excel WebRender fail to load done!!!");
                        break;
                    }
                    Thread.Sleep(1000);
                    timeout++;
                }

                // Download Excel Online Report
                NavigationAction.Instance.DownloadExcelOnlineReport(10);

                // Verify Excel file was downloaded successfully
                verifyPoint = LoginAction.Instance.CheckFileDownloadIsComplete(9, ExcelFileDownloadedPath, downloadedfileName);
                verifyPoints.Add(summaryTC = "Verify '" + downloadedfileName + "' file is downloaded successful", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Paste value for the downloaded file
                System.Threading.Thread.Sleep(2000);
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "Raw Data", "B2:H500", 6, "#,######0.000000"); // 'Raw Data' 6 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B18:C20", 2, "#,##0.00"); // 'Workbench'-Auto Score 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B23:J34", 2, "#,##0.00"); // 'Workbench'-Manager Alpha 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B37:K54", 2, "#,##0.00"); // 'Workbench'-Metrics Over Selected Time Frames 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "C58:F82", 2, "#,##0.00"); // 'Workbench'-Annualized Return 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "C85:F90", 2, "#,##0.00"); // 'Workbench'-Annualized Return 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B93:M104", 2, "#,##0.00"); // 'Workbench'-Net Monthly Return 2 digits

                // Verify content of the downloaded file by doing compare with the baseline excel files
                sheetName = "Raw Data";
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline, sheetName, fileNameBaseline);
                verifyPoints.Add(summaryTC = "Verify the information of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundNameManual + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                //// Verify Excel format of the downloaded file by doing compare with the baseline excel files.
                //sheetName = "WorkBench";
                //verifyPoint = ExcelComparision.IsTheFormatSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline, sheetName, fileNameBaseline);
                //verifyPoints.Add(summaryTC = "Verify Excel Format of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundNameManual + ")", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                // Verify the information of Single Public Report Sheet 'WorkBench'
                sheetName = "WorkBench";
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline, sheetName, fileNameBaseline);
                verifyPoints.Add(summaryTC = "Verify the information of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundNameManual + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Delete WorkBook file
                LoginAction.Instance.DeleteAllFilesWorkBookExcel();
                #endregion

                #region (Greenwoods - Golden China Fund) Single Public Report
                /*
                managerNameManual = "Greenwoods";
                fundNameManual = "Golden China Fund";

                // Switch to the FAD Add-in frame (to interact with Workbench for Excel Online)
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameIdFADAddInCurrentInstance(instance));

                // Search a Fund - Source = Albourne
                FundDashboardAction.Instance.InputNameToSearchFund(10, "greenwoods", managerNameManual, fundNameManual, sourceIcon);

                // Click to expand User Input (section)
                FADAddInAction.Instance.ClickUserInputSubSection(10, dateSection)
                                       .ClickUserInputSubSection(10, customRiskBenchmarkModelling)
                                       .ClickUserInputSubSection(10, feeModelSection);

                // Check if A red notification is shown then waiting for that is disappeared
                System.Threading.Thread.Sleep(1000);
                if (FADAddInAction.Instance.IsElementPresent(FADAddInPage.MessageFundDoesNotHaveReturnData))
                {
                    FADAddInAction.Instance.WaitForLoadingIconToDisappear(10, FADAddInPage.MessageFundDoesNotHaveReturnData);
                }

                #region Input Data for (User Input)
                // User Input
                FADAddInAction.Instance.InputTxtLabelField("Tracking Error", "9");

                //if (urlInstance.Contains("conceptia")) 
                //{
                //    // Custom Risk Benchmark Modelling
                //    //  Delete 2 CRBM previous
                //    FADAddInAction.Instance.ClickCRBMDeleteButton(10, "1").ClickCRBMDeleteButton(10, "1");
                //    System.Threading.Thread.Sleep(500);
                //}

                // Custom Risk Benchmark Modelling
                // Delete 2 CRBM previous
                //FADAddInAction.Instance.ClickCRBMDeleteButton(10, "1").ClickCRBMDeleteButton(10, "1");

                // Add the 1st CRBM
                FADAddInAction.Instance.ClickCRBMAddButton(10)
                                       .InputTxtNameCRBMRow(10, 1, "s&p")
                                       .WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("S&P 500 Industrials Net TR Index")) // old name: S&P 500 Industrials (Sector) Net Total Return Index
                                       .ClickNameCRBMReturnOfResults(10, "S&P 500 Industrials Net TR Index")
                                       .InputNumberBetaCRBMRow(10, 1, "-0.5")
                                       .InputNumberGrossExposureCRBMRow(10, 1, "-1.5");
                // Fee Model Section
                FADAddInAction.Instance.InputTxtLabelField("Management Fee", "3.25")
                                       .ClickAndSelectItemInDropdown(10, "Management Fee Paid", "Monthly")
                                       .InputTxtLabelField("Performance Fee", "14")
                                       .ClickToCheckTheCheckbox("High Water Mark")
                                       .ClickAndSelectItemInDropdown(10, "Catch Up", "No")
                                       //.InputTxtLabelField("Catch Up %-age (if Soft)", "11") // Only enable when "Catch Up" = Yes
                                       .ClickAndSelectItemInDropdown(10, "Crystallization Every X Years", "9")
                                   /// Hurdle Status
                                   .ClickToCheckTheCheckbox("Hurdle Status")
                                       .ClickAndSelectItemInDropdown(10, "Hurdle Fixed or Relative", "Fixed") // Only enable when Hurdle Status is checked (checkbox)
                                       .InputTxtLabelField("Hurdle Rate (%)", "11") // Only enable when Hurdle Fixed or Relative = Fixed
                                                                                    //.InputTxtSearchLabelField("Hurdle Benchmark", "msci") // Only enable when Hurdle Fixed or Relative = Relative
                                                                                    //.WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("MSCI EM LATIN AMERICA ex BRAZIL Net Return USD Index"))
                                                                                    //.ClickNameCRBMReturnOfResults(10, "MSCI EM LATIN AMERICA ex BRAZIL Net Return USD Index")
                .ClickAndSelectItemInDropdown(10, "Hurdle Type", "Hard Hurdle")
                .ClickAndSelectItemInDropdown(10, "Ramp Type", "NAV Dependent");
                #endregion

                // Click on Run button
                FADAddInAction.Instance.ClickRunButton(10)
                                       .WaitForLoadingIconToDisappear(60, FundDashboardPage.loadingIcon)
                                       .WaitforExcelWebRenderDone(60);

                // Download Excel Online Report
                NavigationAction.Instance.DownloadExcelOnlineReport(10);

                // Verify Excel file was downloaded successfully
                verifyPoint = LoginAction.Instance.CheckFileDownloadIsComplete(9, ExcelFileDownloadedPath, downloadedfileName);
                verifyPoints.Add(summaryTC = "The '" + downloadedfileName + "' (" + fundNameManual + ") file is downloaded successful", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Paste value for the downloaded file
                System.Threading.Thread.Sleep(2000);
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "Raw Data", "B2:F500", 6, "#,######0.000000"); // 'Raw Data' 6 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B17:C19", 2, "#,##0.00"); // 'Workbench'-Auto Score 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B22:I39", 2, "#,##0.00"); // 'Workbench'-Manager Alpha 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B42:J56", 2, "#,##0.00"); // 'Workbench'-Metrics Over Selected Time Frames 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "C59:E83", 2, "#,##0.00"); // 'Workbench'-Annualized Return 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "C86:E91", 2, "#,##0.00"); // 'Workbench'-Annualized Return 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B94:M111", 2, "#,##0.00"); // 'Workbench'-Net Monthly Return 2 digits

                // Verify content of the downloaded file by doing compare with the baseline excel files
                sheetName = "Raw Data";
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline2, sheetName, fileNameBaseline2);
                verifyPoints.Add(summaryTC = "Verify the information of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundNameManual + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                //// Verify Excel format of the downloaded file by doing compare with the baseline excel files.
                //sheetName = "WorkBench";
                //verifyPoint = ExcelComparision.IsTheFormatSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline2, sheetName, fileNameBaseline2);
                //verifyPoints.Add(summaryTC = "Verify Excel Format of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundNameManual + ")", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                // Verify the information of Single Public Report Sheet 'WorkBench'
                sheetName = "WorkBench";
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline2, sheetName, fileNameBaseline2);
                verifyPoints.Add(summaryTC = "Verify the information of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundNameManual + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                */
                #endregion

                //Closing Excel Application Process
                foreach (var application in Process.GetProcessesByName("EXCEL"))
                {
                    application.Kill(); System.Threading.Thread.Sleep(1000);
                }

                // Delete Manifest and WorkBook file
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", downloadedfileName);
            }
            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                //Closing Excel Application Process
                foreach (var application in Process.GetProcessesByName("EXCEL"))
                {
                    application.Kill(); System.Threading.Thread.Sleep(1000);
                }

                // Delete Manifest, WorkBook and ExcelCompareDiff file 
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", downloadedfileName);

                // Warning
                ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong! Please check console log.");
            }
            #endregion
        }


        [Test, Category("Regression Testing")]
        public void TC005_ExcelOnline_SinglePublicReport_Solovis()
        {
            #region Variables declare
            string urlInstance = LoginPage.url;
            const string sourceIcon = "S";
            const string dateSection = "Date Section";
            const string customRiskBenchmarkModelling = "Custom Risk Benchmark Modelling";
            const string feeModelSection = "Fee Model Section";
            const string liquiditySection = "Liquidity Section";
            string msg;
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Single Public Report (Solovis) - TC005");
            try
            {
                // Delete Manifest, WorkBook and ExcelCompareDiff file 
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", downloadedfileName);

                // Login Excel Online and then add (manifest.xml) FAD Add-in
                LoginAction.Instance.LoginExcelOnlineSiteAndAddInNoGodaddy();

                // Search a Fund - Source = Solovis
                string managerName = "Laurion Capital Management LP";
                string fundName = "Laurion Capital Ltd.";
                FundDashboardAction.Instance.InputNameToSearchFund(10, "laurion")
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResultsWithItemSource(managerName, sourceIcon))
                                            .ClickFundNameReturnOfResults(10, managerName, sourceIcon)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResultsWithItemSource(fundName, sourceIcon))
                                            .ClickFundNameReturnOfResults(10, fundName, sourceIcon);

                // Check if loading icon shows then wait for it to disappear
                if (FundDashboardAction.Instance.IsElementPresent(FundDashboardPage.loadingIcon))
                {
                    FundDashboardAction.Instance.WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);
                }

                // Click to expand User Input (section)
                FADAddInAction.Instance.ClickUserInputSection()
                                       .ClickUserInputSubSection(10, dateSection)
                                       .ClickUserInputSubSection(10, customRiskBenchmarkModelling)
                                       .ClickUserInputSubSection(10, feeModelSection)
                                       .ClickUserInputSubSection(10, liquiditySection); Thread.Sleep(1000);

                // Check if A red notification is shown then waiting for that is disappeared
                msg = "This fund does not have Return data. Select another data source or upload the fund data.";
                if (FADAddInAction.Instance.IsElementPresent(FADAddInPage.errorInvalidMessageContent(msg)))
                {
                    FADAddInAction.Instance.WaitForLoadingIconToDisappear(10, FADAddInPage.errorInvalidMessageContent(msg));
                }

                // Add an invalid crbm (the 3rd crbm)
                FADAddInAction.Instance.PageDownToScrollDownPage();
                FADAddInAction.Instance.WaitForElementVisible(10, FADAddInPage.deleteCRBMbutton("2"));
                System.Threading.Thread.Sleep(2000);
                FADAddInAction.Instance.ClickCRBMAddButton(10);
                FADAddInAction.Instance.PageDownToScrollDownPage();
                FADAddInAction.Instance.WaitForElementVisible(10, FADAddInPage.deleteCRBMbutton("3"));
                System.Threading.Thread.Sleep(1000);
                FADAddInAction.Instance.InputTxtNameCRBMRow(10, 3, "abcdef")
                                       .WaitForElementInvisible(20, FADAddInPage.nameCRBMLoadingSpinnerRow(3))
                                       .ClickUserInputSubSection(10, customRiskBenchmarkModelling)  // Collapse
                                       .ClickUserInputSubSection(10, customRiskBenchmarkModelling)  // Expand
                                       .WaitForElementVisible(10, FADAddInPage.deleteCRBMbutton("3"));

                // Verify red error message is shown
                msg = "We couldn't find anything matching your search.";
                verifyPoint = msg == FADAddInAction.Instance.RedErrorMessageNameCRBMGetTextRow(10, 3);
                verifyPoints.Add(summaryTC = "Verify the red error message is shown: " + msg + " ", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                System.Threading.Thread.Sleep(3000);
                
                // Check if the 3rd crbm is existing then click Delete button
                if (FADAddInAction.Instance.IsElementPresent(FADAddInPage.deleteCRBMbutton("3")))
                {
                    FADAddInAction.Instance.ClickCRBMDeleteButton(10, "3"); // Delete the 3rd crbm }
                    System.Threading.Thread.Sleep(500);
                }

                // Add a crbm that does not cover the time range of the report
                FADAddInAction.Instance.WaitForElementInvisible(10, FADAddInPage.nameCRBMInputTxtRow(3))
                                       .ClickCRBMAddButton(10) // Click add (+) CRBM button (to add the 3rd crbm)
                                       .WaitForElementVisible(10, FADAddInPage.deleteCRBMbutton("3"));

                System.Threading.Thread.Sleep(1000); // Workaround for issue can't search crbm 
                /*FADAddInAction.Instance.ClickCRBMDeleteButton(10, 3)
                                       .ClickCRBMAddButton(10)
                                       .WaitForElementVisible(10, FADAddInPage.deleteCRBMbutton(3))
                                       .InputTxtNameCRBMRow(10, 3, "S&P US")
                                       .WaitForElementInvisible(20, FADAddInPage.nameCRBMLoadingSpinnerRow(3));
                FADAddInAction.Instance.WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("S&P US High Yield Corporate Bond TR Index"))
                                       .ClickNameCRBMReturnOfResults(10, "S&P US High Yield Corporate Bond TR Index");

                // Verify the error message in red toast is shown
                /// KS-238 The Start Date of the CRBM must be earlier or equal to the start date of the fund (fund return)
                msg = "This Benchmark does not cover the time range of the report. Select another one.";
                verifyPoint = msg == FADAddInAction.Instance.ErrorMessageNameCRBMGetTextRow(10, 3);
                verifyPoints.Add(summaryTC = "Verify the error message in red toast is shown: " + msg + " ", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC); */

                // Add a crbm (the 3rd crbm valid)
                FADAddInAction.Instance.ClickCRBMDeleteButton(10, "3") // Delete the 3rd crbm
                                       .WaitForElementInvisible(10, FADAddInPage.nameCRBMInputTxtRow(3))
                                       .ClickCRBMAddButton(10) // Click add (+) CRBM button (to add the 3rd crbm)
                                       .InputTxtNameCRBMRow(10, 3, "msci")
                                       .WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("MSCI Europe Growth Net TR Index (LCL)"))
                                       .ClickNameCRBMReturnOfResults(10, "MSCI Europe Growth Net TR Index (LCL)")
                                       .InputNumberBetaCRBMRow(10, 3, "-1.5")
                                       .InputNumberGrossExposureCRBMRow(10, 3, "80")
                                       /// Liquidity Section 
                                       .ClickAndSelectItemInDropdown(10, "Lockup", "Soft")
                                       .InputTxtLabelField("Lockup Length (months)", "12")
                                       .ClickAndSelectItemInDropdown(10, "Liquidity Frequency", "Quarterly")
                                       //.InputTxtLabelField("Investor Gate", "")
                                       .ClickAndSelectItemInDropdown(10, "Sidepocket Probability", "100%")
                                       .InputTxtLabelField("Max % of Sidepocket Permitted", "5"); Thread.Sleep(500);

                // wait for the red error message to disappear
                FADAddInAction.Instance.WaitForElementInvisible(10, FADAddInPage.errorInvalidMessage);

                // Click on Run button
                FADAddInAction.Instance.ClickRunButton(10); System.Windows.Forms.SendKeys.SendWait(@"{ENTER}");
                FADAddInAction.Instance.ClickRunButton(10)
                                       .WaitForLoadingIconToDisappear(60, FundDashboardPage.loadingIcon)
                                       .WaitforExcelWebRenderGraphDataAtLastSheet(60); Thread.Sleep(5000); //.WaitforExcelWebRenderDone2(60);

                // Check if the status "Saving.../Saved" is displayed then refresh page, relogin re-enter data
                int timeout = 0;
                while (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == true && timeout < 6)
                {
                    if (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == true && timeout == 4)
                    {
                        Driver.Browser.Navigate().Refresh();
                        LoginAction.Instance.LoginExcelOnlineSiteAfterRefresh();

                        // re enter data to run report
                        // Search a Fund - Source = Albourne
                        managerName = "Laurion Capital Management LP";
                        fundName = "Laurion Capital Ltd.";
                        FundDashboardAction.Instance.InputNameToSearchFund(10, "laurion")
                                                    .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResultsWithItemSource(managerName, sourceIcon))
                                                    .ClickFundNameReturnOfResults(10, managerName, sourceIcon)
                                                    .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResultsWithItemSource(fundName, sourceIcon))
                                                    .ClickFundNameReturnOfResults(10, fundName, sourceIcon);

                        // Check if loading icon shows then wait for it to disappear
                        if (FundDashboardAction.Instance.IsElementPresent(FundDashboardPage.loadingIcon))
                        {
                            FundDashboardAction.Instance.WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);
                        }

                        // Click to expand User Input (section)
                        FADAddInAction.Instance.ClickUserInputSection()
                                               .ClickUserInputSubSection(10, dateSection)
                                               .ClickUserInputSubSection(10, customRiskBenchmarkModelling)
                                               .ClickUserInputSubSection(10, feeModelSection)
                                               .ClickUserInputSubSection(10, liquiditySection); Thread.Sleep(1000);

                        #region Input Data for (User Input)
                        // Check if A red notification is shown then waiting for that is disappeared
                        msg = "This fund does not have Return data. Select another data source or upload the fund data.";
                        if (FADAddInAction.Instance.IsElementPresent(FADAddInPage.errorInvalidMessageContent(msg)))
                        {
                            FADAddInAction.Instance.WaitForElementInvisible(10, FADAddInPage.errorInvalidMessageContent(msg));
                        }

                        // Add an invalid crbm (the 3rd crbm)
                        FADAddInAction.Instance.PageDownToScrollDownPage();
                        FADAddInAction.Instance.WaitForElementVisible(10, FADAddInPage.deleteCRBMbutton("2"));
                        System.Threading.Thread.Sleep(2000);
                        FADAddInAction.Instance.ClickCRBMAddButton(10);
                        FADAddInAction.Instance.PageDownToScrollDownPage();
                        FADAddInAction.Instance.WaitForElementVisible(10, FADAddInPage.deleteCRBMbutton("3"));
                        System.Threading.Thread.Sleep(1000);
                        FADAddInAction.Instance.InputTxtNameCRBMRow(10, 3, "abcdef")
                                               .WaitForElementInvisible(20, FADAddInPage.nameCRBMLoadingSpinnerRow(3))
                                               .ClickUserInputSubSection(10, customRiskBenchmarkModelling)  // Collapse
                                               .ClickUserInputSubSection(10, customRiskBenchmarkModelling)  // Expand
                                               .WaitForElementVisible(10, FADAddInPage.deleteCRBMbutton("3"));

                        // Verify red error message is shown
                        //msg = "We couldn't find anything matching your search.";
                        //verifyPoint = msg == FADAddInAction.Instance.RedErrorMessageNameCRBMGetTextRow(10, 3);
                        //verifyPoints.Add(summaryTC = "Verify the red error message is shown: " + msg + " ", verifyPoint);
                        //ExtReportResult(verifyPoint, summaryTC);
                        System.Threading.Thread.Sleep(3000);

                        // Check if the 3rd crbm is existing then click Delete button
                        if (FADAddInAction.Instance.IsElementPresent(FADAddInPage.deleteCRBMbutton("3")))
                        {
                            FADAddInAction.Instance.ClickCRBMDeleteButton(10, "3"); // Delete the 3rd crbm }
                            System.Threading.Thread.Sleep(500);
                        }

                        // Add a crbm that does not cover the time range of the report
                        FADAddInAction.Instance.WaitForElementInvisible(10, FADAddInPage.nameCRBMInputTxtRow(3))
                                               .ClickCRBMAddButton(10) // Click add (+) CRBM button (to add the 3rd crbm)
                                               .WaitForElementVisible(10, FADAddInPage.deleteCRBMbutton("3"));

                        System.Threading.Thread.Sleep(1000); // Workaround for issue can't search crbm 

                        // Add a crbm (the 3rd crbm valid)
                        FADAddInAction.Instance.ClickCRBMDeleteButton(10, "3") // Delete the 3rd crbm
                                               .WaitForElementInvisible(10, FADAddInPage.nameCRBMInputTxtRow(3))
                                               .ClickCRBMAddButton(10) // Click add (+) CRBM button (to add the 3rd crbm)
                                               .InputTxtNameCRBMRow(10, 3, "msci")
                                               .WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("MSCI Europe Growth Net TR Index (LCL)"))
                                               .ClickNameCRBMReturnOfResults(10, "MSCI Europe Growth Net TR Index (LCL)")
                                               .InputNumberBetaCRBMRow(10, 3, "-1.5")
                                               .InputNumberGrossExposureCRBMRow(10, 3, "80")
                                               /// Liquidity Section 
                                               .ClickAndSelectItemInDropdown(10, "Lockup", "Soft")
                                               .InputTxtLabelField("Lockup Length (months)", "12")
                                               .ClickAndSelectItemInDropdown(10, "Liquidity Frequency", "Quarterly")
                                               //.InputTxtLabelField("Investor Gate", "")
                                               .ClickAndSelectItemInDropdown(10, "Sidepocket Probability", "100%")
                                               .InputTxtLabelField("Max % of Sidepocket Permitted", "5"); Thread.Sleep(500);

                        // wait for the red error message to disappear
                        FADAddInAction.Instance.WaitForElementInvisible(10, FADAddInPage.errorInvalidMessage);
                        #endregion

                        // Click on Run button
                        FADAddInAction.Instance.ClickRunButton(10)
                                               .WaitForLoadingIconToDisappear(60, FundDashboardPage.loadingIcon)
                                               .WaitforExcelWebRenderDone2(60);
                        break;
                    }
                    if (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == false)
                    {
                        Thread.Sleep(11000);

                        // Check if the dialog "Invalid sheet" is displayed then click on OK button
                        if (FADAddInAction.Instance.IsElementPresent(NavigationPage.okInvalidSheetButtonExcelOnline))
                        {
                            NavigationAction.Instance.ClicKOkButtonInvalidSheetExcelOnline(10); Thread.Sleep(1000);
                        }
                        Thread.Sleep(1000);
                        break;
                    }
                    if (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == true && timeout == 5)
                    {
                        Console.WriteLine("timeout - Excel WebRender fail to load done!!!");
                        break;
                    }
                    Thread.Sleep(1000);
                    timeout++;
                }

                // Download Excel Online Report
                NavigationAction.Instance.DownloadExcelOnlineReport(10);

                // Verify Excel file was downloaded successfully
                verifyPoint = LoginAction.Instance.CheckFileDownloadIsComplete(9, ExcelFileDownloadedPath, downloadedfileName);
                verifyPoints.Add(summaryTC = "Verify the '" + downloadedfileName + "' file is downloaded successful", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Paste value for the downloaded file
                System.Threading.Thread.Sleep(2000);
                sheetName = "Raw Data";
                string columnName = "MSCI ACWI IMI with USA Net TR Index (USD)";
                ExcelComparision.ReOrderColumnsWithConditions(ExcelFileDownloadedPath + @"\" + downloadedfileName, sheetName, "E1", columnName, "F:F", "E:E");
                sheetName = "WorkBench";
                ExcelComparision.SortColumnRange(ExcelFileDownloadedPath + @"\" + downloadedfileName, sheetName, "A12:C14"); // Sort CRBM - Index Name
                ExcelComparision.ReOrderColumnsWithConditions(ExcelFileDownloadedPath + @"\" + downloadedfileName, sheetName, "D64", columnName, "D64:D97", "F64"); // Sort crbm at 'Annualized Return'
                /// Get 2 Digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "Raw Data", "B2:H1000", 2, "#,##0.00"); // if get 'Raw Data' 6 digits --> 6, "#,######0.000000"
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B17:C19", 2, "#,##0.00"); // 'Workbench'-Auto Score 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B22:J42", 2, "#,##0.00"); // 'Workbench'-Manager Alpha 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B45:K62", 2, "#,##0.00"); // 'Workbench'-Metrics Over Selected Time Frames 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "C65:G89", 2, "#,##0.00"); // 'Workbench'-Annualized Return 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "C92:G97", 2, "#,##0.00"); // 'Workbench'-Annualized Return 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B100:M120", 2, "#,##0.00"); // 'Workbench'-Net Monthly Return 2 digits

                //Verify content of the downloaded file by doing compare with the baseline excel files.
                /// Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
                string? fileNameBaseline = null;
                if (urlInstance.Contains("sandbox"))
                {
                    fileNameBaseline = "Book_Laurion.xlsx";
                }
                if (urlInstance.Contains("conceptia"))
                {
                    fileNameBaseline = @"Book_Laurion_Staging.xlsx";
                }
                sheetName = "Raw Data";
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline, sheetName, fileNameBaseline);
                verifyPoints.Add(summaryTC = "Verify the information of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundName + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Verify the information of Single Public Report Sheet 'WorkBench'
                sheetName = "WorkBench";
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline, sheetName, fileNameBaseline);
                verifyPoints.Add(summaryTC = "Verify the information of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundName + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                //Closing Excel Application Process
                foreach (var application in Process.GetProcessesByName("EXCEL"))
                {
                    application.Kill(); System.Threading.Thread.Sleep(1000);
                }

                // Delete Manifest and WorkBook file
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", downloadedfileName);
            }
            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                //Closing Excel Application Process
                foreach (var application in Process.GetProcessesByName("EXCEL"))
                {
                    application.Kill(); System.Threading.Thread.Sleep(1000);
                }

                // Delete Manifest, WorkBook and ExcelCompareDiff file 
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", downloadedfileName);

                // Warning
                ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong! Please check console log.");
            }
            #endregion
        }

        [Test, Category("Regression Testing")]
        public void TC006_ExcelOnline_SinglePublicReport_AEvestment()
        {
            #region Variables declare
            string urlInstance = LoginPage.url;
            const string sourceIcon = "E";
            const string dateSection = "Date Section";
            const string customRiskBenchmarkModelling = "Custom Risk Benchmark Modelling";
            const string feeModelSection = "Fee Model Section";
            const string liquiditySection = "Liquidity Section";
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Single Public Report (Alt Evestment) - TC006");
            try
            {
                // Delete Manifest, WorkBook and ExcelCompareDiff file 
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", downloadedfileName);

                // Login Excel Online and then add (manifest.xml) FAD Add-in
                LoginAction.Instance.LoginExcelOnlineSiteAndAddInNoGodaddy();

                // Search a Fund - Source = Solovis
                string managerName = "Amber Capital";
                string fundName = "Amber European Long Opportunities Fund";
                FundDashboardAction.Instance.InputNameToSearchFund(10, "amber", managerName, fundName, sourceIcon);

                // Click to expand User Input (section)
                FADAddInAction.Instance.ClickUserInputSection()
                                       .ClickUserInputSubSection(10, dateSection)
                                       .ClickUserInputSubSection(10, customRiskBenchmarkModelling)
                                       .ClickUserInputSubSection(10, feeModelSection)
                                       .ClickUserInputSubSection(10, liquiditySection);

                // Check if A red notification is shown then waiting for that is disappeared
                if (FADAddInAction.Instance.IsElementPresent(FADAddInPage.errorInvalidMessage))
                {
                    FADAddInAction.Instance.WaitForLoadingIconToDisappear(10, FADAddInPage.errorInvalidMessage);
                }

                /// Liquidity Section 
                FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, "Lockup", "Soft")
                                       .InputTxtLabelField("Lockup Length (months)", "12")
                                       .ClickAndSelectItemInDropdown(10, "Liquidity Frequency", "Monthly")
                                       .InputTxtLabelField("Investor Gate", "71")
                                       .ClickAndSelectItemInDropdown(10, "Sidepocket Probability", "50%")
                                       .InputTxtLabelField("Max % of Sidepocket Permitted", "91"); Thread.Sleep(500);
                FADAddInAction.Instance.WaitForElementInvisible(10, FADAddInPage.errorInvalidMessageContent("All the liquidity parameters need to be populated except for Investor Gate"));

                // Click on Run button
                FADAddInAction.Instance.ClickRunButton(10)
                                       .WaitForLoadingIconToDisappear(60, FundDashboardPage.loadingIcon)
                                       .WaitforExcelWebRenderGraphDataAtLastSheet(60); Thread.Sleep(5000); //.WaitforExcelWebRenderDone2(60);

                // Check if the status "Saving.../Saved" is displayed then refresh page, relogin re-enter data
                int timeout = 0;
                while (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == true && timeout < 6)
                {
                    if (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == true && timeout == 4)
                    {
                        Driver.Browser.Navigate().Refresh();
                        LoginAction.Instance.LoginExcelOnlineSiteAfterRefresh();

                        // re enter data to run report
                        // Search a Fund - Source = Solovis
                        managerName = "Amber Capital";
                        fundName = "Amber European Long Opportunities Fund";
                        FundDashboardAction.Instance.InputNameToSearchFund(10, "amber", managerName, fundName, sourceIcon);

                        // Click to expand User Input (section)
                        FADAddInAction.Instance.ClickUserInputSection()
                                               .ClickUserInputSubSection(10, dateSection)
                                               .ClickUserInputSubSection(10, customRiskBenchmarkModelling)
                                               .ClickUserInputSubSection(10, feeModelSection)
                                               .ClickUserInputSubSection(10, liquiditySection);

                        // Check if A red notification is shown then waiting for that is disappeared
                        if (FADAddInAction.Instance.IsElementPresent(FADAddInPage.errorInvalidMessage))
                        {
                            FADAddInAction.Instance.WaitForLoadingIconToDisappear(10, FADAddInPage.errorInvalidMessage);
                        }

                        /// Liquidity Section 
                        FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, "Lockup", "Soft")
                                               .InputTxtLabelField("Lockup Length (months)", "12")
                                               .ClickAndSelectItemInDropdown(10, "Liquidity Frequency", "Monthly")
                                               .InputTxtLabelField("Investor Gate", "71")
                                               .ClickAndSelectItemInDropdown(10, "Sidepocket Probability", "50%")
                                               .InputTxtLabelField("Max % of Sidepocket Permitted", "91"); Thread.Sleep(500);
                        FADAddInAction.Instance.WaitForElementInvisible(10, FADAddInPage.errorInvalidMessageContent("All the liquidity parameters need to be populated except for Investor Gate"));

                        // Click on Run button
                        FADAddInAction.Instance.ClickRunButton(10)
                                               .WaitForLoadingIconToDisappear(60, FundDashboardPage.loadingIcon)
                                               .WaitforExcelWebRenderDone2(60);
                        break;
                    }
                    if (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == false)
                    {
                        Thread.Sleep(11000);

                        // Check if the dialog "Invalid sheet" is displayed then click on OK button
                        if (FADAddInAction.Instance.IsElementPresent(NavigationPage.okInvalidSheetButtonExcelOnline))
                        {
                            NavigationAction.Instance.ClicKOkButtonInvalidSheetExcelOnline(10); Thread.Sleep(1000);
                        }
                        Thread.Sleep(1000);
                        break;
                    }
                    if (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == true && timeout == 5)
                    {
                        Console.WriteLine("timeout - Excel WebRender fail to load done!!!");
                        break;
                    }
                    Thread.Sleep(1000);
                    timeout++;
                }

                // Download Excel Online Report
                NavigationAction.Instance.DownloadExcelOnlineReport(10);

                // Verify Excel file was downloaded successfully
                verifyPoint = LoginAction.Instance.CheckFileDownloadIsComplete(9, ExcelFileDownloadedPath, downloadedfileName);
                verifyPoints.Add(summaryTC = "Verify the '" + downloadedfileName + "' file is downloaded successful", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                //Verify content of the downloaded file by doing compare with the baseline excel files.
                /// Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
                string? fileNameBaseline = null;
                if (urlInstance.Contains("sandbox"))
                {
                    fileNameBaseline = "Book_Amber_AEvestment_Sandbox.xlsx";
                }
                if (urlInstance.Contains("conceptia"))
                {
                    fileNameBaseline = @"Book_Amber_AEvestment_Staging.xlsx";
                }

                // Paste value for Downloaded File
                System.Threading.Thread.Sleep(2000);
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "Raw Data", "B2:F1000", 2, "#,##0.00"); // if get 'Raw Data' 6 digits-- > 6, "#,######0.000000"
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B15:C17", 2, "#,##0.00"); // 'Workbench'-Auto Score 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B20:J24", 2, "#,##0.00"); // 'Workbench'-Manager Alpha 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B27:K44", 2, "#,##0.00"); // 'Workbench'-Metrics Over Selected Time Frames 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "C48:D68", 2, "#,##0.00"); // 'Workbench'-Annualized Return 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "C71:D76", 2, "#,##0.00"); // 'Workbench'-Annualized Return 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B79:M83", 2, "#,##0.00"); // 'Workbench'-Net Monthly Return 2 digits

                sheetName = "Raw Data";
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline, sheetName, fileNameBaseline);
                verifyPoints.Add(summaryTC = "Verify the information of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundName + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Verify the information of Single Public Report Sheet 'WorkBench'
                sheetName = "WorkBench";
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline, sheetName, fileNameBaseline);
                verifyPoints.Add(summaryTC = "Verify the information of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundName + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                //Closing Excel Application Process
                foreach (var application in Process.GetProcessesByName("EXCEL"))
                {
                    application.Kill(); System.Threading.Thread.Sleep(1000);
                }

                // Delete Manifest and WorkBook file
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", downloadedfileName);
            }
            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                //Closing Excel Application Process
                foreach (var application in Process.GetProcessesByName("EXCEL"))
                {
                    application.Kill(); System.Threading.Thread.Sleep(1000);
                }

                // Delete Manifest, WorkBook and ExcelCompareDiff file 
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", downloadedfileName);

                // Warning
                ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong! Please check console log.");
            }
            #endregion
        }

        [Test, Category("Regression Testing")]
        public void TC007_ExcelOnline_SinglePublicReport_TEvestment()
        {
            #region Variables declare
            string urlInstance = LoginPage.url;
            const string sourceIcon = "E";
            const string dateSection = "Date Section";
            const string customRiskBenchmarkModelling = "Custom Risk Benchmark Modelling";
            const string feeModelSection = "Fee Model Section";
            const string liquiditySection = "Liquidity Section";
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Single Public Report (Traditional Evestment) - TC007");
            try
            {
                // Delete Manifest, WorkBook and ExcelCompareDiff file 
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", downloadedfileName);

                // Login Excel Online and then add (manifest.xml) FAD Add-in
                LoginAction.Instance.LoginExcelOnlineSiteAndAddInNoGodaddy();

                // Search a Fund - Source = Solovis
                string managerName = "John Hancock Investments";
                string fundName = "John Hancock Multimanager Lifetime Portfolios 2035";
                FundDashboardAction.Instance.InputNameToSearchFund(10, fundName, managerName, fundName, sourceIcon);

                // Click to expand User Input (section)
                FADAddInAction.Instance.ClickUserInputSection()
                                       .ClickUserInputSubSection(10, dateSection)
                                       .ClickUserInputSubSection(10, customRiskBenchmarkModelling)
                                       .ClickUserInputSubSection(10, feeModelSection)
                                       .ClickUserInputSubSection(10, liquiditySection);

                // Check if A red notification is shown then waiting for that is disappeared
                if (FADAddInAction.Instance.IsElementPresent(FADAddInPage.errorInvalidMessage))
                {
                    //FADAddInAction.Instance.WaitForLoadingIconToDisappear(10, FADAddInPage.MessageFundDoesNotHaveReturnData);
                    FADAddInAction.Instance.WaitForElementInvisible(10, FADAddInPage.errorInvalidMessage);
                }

                /// Liquidity Section 
                FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, "Lockup", "Hard")
                                       .InputTxtLabelField("Lockup Length (months)", "12")
                                       .ClickAndSelectItemInDropdown(10, "Liquidity Frequency", "Monthly")
                                       .InputTxtLabelField("Investor Gate", "60.55")
                                       .ClickAndSelectItemInDropdown(10, "Sidepocket Probability", "50%")
                                       .InputTxtLabelField("Max % of Sidepocket Permitted", "90.55"); Thread.Sleep(500);
                FADAddInAction.Instance.WaitForElementInvisible(10, FADAddInPage.errorInvalidMessageContent("All the liquidity parameters need to be populated except for Investor Gate"));

                // Click on Run button
                FADAddInAction.Instance.ClickRunButton(10); System.Windows.Forms.SendKeys.SendWait(@"{ENTER}");
                FADAddInAction.Instance.ClickRunButton(10)
                                       .WaitForLoadingIconToDisappear(60, FundDashboardPage.loadingIcon)
                                       .WaitforExcelWebRenderGraphDataAtLastSheet(60); Thread.Sleep(5000); //.WaitforExcelWebRenderDone2(60);

                // Check if the status "Saving.../Saved" is displayed then refresh page, relogin re-enter data
                int timeout = 0;
                while (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == true && timeout < 6)
                {
                    if (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == true && timeout == 4)
                    {
                        Driver.Browser.Navigate().Refresh();
                        LoginAction.Instance.LoginExcelOnlineSiteAfterRefresh();

                        // re enter data to run report
                        managerName = "John Hancock Investments";
                        fundName = "John Hancock Multimanager Lifetime Portfolios 2035";
                        FundDashboardAction.Instance.InputNameToSearchFund(10, fundName, managerName, fundName, sourceIcon);

                        // Click to expand User Input (section)
                        FADAddInAction.Instance.ClickUserInputSection()
                                               .ClickUserInputSubSection(10, dateSection)
                                               .ClickUserInputSubSection(10, customRiskBenchmarkModelling)
                                               .ClickUserInputSubSection(10, feeModelSection)
                                               .ClickUserInputSubSection(10, liquiditySection);

                        // Check if A red notification is shown then waiting for that is disappeared
                        if (FADAddInAction.Instance.IsElementPresent(FADAddInPage.errorInvalidMessage))
                        {
                            FADAddInAction.Instance.WaitForElementInvisible(10, FADAddInPage.errorInvalidMessage);
                        }

                        /// Liquidity Section 
                        FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, "Lockup", "Hard")
                                               .InputTxtLabelField("Lockup Length (months)", "12")
                                               .ClickAndSelectItemInDropdown(10, "Liquidity Frequency", "Monthly")
                                               .InputTxtLabelField("Investor Gate", "60.55")
                                               .ClickAndSelectItemInDropdown(10, "Sidepocket Probability", "50%")
                                               .InputTxtLabelField("Max % of Sidepocket Permitted", "90.55"); Thread.Sleep(500);
                        FADAddInAction.Instance.WaitForElementInvisible(10, FADAddInPage.errorInvalidMessageContent("All the liquidity parameters need to be populated except for Investor Gate"));

                        // Click on Run button
                        FADAddInAction.Instance.ClickRunButton(10)
                                               .WaitForLoadingIconToDisappear(60, FundDashboardPage.loadingIcon)
                                               .WaitforExcelWebRenderDone2(60);
                        break;
                    }
                    if (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == false)
                    {
                        Thread.Sleep(11000);

                        // Check if the dialog "Invalid sheet" is displayed then click on OK button
                        if (FADAddInAction.Instance.IsElementPresent(NavigationPage.okInvalidSheetButtonExcelOnline))
                        {
                            NavigationAction.Instance.ClicKOkButtonInvalidSheetExcelOnline(10); Thread.Sleep(1000);
                        }
                        Thread.Sleep(1000);
                        break;
                    }
                    if (FADAddInAction.Instance.IsElementPresent(By.XPath(bookSavedTitleLoadDone)) == true && timeout == 5)
                    {
                        Console.WriteLine("timeout - Excel WebRender fail to load done!!!");
                        break;
                    }
                    Thread.Sleep(1000);
                    timeout++;
                }

                // Download Excel Online Report
                NavigationAction.Instance.DownloadExcelOnlineReport(10);

                // Verify Excel file was downloaded successfully
                verifyPoint = LoginAction.Instance.CheckFileDownloadIsComplete(9, ExcelFileDownloadedPath, downloadedfileName);
                verifyPoints.Add(summaryTC = "Verify the '" + downloadedfileName + "' file is downloaded successful", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                //Verify content of the downloaded file by doing compare with the baseline excel files.
                /// Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
                string? fileNameBaseline = null;
                if (urlInstance.Contains("sandbox"))
                {
                    fileNameBaseline = "Book_John2035_TEvestment_Sandbox.xlsx";
                }
                if (urlInstance.Contains("conceptia"))
                {
                    fileNameBaseline = @"Book_John2035_TEvestment_Staging.xlsx";
                }

                // Paste value for Downloaded File
                System.Threading.Thread.Sleep(2000);
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "Raw Data", "B2:F1000", 2, "#,##0.00"); // if get 'Raw Data' 6 digits-- > 6, "#,######0.000000"
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B15:C17", 2, "#,##0.00"); // 'Workbench'-Auto Score 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B20:J39", 2, "#,##0.00"); // 'Workbench'-Manager Alpha 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B42:K59", 2, "#,##0.00"); // 'Workbench'-Metrics Over Selected Time Frames 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "C63:D87", 2, "#,##0.00"); // 'Workbench'-Annualized Return 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "C90:D95", 2, "#,##0.00"); // 'Workbench'-Annualized Return 2 digits
                ExcelComparision.MultiplyPercentAndGetXDigits(ExcelFileDownloadedPath + @"\" + downloadedfileName, "WorkBench", "B98:M117", 2, "#,##0.00"); // 'Workbench'-Net Monthly Return 2 digits

                // Verify the information of Single Public Report Sheet 'Raw Data'
                sheetName = "Raw Data";
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline, sheetName, fileNameBaseline);
                verifyPoints.Add(summaryTC = "Verify the information of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundName + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Verify the information of Single Public Report Sheet 'WorkBench'
                sheetName = "WorkBench";
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline, sheetName, fileNameBaseline);
                verifyPoints.Add(summaryTC = "Verify the information of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundName + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                //Closing Excel Application Process
                foreach (var application in Process.GetProcessesByName("EXCEL"))
                {
                    application.Kill(); System.Threading.Thread.Sleep(1000);
                }

                // Delete Manifest and WorkBook file
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", downloadedfileName);
            }
            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                //Closing Excel Application Process
                foreach (var application in Process.GetProcessesByName("EXCEL"))
                {
                    application.Kill(); System.Threading.Thread.Sleep(1000);
                }

                // Delete Manifest, WorkBook and ExcelCompareDiff file 
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", downloadedfileName);

                // Warning
                ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong! Please check console log.");
            }
            #endregion
        }
    }
}
