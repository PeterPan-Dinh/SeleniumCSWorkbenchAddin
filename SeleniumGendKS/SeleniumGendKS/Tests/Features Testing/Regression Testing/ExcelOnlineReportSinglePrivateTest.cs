using AventStack.ExtentReports;
using NUnit.Framework;
using SeleniumGendKS.Core.BaseTestCase;
using SeleniumGendKS.Core.ExcelComparision;
using SeleniumGendKS.Pages;
using System.Diagnostics;
using System.Reflection;

namespace SeleniumGendKS.Tests.Features_Testing.Regression_Testing
{
    [TestFixture, Order(8)]
    internal class ExcelOnlineReportSinglePrivateTest : BaseTestCase
    {
        // Variables declare
        [Obsolete]
        readonly ExtentReports rep = ExtReportgetInstance();
        private string? summaryTC;
        private string downloadedfileName = LoginPage.downloadedBook;
        ///Determine the folder which Excel files were saved
        private string ExcelFileDownloadedPath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads";
        private string ExcelFileBaselinePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Tests\Documents\Excel files\SinglePrivateReport_Baseline\");
        private const string sheetName = "Single Manager Dashboard Output";

        [Test, Category("Regression Testing")]
        public void TC001_ExcelOnline_SinglePrivateReport()
        {
            #region Variables declare
            string urlInstance = LoginPage.url;
            string cambridgeFund = "VGO Capital Partners";
            const string sourceIcon = "C";
            const string asOfDate = "As of Date";
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Single Private Report - TC001");
            try
            {
                // Delete Manifest, WorkBook and ExcelCompareDiff file 
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", downloadedfileName);

                // Login Excel Online and then add (manifest.xml) FAD Add-in
                LoginAction.Instance.LoginExcelOnlineSiteAndAddInNoGodaddy(urlInstance);

                // Search a Cambridge Fund (VGO) (KS-455 the Firm should be searched in the search bar)
                FundDashboardAction.Instance.InputNameToSearchFund(10, "vgo ")
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResultsWithItemSource(cambridgeFund, sourceIcon))
                                            .ClickFundNameReturnOfResults(10, cambridgeFund, sourceIcon); Thread.Sleep(1000);

                // Check if loading icon shows then wait for it to disappear
                if (FundDashboardAction.Instance.IsElementPresent(FundDashboardPage.loadingIcon))
                {
                    FundDashboardAction.Instance.WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);
                }

                /// User Input
                // Select Report Type = Single Manager Dashboard
                Thread.Sleep(1000);
                FADAddInAction.Instance.ClickToCheckTheCheckbox("Single Manager Dashboard")
                                       .ClickAndSelectItemInDropdown(10, FADAddInPage.dataSource, "Manual"); // --> Select Data Source = Manual
                                       
                // Verify label of fields in Single Manager Dashboard
                verifyPoint = asOfDate == FADAddInAction.Instance.IsDatePickerLabelShown(10);
                verifyPoints.Add(summaryTC = "Verify label date-picker of '" + FADAddInPage.dateSelection + "'is shown correctly: '" + asOfDate + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Input data for Report - Single Manager Dashboard
                FADAddInAction.Instance.WaitForElementVisible(10, FADAddInPage.labelDropdown(FADAddInPage.assetClass))
                                       .InputTxtDatePickerTitle(10, FADAddInPage.dateSelection, "2021", "Jun", "30")
                                       .PageDownToScrollDownPage() //.ClickCRBMAddButton(10)
                                       .CheckIfExistingCRBMThenDeleteAll(10, sourceIcon)
                                       .InputTxtNameCRBMRow(10, 1, "s&p") // input 1st CRBM
                                       .WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("S&P 500 Index"))
                                       .ClickNameCRBMReturnOfResults(10, "S&P 500 Index").WaitForElementInvisible(10, AddEditFundPage.overlayDropdown)
                                       //.InputNumberBetaCRBMRow(10, 1, "1") // --> issue on Conceptia: show alert 'Cannot retrieve ...'
                                       .InputNumberGrossExposureCRBMRow(10, 1, "80")
                                       .ClickAndSelectItemInDropdown(10, FADAddInPage.assetClass, "Buyout")
                                       .ClickAndSelectItemInDropdown(10, FADAddInPage.geography, "Africa")
                                       .ClickRunButton(10) // Click Run button to run Report
                                       .WaitForLoadingIconToDisappear(60, FundDashboardPage.loadingIcon)
                                       .WaitforExcelWebSinglePrivateReportRenderDone(60);

                // Download Excel Online Report
                NavigationAction.Instance.DownloadExcelOnlineReport(10);

                // Verify Excel file was downloaded successfully
                verifyPoint = LoginAction.Instance.CheckFileDownloadIsComplete(9, ExcelFileDownloadedPath, downloadedfileName);
                verifyPoints.Add(summaryTC = "Verify the '"+ downloadedfileName + "' file (VGO) is downloaded successful", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
                string? instanceName = null;
                string? fileNameBaseline_VGO = null;
                string? fileNameBaseline_GSR = null;
                if (urlInstance.Contains("sandbox")) 
                {
                    instanceName = "sandbox";
                    fileNameBaseline_VGO = @"Book_VGO_Sandbox.xlsx";
                    fileNameBaseline_GSR = @"Book_GSR_Sandbox.xlsx";
                }
                if (urlInstance.Contains("conceptia")) 
                {
                    instanceName = "conceptia";
                    fileNameBaseline_VGO = @"Book_VGO_Staging.xlsx";
                    fileNameBaseline_GSR = @"Book_GSR_Staging.xlsx";
                }

                #region Verify the information of Single Manager Dashboard Output Report Excel is generated (VGO - source: manual)
                // Verify Excel format of the downloaded file by doing compare with the baseline excel files.
                verifyPoint = ExcelComparision.IsTheFormatSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline_VGO, sheetName, fileNameBaseline_VGO);
                verifyPoints.Add(summaryTC = "Verify Excel format of Single Manager Dashboard Output Report Excel (source = Manual) is generated correctly. (" + cambridgeFund + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Verify content of the downloaded file by doing compare with the baseline excel files
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline_VGO, sheetName, fileNameBaseline_VGO);
                verifyPoints.Add(summaryTC = "Verify the information of Single Manager Dashboard Output Report Excel (source = Manual) is generated correctly. (" + cambridgeFund + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Delete WorkBook file
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "Book");
                #endregion

                #region Verify the information of Single Manager Dashboard Output Report Excel is generated correctly (1315 - source: Cambridge)
                // Switch to the FAD Add-in frame (to interact with Workbench for Excel Online)
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameIdFADAddInCurrentInstance(instanceName));

                // Search a Cambridge Fund (GSR Ventures)
                cambridgeFund = "GSR Ventures";
                FundDashboardAction.Instance.InputNameToSearchFund(10, "GSR Ve")
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResultsWithItemSource(cambridgeFund, sourceIcon))
                                            .ClickFundNameReturnOfResults(10, cambridgeFund, sourceIcon); Thread.Sleep(1000);

                // Check if loading icon shows then wait for it to disappear
                if (FundDashboardAction.Instance.IsElementPresent(FundDashboardPage.loadingIcon))
                {
                    FundDashboardAction.Instance.WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);
                }

                /// User Input
                // Select Report Type = Single Manager Dashboard
                Thread.Sleep(1000);
                FADAddInAction.Instance.ClickToCheckTheCheckbox("Single Manager Dashboard")
                                       .ClickAndSelectItemInDropdown(10, FADAddInPage.dataSource, "Cambridge"); // --> Select Data Source = Cambridge

                // Input data for Report - Single Manager Dashboard
                FADAddInAction.Instance.InputTxtDatePickerTitle(10, FADAddInPage.dateSelection, "2021", "Jun", "30") // old: "2021", "Dec", "31"
                                       .ClickAndSelectItemInDropdown(10, FADAddInPage.assetClass, "Venture")
                                       .ClickAndSelectItemInDropdown(10, FADAddInPage.geography, "United States")
                                       .ClickRunButton(10) // Click Run button to run Report
                                       .WaitForLoadingIconToDisappear(60, FundDashboardPage.loadingIcon)
                                       .WaitforExcelWebSinglePrivateReportRenderDone(60);

                // Download Excel Online Report
                NavigationAction.Instance.DownloadExcelOnlineReport(10);

                // Verify Excel file was downloaded successfully
                verifyPoint = LoginAction.Instance.CheckFileDownloadIsComplete(9, ExcelFileDownloadedPath, downloadedfileName);
                verifyPoints.Add(summaryTC = "Verify the '" + downloadedfileName + "' file (GSR Ventures) is downloaded successful", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Verify Excel format of the downloaded file by doing compare with the baseline excel files.
                verifyPoint = ExcelComparision.IsTheFormatSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline_GSR, sheetName, fileNameBaseline_GSR);
                verifyPoints.Add(summaryTC = "Verify Excel format of Single Manager Dashboard Output Report Excel (source = Cambridge) is generated correctly. (" + cambridgeFund + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Verify content of the downloaded file by doing compare with the baseline excel files.
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline_GSR, sheetName, fileNameBaseline_GSR);
                verifyPoints.Add(summaryTC = "Verify the information of Single Manager Dashboard Output Report Excel (source = Cambridge) is generated correctly. (" + cambridgeFund + ")", verifyPoint);
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

                // Delete Manifest, WorkBook files 
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
