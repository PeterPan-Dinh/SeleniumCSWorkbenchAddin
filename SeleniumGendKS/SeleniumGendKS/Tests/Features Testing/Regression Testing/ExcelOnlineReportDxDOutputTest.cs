using AventStack.ExtentReports;
using NUnit.Framework;
using OpenQA.Selenium;
using SeleniumGendKS.Core.BaseTestCase;
using SeleniumGendKS.Core.ExcelComparision;
using SeleniumGendKS.Core.Selenium;
using SeleniumGendKS.Pages;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;

namespace SeleniumGendKS.Tests.Features_Testing.Regression_Testing
{
    [TestFixture, Order(7)]
    internal class ExcelOnlineReportDxDOutputTest : BaseTestCase
    {
        // Variables declare
        [Obsolete]
        readonly ExtentReports rep = ExtReportgetInstance();
        private string? summaryTC;
        private string downloadedfileName = LoginPage.downloadedBook;
        ///Determine the folder which Excel files were saved
        private string ExcelFileDownloadedPath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads";
        private string ExcelFileBaselinePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Tests\Documents\Excel files\DxDOutputReport_Baseline\");
        private const string sheetName = "DxD Output";

        [Test, Category("Regression Testing")]
        public void TC001_ExcelOnline_DxDOutputReport()
        {
            #region Variables declare
            string urlInstance = LoginPage.url;
            string cambridgeFund = "VGO Capital Partners";
            const string sourceIcon = "C";
            const string effectiveDate = "Effective Date*";
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - DxD Output Report - TC001");
            try
            {
                // Delete Manifest, WorkBook files
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
                // Select Report Type = Deal by Deal
                Thread.Sleep(1000);
                FADAddInAction.Instance.ClickToCheckTheCheckbox("Deal by Deal");

                // Verify label of fields in Deal by Deal
                verifyPoint = effectiveDate == FADAddInAction.Instance.IsDatePickerLabelShown(10);
                verifyPoints.Add(summaryTC = "Verify label date-picker of '"+ FADAddInPage.dateSelection + "'is shown correctly: '" + effectiveDate + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Input data for Report - Deal by Deal
                FADAddInAction.Instance.WaitForElementVisible(10, FADAddInPage.labelDropdown(FADAddInPage.weightingMethodology))
                                       .InputTxtDatePickerTitle(10, FADAddInPage.dateSelection, "2021", "Jun", "30") //.InputTxtDatePickerLabel(10, dateSelection, "2022", "Jun", "30")
                                       .CheckIfExistingCRBMThenDeleteAll(10, sourceIcon)
                                       .InputTxtNameCRBMRow(10, 1, "s&p") // input 1st CRBM
                                       .WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("S&P 500 Index"))
                                       .ClickNameCRBMReturnOfResults(10, "S&P 500 Index").WaitForElementInvisible(10, AddEditFundPage.overlayDropdown)
                                       .InputNumberBetaCRBMRow(10, 1, "-0.5")
                                       .InputNumberGrossExposureCRBMRow(10, 1, "80")
                                       .ClickCRBMAddButton(10) // Add 2nd CRBM
                                       .InputTxtNameCRBMRow(10, 2, "msci")
                                       .WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("MSCI Emerging Markets Latin America ex Brazil Net TR Index (USD)"))
                                       .ClickNameCRBMReturnOfResults(10, "MSCI Emerging Markets Latin America ex Brazil Net TR Index (USD)").WaitForElementInvisible(10, AddEditFundPage.overlayDropdown)
                                       .InputNumberBetaCRBMRow(10, 2, "1.5")
                                       .InputNumberGrossExposureCRBMRow(10, 2, "-4.5")
                                       .ClickAndSelectItemInDropdown(10, FADAddInPage.weightingMethodology, "Percent Fund Final") 
                                       .ClickAndSelectItemInDropdown(10, FADAddInPage.attributionCategory1, "% of Fund")
                                       .ClickAndSelectItemInDropdown(10, FADAddInPage.attributionCategory2, "Equal Weight")
                                       .ClickAndSelectItemInDropdown(10, FADAddInPage.attributionCategory3, "Invested Capital")
                                       .ClickRunButton(10) // Click Run button to run Report
                                       .WaitForLoadingIconToDisappear(60, FundDashboardPage.loadingIcon)
                                       .WaitforExcelWebDxDOutputRenderDone(60);

                // Download Excel Online Report
                NavigationAction.Instance.DownloadExcelOnlineReport(10);
                
                // Verify Excel file was downloaded successfully
                verifyPoint = LoginAction.Instance.CheckFileDownloadIsComplete(9, ExcelFileDownloadedPath, downloadedfileName);
                verifyPoints.Add(summaryTC = "Verify the '" + downloadedfileName + "' file (Weighting Methodology = Percent Fund Final) is downloaded successful", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
                string? instanceName = null;
                string? fileNameBaseline_WM_Perc = null;
                string? fileNameBaseline_WM_Equal = null;
                string? fileNameBaseline_WM_Custom = null;
                string? fileNameBaseline_Pear_WM_Perc = null;
                if (urlInstance.Contains("sandbox"))
                {
                    instanceName = "sandbox";
                    fileNameBaseline_WM_Perc = @"Book_VGO_Sandbox_WM_Perc.xlsx";
                    //fileNameBaseline_WM_Equal = @"Book_VGO_Sandbox_WM_Equal";
                    //fileNameBaseline_WM_Custom = @"Book_VGO_Sandbox_WM_Custom";
                    //fileNameBaseline_Pear_WM_Perc = @"Book_PEAR_Sandbox_WM_Perc";
                }
                if (urlInstance.Contains("conceptia"))
                {
                    instanceName = "conceptia";
                    fileNameBaseline_WM_Perc = @"Book_VGO_Staging_WM_Perc.xlsx";
                    //fileNameBaseline_WM_Equal = @"Book_VGO_Staging_WM_Equal";
                    //fileNameBaseline_WM_Custom = @"Book_VGO_Staging_WM_Custom";
                    //fileNameBaseline_Pear_WM_Perc = @"Book_PEAR_Staging_WM_Perc";
                }

                // Paste value for Downloaded File
                System.Threading.Thread.Sleep(2000);
                ExcelComparision.MultiplyPercentAndGet2Digits(ExcelFileDownloadedPath + @"\" + downloadedfileName, sheetName, "L2:L4");
                ExcelComparision.MultiplyPercentAndGet2Digits(ExcelFileDownloadedPath + @"\" + downloadedfileName, sheetName, "O3:Q5");
                ExcelComparision.MultiplyPercentAndGet2Digits(ExcelFileDownloadedPath + @"\" + downloadedfileName, sheetName, "O9:Q11");
                ExcelComparision.MultiplyPercentAndGet2Digits(ExcelFileDownloadedPath + @"\" + downloadedfileName, sheetName, "O15:Q15");
                ExcelComparision.MultiplyPercentAndGet2Digits(ExcelFileDownloadedPath + @"\" + downloadedfileName, sheetName, "K9:L13");
                ExcelComparision.MultiplyPercentAndGet2Digits(ExcelFileDownloadedPath + @"\" + downloadedfileName, sheetName, "D13:G14");
                ExcelComparision.MultiplyPercentAndGet2Digits(ExcelFileDownloadedPath + @"\" + downloadedfileName, sheetName, "K16:L23");
                ExcelComparision.MultiplyPercentAndGet2Digits(ExcelFileDownloadedPath + @"\" + downloadedfileName, sheetName, "F28:I52");
                ExcelComparision.MultiplyPercentAndGet2Digits(ExcelFileDownloadedPath + @"\" + downloadedfileName, sheetName, "L28:R52");
                ExcelComparision.SortColumnRange(ExcelFileDownloadedPath + @"\" + downloadedfileName, sheetName, "N3:Q5"); // Sort Attribution 1
                ExcelComparision.SortColumnRange(ExcelFileDownloadedPath + @"\" + downloadedfileName, sheetName, "N9:Q11"); // Sort Attribution 2
                ExcelComparision.SortColumnRange(ExcelFileDownloadedPath + @"\" + downloadedfileName, sheetName, "A28:R50"); // Sort Company Name

                // Verify content of the downloaded file by doing compare with the baseline excel files.
                #region Verify the information of DxD Output Report Excel is generated with Weighting Methodology = Percent Fund Final
                // Verify content of the downloaded file by doing compare with the baseline excel files.
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline_WM_Perc, sheetName, fileNameBaseline_WM_Perc);
                verifyPoints.Add(summaryTC = "Verify the information of DxD Output Report Excel is generated correctly (Weighting Methodology = Percent Fund Final). (" + cambridgeFund + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Verify Excel format of the downloaded file by doing compare with the baseline excel files.
                verifyPoint = ExcelComparision.IsTheFormatSheetsEqual(ExcelFileDownloadedPath + @"\" + downloadedfileName, ExcelFileBaselinePath + fileNameBaseline_WM_Perc, sheetName, fileNameBaseline_WM_Perc);
                verifyPoints.Add(summaryTC = "Verify Excel Format of DxD Output Report Excel is generated correctly (Weighting Methodology = Percent Fund Final). (" + cambridgeFund + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                //// Delete WorkBook file
                //LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "Book");
                #endregion

                #region Verify the information of DxD Output Report Excel is generated with Weighting Methodology = Equal Weight
                /*
                // Switch to the FAD Add-in frame (to interact with Workbench for Excel Online)
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameIdFADAddInCurrentInstance(instanceName));

                // Select Weighting Methodology = Equal Weight
                FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, FADAddInPage.weightingMethodology, "Equal Weight")
                                       .ClickRunButton(10) // Click Run button to run Report
                                       .WaitForLoadingIconToDisappear(60, FundDashboardPage.loadingIcon)
                                       .WaitforExcelWebDxDOutputRenderDone(60);

                // Download Excel Online Report
                NavigationAction.Instance.DownloadExcelOnlineReport(10);

                // Verify Excel file was downloaded successfully
                verifyPoint = LoginAction.Instance.CheckFileDownloadIsComplete(9, ExcelFileDownloadedPath, "Book.xlsx");
                verifyPoints.Add(summaryTC = "Verify the Book.xlsx file (Weighting Methodology = Equal Weight) is downloaded successful", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Verify the information of DxD Output Report Excel is generated with Weighting Methodology = Equal Weight
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + @"Book.xlsx", ExcelFileBaselinePath + fileNameBaseline_WM_Equal, sheetName);
                verifyPoints.Add(summaryTC = "Verify the information of DxD Output Report Excel is generated correctly (Weighting Methodology = Equal Weight). (" + cambridgeFund + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Delete WorkBook file
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "Book");
                */
                #endregion

                #region Verify the information of DxD Output Report Excel is generated with Weighting Methodology = Custom Weight
                /*
                // Switch to the FAD Add-in frame (to interact with Workbench for Excel Online)
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameIdFADAddInCurrentInstance(instanceName));

                // Select Weighting Methodology = Equal Weight
                FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, FADAddInPage.weightingMethodology, "Custom Weight")
                                       .ClickRunButton(10) // Click Run button to run Report
                                       .WaitForLoadingIconToDisappear(60, FundDashboardPage.loadingIcon)
                                       .WaitforExcelWebDxDOutputRenderDone(60);

                // Download Excel Online Report
                NavigationAction.Instance.DownloadExcelOnlineReport(10);

                // Verify Excel file was downloaded successfully
                verifyPoint = LoginAction.Instance.CheckFileDownloadIsComplete(9, ExcelFileDownloadedPath, "Book.xlsx");
                verifyPoints.Add(summaryTC = "Verify the Book.xlsx file (Weighting Methodology = Custom Weight) is downloaded successful", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Verify the information of DxD Output Report Excel is generated with Weighting Methodology = Custom Weight
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + @"Book.xlsx", ExcelFileBaselinePath + fileNameBaseline_WM_Custom, sheetName);
                verifyPoints.Add(summaryTC = "Verify the information of DxD Output Report Excel is generated correctly (Weighting Methodology = Custom Weight). (" + cambridgeFund + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Delete WorkBook file
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "Book");
                */
                #endregion

                #region Verify the information of (PEAR) DxD Output Report Excel is generated with Weighting Methodology = Percent Fund Final
                /*
                // Switch to the FAD Add-in frame (to interact with Workbench for Excel Online)
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameIdFADAddInCurrentInstance(LoginPage.instanceName));

                // Search a Cambridge Fund (Pear Ventures)
                cambridgeFund = "Pear Ventures";
                FundDashboardAction.Instance.InputNameToSearchFund(10, "Pear")
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResultsWithItemSource(cambridgeFund, sourceIcon))
                                            .ClickFundNameReturnOfResults(10, cambridgeFund, sourceIcon)
                                            .WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);

                // User Input
                FADAddInAction.Instance.ClickAndSelectItemInDropdown(10, FADAddInPage.dataSource, "Manual") // --> Select Data Source = Manual
                                       .InputTxtDatePickerTitle(10, FADAddInPage.dateSelection, "2022", "Mar", "31")
                                       .ClickCRBMDeleteButton(10, 2) // --> Delete the 2nd crbm
                                       .WaitForElementInvisible(10, FADAddInPage.nameCRBMInputTxtRow(2))
                                       .InputNumberBetaCRBMRow(10, 1, "1")
                                       .InputNumberGrossExposureCRBMRow(10, 1, "80")
                                       .ClickAndSelectItemInDropdown(10, FADAddInPage.weightingMethodology, "Percent Fund Final") // issue not stable:  Percent Fund Final --> % of Fund
                                       .ClickAndSelectItemInDropdown(10, FADAddInPage.attributionCategory1, "% of Fund")
                                       .ClickAndSelectItemInDropdown(10, FADAddInPage.attributionCategory2, "% of Fund")
                                       .ClickAndSelectItemInDropdown(10, FADAddInPage.attributionCategory3, "% of Fund")
                                       .ClickRunButton(10) // --> Click Run button to run Report
                                       .WaitForLoadingIconToDisappear(60, FundDashboardPage.loadingIcon)
                                       .WaitforExcelWebDxDOutputRenderDone(60);

                // Download Excel Online Report
                NavigationAction.Instance.DownloadExcelOnlineReport(10);

                // Verify Excel file was downloaded successfully
                verifyPoint = LoginAction.Instance.CheckFileDownloadIsComplete(9, ExcelFileDownloadedPath, "Book.xlsx");
                verifyPoints.Add(summaryTC = "Verify the (PEAR) Book.xlsx file (Weighting Methodology = Percent Fund Final) is downloaded successful", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Verify content of the downloaded file by doing compare with the baseline excel files.
                verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + @"Book.xlsx", ExcelFileBaselinePath + fileNameBaseline_Pear_WM_Perc, sheetName);
                verifyPoints.Add(summaryTC = "Verify the information of DxD Output Report Excel is generated correctly (Weighting Methodology = Percent Fund Final). (" + cambridgeFund + ")", verifyPoint);
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
    }
}
