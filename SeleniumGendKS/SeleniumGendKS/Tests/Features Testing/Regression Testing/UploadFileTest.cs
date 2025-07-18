using AventStack.ExtentReports;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using OpenQA.Selenium;
using SeleniumGendKS.Core.BaseTestCase;
using SeleniumGendKS.Core.DBConnection;
using SeleniumGendKS.Core.FilesComparision;
using SeleniumGendKS.Core.Selenium;
using SeleniumGendKS.Pages;
using System.Globalization;
using System.Reflection;
using static System.Windows.Forms.AxHost;

namespace SeleniumGendKS.Tests.Features_Testing.Regression_Testing
{
    [TestFixture]
    internal class UploadFileTest : BaseTestCase
    {
        // Variables declare
        [Obsolete]
        readonly ExtentReports rep = ExtReportgetInstance();
        private string? summaryTC;

        [Test, Category("Regression Testing")]
        public void TC001_UploadFile_ManualFund_CSV_invalid_file()
        {
            #region Variables declare
            string? inputSearch = "qa test 05";
            string? managerNameManual = "QA Test 05";
            string? fundNameManual = "Main QA Test 05";
            const string fileType = "File Type";
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenDKS-UploadFileTest - ManualFund - Invalid CSV File - TC001");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Log into the application
                LoginAction.Instance.LoginSiteNoGodaddy();

                // Search a Manual Fund (QA Test 05)
                FundDashboardAction.Instance.InputNameToSearchFund(10, inputSearch)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(managerNameManual))
                                            .ClickFundNameReturnOfResults(10, managerNameManual)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(fundNameManual))
                                            .ClickFundNameReturnOfResults(10, fundNameManual)
                                            .WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);

                // Verify the fund name (manual fund) is shown correctly after searching
                verifyPoint = fundNameManual == FundDashboardAction.Instance.FundNameTitleGetText(10);
                verifyPoints.Add(summaryTC = "Verify the fund name (Manual Fund) is shown correctly after searching: (" + fundNameManual + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click on Data Status tab
                FundDashboardAction.Instance.ClickDataStatusTab(10);

                // Click on Upload File button
                UploadFileAction.Instance.ClickUploadFileButton(10);

                #region Exposure (Upload invalid file)
                // Input data for File Type (dropdown)
                UploadFileAction.Instance.ClickAndSelectItemInDropdown(10, fileType, "Exposure");

                // Upload file (txt file)
                string fileName = "txt_file.txt";
                string filePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Tests\Documents\Empty files\");
                UploadFileAction.Instance.UploadFileInput(filePath + fileName);

                // Verify a red toast message is shown after uploading an invalid file (txt file) (File Type: Exposure)
                verifyPoint = fileName + ": Invalid file type," == UploadFileAction.Instance.ToastMessageInvalidFileGetText(10)
                           && "allowed file types: .csv." == UploadFileAction.Instance.ToastMessageInvalidFileDetailGetText(10);
                verifyPoints.Add(summaryTC = "Verify a red toast message is shown after uploading an invalid file (Txt) (File Type: Exposure): (" + fileName + ": Invalid file type, allowed file types: .csv.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click (x) to close the red toast message
                UploadFileAction.Instance.ClickCloseToastMessageButton(10)
                                         .WaitForElementInvisible(10, UploadFilePage.toastMessageInvalidFile);

                // Upload file (Excel file)
                fileName = "Excel_file.xlsx";
                UploadFileAction.Instance.UploadFileInput(filePath + fileName);

                // Verify a red toast message is shown after uploading an invalid file (Excel file) (File Type: Exposure)
                verifyPoint = fileName + ": Invalid file type," == UploadFileAction.Instance.ToastMessageInvalidFileGetText(10)
                           && "allowed file types: .csv." == UploadFileAction.Instance.ToastMessageInvalidFileDetailGetText(10);
                verifyPoints.Add(summaryTC = "Verify a red toast message is shown after uploading an invalid file (Excel) (File Type: Exposure): (" + fileName + ": Invalid file type, allowed file types: .csv.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click (x) to close the red toast message
                UploadFileAction.Instance.ClickCloseToastMessageButton(10)
                                         .WaitForElementInvisible(10, UploadFilePage.toastMessageInvalidFile);

                // Upload file (Word file)
                fileName = "Word_file.docx";
                UploadFileAction.Instance.UploadFileInput(filePath + fileName);

                // Verify a red toast message is shown after uploading an invalid file (Word file) (File Type: Exposure)
                verifyPoint = fileName + ": Invalid file type," == UploadFileAction.Instance.ToastMessageInvalidFileGetText(10)
                           && "allowed file types: .csv." == UploadFileAction.Instance.ToastMessageInvalidFileDetailGetText(10);
                verifyPoints.Add(summaryTC = "Verify a red toast message is shown after uploading an invalid file (Word) (File Type: Exposure): (" + fileName + ": Invalid file type, allowed file types: .csv.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                #endregion

                #region Fund AUM (Upload invalid file)
                // Click (x) to close the red toast message
                UploadFileAction.Instance.ClickCloseToastMessageButton(10)
                                         .WaitForElementInvisible(10, UploadFilePage.toastMessageInvalidFile);

                // Input data for File Type (dropdown)
                UploadFileAction.Instance.ClickAndSelectItemInDropdown(10, fileType, "Fund AUM");

                // Upload file (txt file)
                fileName = "txt_file.txt";
                filePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Tests\Documents\Empty files\");
                UploadFileAction.Instance.UploadFileInput(filePath + fileName);

                // Verify a red toast message is shown after uploading an invalid file (txt file) (File Type: Fund AUM)
                verifyPoint = fileName + ": Invalid file type," == UploadFileAction.Instance.ToastMessageInvalidFileGetText(10)
                           && "allowed file types: .csv." == UploadFileAction.Instance.ToastMessageInvalidFileDetailGetText(10);
                verifyPoints.Add(summaryTC = "Verify a red toast message is shown after uploading an invalid file (Txt) (File Type: Fund AUM): (" + fileName + ": Invalid file type, allowed file types: .csv.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click (x) to close the red toast message
                UploadFileAction.Instance.ClickCloseToastMessageButton(10)
                                         .WaitForElementInvisible(10, UploadFilePage.toastMessageInvalidFile);

                // Upload file (Excel file)
                fileName = "Excel_file.xlsx";
                UploadFileAction.Instance.UploadFileInput(filePath + fileName);

                // Verify a red toast message is shown after uploading an invalid file (Excel file) (File Type: Fund AUM)
                verifyPoint = fileName + ": Invalid file type," == UploadFileAction.Instance.ToastMessageInvalidFileGetText(10)
                           && "allowed file types: .csv." == UploadFileAction.Instance.ToastMessageInvalidFileDetailGetText(10);
                verifyPoints.Add(summaryTC = "Verify a red toast message is shown after uploading an invalid file (Excel) (File Type: Fund AUM): (" + fileName + ": Invalid file type, allowed file types: .csv.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click (x) to close the red toast message
                UploadFileAction.Instance.ClickCloseToastMessageButton(10)
                                         .WaitForElementInvisible(10, UploadFilePage.toastMessageInvalidFile);

                // Upload file (Word file)
                fileName = "Word_file.docx";
                UploadFileAction.Instance.UploadFileInput(filePath + fileName);

                // Verify a red toast message is shown after uploading an invalid file (Word file) (File Type: Fund AUM)
                verifyPoint = fileName + ": Invalid file type," == UploadFileAction.Instance.ToastMessageInvalidFileGetText(10)
                           && "allowed file types: .csv." == UploadFileAction.Instance.ToastMessageInvalidFileDetailGetText(10);
                verifyPoints.Add(summaryTC = "Verify a red toast message is shown after uploading an invalid file (Word) (File Type: Fund AUM): (" + fileName + ": Invalid file type, allowed file types: .csv.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                #endregion

                #region Fund Returns (Upload invalid file)
                // Click (x) to close the red toast message
                UploadFileAction.Instance.ClickCloseToastMessageButton(10)
                                         .WaitForElementInvisible(10, UploadFilePage.toastMessageInvalidFile);

                // Input data for File Type (dropdown)
                UploadFileAction.Instance.ClickAndSelectItemInDropdown(10, fileType, "Fund Returns");

                // Upload file (txt file)
                fileName = "txt_file.txt";
                filePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Tests\Documents\Empty files\");
                UploadFileAction.Instance.UploadFileInput(filePath + fileName);

                // Verify a red toast message is shown after uploading an invalid file (txt file) (File Type: Fund Returns)
                verifyPoint = fileName + ": Invalid file type," == UploadFileAction.Instance.ToastMessageInvalidFileGetText(10)
                           && "allowed file types: .csv." == UploadFileAction.Instance.ToastMessageInvalidFileDetailGetText(10);
                verifyPoints.Add(summaryTC = "Verify a red toast message is shown after uploading an invalid file (Txt) (File Type: Fund Returns): (" + fileName + ": Invalid file type, allowed file types: .csv.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click (x) to close the red toast message
                UploadFileAction.Instance.ClickCloseToastMessageButton(10)
                                         .WaitForElementInvisible(10, UploadFilePage.toastMessageInvalidFile);

                // Upload file (Excel file)
                fileName = "Excel_file.xlsx";
                UploadFileAction.Instance.UploadFileInput(filePath + fileName);

                // Verify a red toast message is shown after uploading an invalid file (Excel file) (File Type: Fund Returns)
                verifyPoint = fileName + ": Invalid file type," == UploadFileAction.Instance.ToastMessageInvalidFileGetText(10)
                           && "allowed file types: .csv." == UploadFileAction.Instance.ToastMessageInvalidFileDetailGetText(10);
                verifyPoints.Add(summaryTC = "Verify a red toast message is shown after uploading an invalid file (Excel) (File Type: Fund Returns): (" + fileName + ": Invalid file type, allowed file types: .csv.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click (x) to close the red toast message
                UploadFileAction.Instance.ClickCloseToastMessageButton(10)
                                         .WaitForElementInvisible(10, UploadFilePage.toastMessageInvalidFile);

                // Upload file (Word file)
                fileName = "Word_file.docx";
                UploadFileAction.Instance.UploadFileInput(filePath + fileName);

                // Verify a red toast message is shown after uploading an invalid file (Word file) (File Type: Fund Returns)
                verifyPoint = fileName + ": Invalid file type," == UploadFileAction.Instance.ToastMessageInvalidFileGetText(10)
                           && "allowed file types: .csv." == UploadFileAction.Instance.ToastMessageInvalidFileDetailGetText(10);
                verifyPoints.Add(summaryTC = "Verify a red toast message is shown after uploading an invalid file (Word) (File Type: Fund Returns): (" + fileName + ": Invalid file type, allowed file types: .csv.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
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

        [Test, Category("Regression Testing"), Ignore("")]
        public void TC002_UploadFile_PDF_CRBM_invalid_file()
        {
            #region Variables declare
            const string fileType = "File Type";
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenDKS-UploadFileTest - PDF-CRBM - Invalid File - TC002");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Log into the application
                LoginAction.Instance.LoginSiteNoGodaddy();

                // Click on "PDF" upload icon
                LoginAction.Instance.ClickUploadPDFFilesIcon(10);

                #region Historical Exposure - PDF (Upload invalid file)
                // Input data for File Type (dropdown)
                UploadFileAction.Instance.ClickAndSelectItemInDropdown(10, fileType, "Historical Exposure - PDF");

                // Upload file (txt file)
                string fileName = "txt_file.txt";
                string filePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Tests\Documents\Empty files\");
                UploadFileAction.Instance.UploadFileInput(filePath + fileName);

                // Verify a red toast message is shown after uploading an invalid file (txt file) (File Type: Historical Exposure - PDF)
                verifyPoint = fileName + ": Invalid file type," == UploadFileAction.Instance.ToastMessageInvalidFileGetText(10)
                           && "allowed file types: .pdf." == UploadFileAction.Instance.ToastMessageInvalidFileDetailGetText(10);
                verifyPoints.Add(summaryTC = "Verify a red toast message is shown after uploading an invalid file (Txt) (File Type: Historical Exposure - PDF): (" + fileName + ": Invalid file type, allowed file types: .pdf.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click (x) to close the red toast message
                UploadFileAction.Instance.ClickCloseToastMessageButton(10)
                                         .WaitForElementInvisible(10, UploadFilePage.toastMessageInvalidFile);

                // Upload file (Excel file)
                fileName = "Excel_file.xlsx";
                UploadFileAction.Instance.UploadFileInput(filePath + fileName);

                // Verify a red toast message is shown after uploading an invalid file (Excel file) (File Type: Historical Exposure - PDF)
                verifyPoint = fileName + ": Invalid file type," == UploadFileAction.Instance.ToastMessageInvalidFileGetText(10)
                           && "allowed file types: .pdf." == UploadFileAction.Instance.ToastMessageInvalidFileDetailGetText(10);
                verifyPoints.Add(summaryTC = "Verify a red toast message is shown after uploading an invalid file (Excel) (File Type: Historical Exposure - PDF): (" + fileName + ": Invalid file type, allowed file types: .pdf.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click (x) to close the red toast message
                UploadFileAction.Instance.ClickCloseToastMessageButton(10)
                                         .WaitForElementInvisible(10, UploadFilePage.toastMessageInvalidFile);

                // Upload file (Word file)
                fileName = "Word_file.docx";
                UploadFileAction.Instance.UploadFileInput(filePath + fileName);

                // Verify a red toast message is shown after uploading an invalid file (Word file) (File Type: Historical Exposure - PDF)
                verifyPoint = fileName + ": Invalid file type," == UploadFileAction.Instance.ToastMessageInvalidFileGetText(10)
                           && "allowed file types: .pdf." == UploadFileAction.Instance.ToastMessageInvalidFileDetailGetText(10);
                verifyPoints.Add(summaryTC = "Verify a red toast message is shown after uploading an invalid file (Word) (File Type: Historical Exposure - PDF): (" + fileName + ": Invalid file type, allowed file types: .pdf.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                #endregion

                #region Performance History - PDF (Upload invalid file)
                // Click (x) to close the red toast message
                UploadFileAction.Instance.ClickCloseToastMessageButton(10)
                                         .WaitForElementInvisible(10, UploadFilePage.toastMessageInvalidFile);

                // Input data for File Type (dropdown)
                UploadFileAction.Instance.ClickAndSelectItemInDropdown(10, fileType, "Performance History - PDF");

                // Upload file (txt file)
                fileName = "txt_file.txt";
                UploadFileAction.Instance.UploadFileInput(filePath + fileName);

                // Verify a red toast message is shown after uploading an invalid file (txt file) (File Type: Performance History - PDF)
                verifyPoint = fileName + ": Invalid file type," == UploadFileAction.Instance.ToastMessageInvalidFileGetText(10)
                           && "allowed file types: .pdf." == UploadFileAction.Instance.ToastMessageInvalidFileDetailGetText(10);
                verifyPoints.Add(summaryTC = "Verify a red toast message is shown after uploading an invalid file (Txt) (File Type: Performance History - PDF): (" + fileName + ": Invalid file type, allowed file types: .pdf.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click (x) to close the red toast message
                UploadFileAction.Instance.ClickCloseToastMessageButton(10)
                                         .WaitForElementInvisible(10, UploadFilePage.toastMessageInvalidFile);

                // Upload file (Excel file)
                fileName = "Excel_file.xlsx";
                UploadFileAction.Instance.UploadFileInput(filePath + fileName);

                // Verify a red toast message is shown after uploading an invalid file (Excel file) (File Type: Performance History - PDF)
                verifyPoint = fileName + ": Invalid file type," == UploadFileAction.Instance.ToastMessageInvalidFileGetText(10)
                           && "allowed file types: .pdf." == UploadFileAction.Instance.ToastMessageInvalidFileDetailGetText(10);
                verifyPoints.Add(summaryTC = "Verify a red toast message is shown after uploading an invalid file (Excel) (File Type: Performance History - PDF): (" + fileName + ": Invalid file type, allowed file types: .pdf.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click (x) to close the red toast message
                UploadFileAction.Instance.ClickCloseToastMessageButton(10)
                                         .WaitForElementInvisible(10, UploadFilePage.toastMessageInvalidFile);

                // Upload file (Word file)
                fileName = "Word_file.docx";
                UploadFileAction.Instance.UploadFileInput(filePath + fileName);

                // Verify a red toast message is shown after uploading an invalid file (Word file) (File Type: Performance History - PDF)
                verifyPoint = fileName + ": Invalid file type," == UploadFileAction.Instance.ToastMessageInvalidFileGetText(10)
                           && "allowed file types: .pdf." == UploadFileAction.Instance.ToastMessageInvalidFileDetailGetText(10);
                verifyPoints.Add(summaryTC = "Verify a red toast message is shown after uploading an invalid file (Word) (File Type: Performance History - PDF): (" + fileName + ": Invalid file type, allowed file types: .pdf.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                #endregion

                #region INDEX - XLSX (Upload invalid file)
                // Click (x) to close the red toast message
                UploadFileAction.Instance.ClickCloseToastMessageButton(10)
                                         .WaitForElementInvisible(10, UploadFilePage.toastMessageInvalidFile);

                // Input data for File Type (dropdown)
                UploadFileAction.Instance.ClickAndSelectItemInDropdown(10, fileType, "INDEX - XLSX");

                // Upload file (txt file)
                fileName = "txt_file.txt";
                UploadFileAction.Instance.UploadFileInput(filePath + fileName);

                // Verify a red toast message is shown after uploading an invalid file (txt file) (File Type: CRBM - XLSX)
                verifyPoint = fileName + ": Invalid file type," == UploadFileAction.Instance.ToastMessageInvalidFileGetText(10)
                           && "allowed file types: .xls,.xlsx." == UploadFileAction.Instance.ToastMessageInvalidFileDetailGetText(10);
                verifyPoints.Add(summaryTC = "Verify a red toast message is shown after uploading an invalid file (Txt) (File Type: CRBM - XLSX): (" + fileName + ": Invalid file type, allowed file types: .xls,.xlsx.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click (x) to close the red toast message
                UploadFileAction.Instance.ClickCloseToastMessageButton(10)
                                         .WaitForElementInvisible(10, UploadFilePage.toastMessageInvalidFile);

                // Upload file (Word file)
                fileName = "Word_file.docx";
                UploadFileAction.Instance.UploadFileInput(filePath + fileName);

                // Verify a red toast message is shown after uploading an invalid file (Word file) (File Type: CRBM - XLSX)
                verifyPoint = fileName + ": Invalid file type," == UploadFileAction.Instance.ToastMessageInvalidFileGetText(10)
                           && "allowed file types: .xls,.xlsx." == UploadFileAction.Instance.ToastMessageInvalidFileDetailGetText(10);
                verifyPoints.Add(summaryTC = "Verify a red toast message is shown after uploading an invalid file (Word) (File Type: CRBM - XLSX): (" + fileName + ": Invalid file type, allowed file types: .xls,.xlsx.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
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

        [Test, Category("Regression Testing")]
        public void TC003_UploadFile_Cambridge_invalid_file()
        {
            #region Variables declare
            string? inputSearch = "agfe";
            string? firmCambridge = "AgFe";
            const string fileType = "File Type";
            const string currency = "Currency";
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenDKS-UploadFileTest - Cambridge Invalid File - TC003");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Log into the application
                LoginAction.Instance.LoginSiteNoGodaddy();

                // Search a Cambridge Fund (AgFe)
                FundDashboardAction.Instance.InputNameToSearchFund(10, inputSearch)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(firmCambridge))
                                            .ClickFundNameReturnOfResults(10, firmCambridge); Thread.Sleep(1000);

                // Check if loading icon shows then wait for it to disappear
                if (FundDashboardAction.Instance.IsElementPresent(FundDashboardPage.loadingIcon))
                {
                    FundDashboardAction.Instance.WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);
                }

                // Verify the Firm name (Cambridge) is shown correctly after searching
                Thread.Sleep(1000);
                verifyPoint = firmCambridge == FundDashboardAction.Instance.FundNameTitleGetText(10);
                verifyPoints.Add(summaryTC = "Verify the Firm Name (Cambridge) is shown correctly after searching: (" + firmCambridge + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click on Data Status tab
                FundDashboardAction.Instance.ClickDataStatusTab(10);

                // Click on Upload File button
                UploadFileAction.Instance.ClickUploadFileButton(10);

                // Input data for File Type (dropdown), Currency (dropdown), As of date (date-pickker)
                UploadFileAction.Instance.ClickAndSelectItemInDropdown(10, fileType, "Fund Information")
                                         .ClickAndSelectItemInDropdown(10, currency, "VND")
                                         .InputAsOfDate(10, "05-04-2022");

                // Upload file (txt file)
                string fileName = "txt_file.txt";
                string filePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Tests\Documents\Empty files\");
                UploadFileAction.Instance.UploadFileInput(filePath + fileName);

                // Verify a red toast message is shown after uploading an invalid file (txt file)
                verifyPoint = fileName + ": Invalid file type," == UploadFileAction.Instance.ToastMessageInvalidFileGetText(10)
                           && "allowed file types: .csv." == UploadFileAction.Instance.ToastMessageInvalidFileDetailGetText(10);
                verifyPoints.Add(summaryTC = "Verify a red toast message is shown after uploading an invalid file (Txt): (" + fileName + ": Invalid file type, allowed file types: .csv.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click (x) to close the red toast message
                UploadFileAction.Instance.ClickCloseToastMessageButton(10)
                                         .WaitForElementInvisible(10, UploadFilePage.toastMessageInvalidFile);

                // Upload file (Excel file)
                fileName = "Excel_file.xlsx";
                UploadFileAction.Instance.UploadFileInput(filePath + fileName);

                // Verify a red toast message is shown after uploading an invalid file (Excel file)
                verifyPoint = fileName + ": Invalid file type," == UploadFileAction.Instance.ToastMessageInvalidFileGetText(10)
                           && "allowed file types: .csv." == UploadFileAction.Instance.ToastMessageInvalidFileDetailGetText(10);
                verifyPoints.Add(summaryTC = "Verify a red toast message is shown after uploading an invalid file (Excel): (" + fileName + ": Invalid file type, allowed file types: .csv.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click (x) to close the red toast message
                UploadFileAction.Instance.ClickCloseToastMessageButton(10)
                                         .WaitForElementInvisible(10, UploadFilePage.toastMessageInvalidFile);

                // Upload file (Word file)
                fileName = "Word_file.docx";
                UploadFileAction.Instance.UploadFileInput(filePath + fileName);

                // Verify a red toast message is shown after uploading an invalid file (Word file)
                verifyPoint = fileName + ": Invalid file type," == UploadFileAction.Instance.ToastMessageInvalidFileGetText(10)
                           && "allowed file types: .csv." == UploadFileAction.Instance.ToastMessageInvalidFileDetailGetText(10);
                verifyPoints.Add(summaryTC = "Verify a red toast message is shown after uploading an invalid file (Word): (" + fileName + ": Invalid file type, allowed file types: .csv.)", verifyPoint);
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
        public void TC004_UploadFile_ManualFund_CSV_Valid_file_Exposure()
        {
            // Notes: Please turn on your VPN before running this test cases

            #region Variables declare
            string urlInstance = LoginPage.url;
            string? inputSearch = null;
            string? managerNameManual = null;
            string? fundNameManual = null;
            int? fundid = null;
            string exposure = FundDashboardPage.exposure;          
            string rowCounts = FundDashboardPage.rowCounts; //string rowCount = FundDashboardPage.rowCount;
            string startDate = FundDashboardPage.startDate;
            string endDate = FundDashboardPage.endDate; //string asOfDate = FundDashboardPage.asOfDate;
            const string fileType = "File Type";
            string fileName = "Upload_Exposure_2rows.csv";
            string filePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Tests\Documents\CSV files\");
            string timestamp = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff");
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenDKS-UploadFileTest - ManualFund - (Exposure) Valid CSV File - TC004");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Log into the application
                LoginAction.Instance.LoginSiteNoGodaddy();

                // Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
                string? exposureNRowCountVal = null;
                string? exposureStartDateVal = null;
                string? exposureEndDateVal = null;
                if (urlInstance.Contains("sandbox"))
                {
                    inputSearch = "qa test 04";
                    managerNameManual = "QA Test 04";
                    fundNameManual = "Main Fund of QA Test 04 update 01";
                    fundid = 19;
                    exposureNRowCountVal = "2";
                    exposureStartDateVal = "03/31/2019";
                    exposureEndDateVal = "04/30/2020";
                }
                if (urlInstance.Contains("conceptia"))
                {
                    inputSearch = "qa test 05";
                    managerNameManual = "QA Test 05";
                    fundNameManual = "Main QA Test 05";
                    fundid = 12;
                    exposureNRowCountVal = "2";
                    exposureStartDateVal = "03/31/2019";
                    exposureEndDateVal = "04/30/2020";
                }

                // Search a Manual Fund (QA Test 04)
                FundDashboardAction.Instance.InputNameToSearchFund(10, inputSearch)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(managerNameManual))
                                            .ClickFundNameReturnOfResults(10, managerNameManual)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(fundNameManual))
                                            .ClickFundNameReturnOfResults(10, fundNameManual)
                                            .WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);

                // Verify the fund name (manual fund) is shown correctly after searching
                verifyPoint = fundNameManual == FundDashboardAction.Instance.FundNameTitleGetText(10);
                verifyPoints.Add(summaryTC = "Verify the fund name (Manual Fund) is shown correctly after searching: (" + fundNameManual + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                #region Upload CSV File (Exposure)
                // Click on Data Status tab
                FundDashboardAction.Instance.ClickDataStatusTab(10);

                // Click on Upload File button
                UploadFileAction.Instance.ClickUploadFileButton(10);

                // Input data for File Type (dropdown)
                UploadFileAction.Instance.ClickAndSelectItemInDropdown(10, fileType, "Exposure");

                // Upload file (csv file - Exposure)
                UploadFileAction.Instance.UploadFileInput(filePath + fileName);

                // Way 1:
                // Fields Mapping - Select fields for 'Source' and 'Destination'
                UploadFileAction.Instance.ClickDestinationDropdown(10, 1)
                                         .WaitForAllElementsVisible(10, UploadFilePage.overlayDropdown)
                                         .PressDownArrowKeyUntilElementIsVisible(20);

                // Select item for 'Destination' dropdown
                UploadFileAction.Instance.WaitForAllElementsVisible(10, UploadFilePage.destinationDropdownItemSelect(1, UploadFilePage.date))
                                         .SelectItemInDestinationDropdown(10, 1, UploadFilePage.date)
                                         .ClickAndSelectItemInDestinationDropdown(10, 2, UploadFilePage.exposure)
                                         .ClickAndSelectItemInDestinationDropdown(10, 3, UploadFilePage.currency)
                                         .ClickAndSelectItemInDestinationDropdown(10, 4, UploadFilePage.share_class_id);

                // Click Replace button
                UploadFileAction.Instance.ClickUploadLabelButton(10, "Replace") //.ClickDoneButton();
                                         .WaitForAllElementsVisible(10, UploadFilePage.dialogWarning)
                                         .ClickButtonInDialog(10, "Replace");

                // Verify a green toast message is shown after uploading a valid file  (File Type: Exposure)
                verifyPoint = "Success" == UploadFileAction.Instance.toastMessageAlertGetText(10, 1)
                            && "File uploaded successfully." == UploadFileAction.Instance.toastMessageAlertGetText(10, 2);
                verifyPoints.Add(summaryTC = "Verify a green toast message is shown after uploading a valid file (File Type: Exposure): (Success File uploaded successfully.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Wait for toast message is disappeared
                UploadFileAction.Instance.WaitForElementInvisible(10, UploadFilePage.toastMessage(1));
                #endregion

                #region Verify the database value in the table after uploading CSV file into DB
                const string db = "ks_model";
                const string table = "manual_fund_exposure";
                const string dataSource = "manual";
                string query = "SELECT * FROM " + db + "." + table + " WHERE fund_id=" + fundid + " AND _time_ >= timestamp '" + timestamp + "'";

                // Wait until all files are loaded
                System.Threading.Thread.Sleep(10000);

                // Send API Requests and Get API Responses (to retrieve data)
                var sendResponse = DatabaseConnection.GetDatalakeANSISQLQueries(db, query);
                List<JObject> jsonDB = JsonConvert.DeserializeObject<List<JObject>>(sendResponse.Content);

                // Convert csv (file) to json object (--> csv from file PATH)
                var convert = DatabaseConnection.ConvertCsvFileToJsonObject(filePath + fileName);
                List<JObject> jsFile = JsonConvert.DeserializeObject<List<JObject>>(convert);

                // Verify number of rows are loaded
                verifyPoint = jsonDB.Count == jsFile.Count;
                verifyPoints.Add(summaryTC = "Verify number of rows (Exposure) are loaded = " + jsFile.Count + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // JsonFile (csv file) - Rename Keys (field name) and change Date format in JObject
                DatabaseConnection.ChangeDateFormatInJObject(jsFile, "Year", "yyyy-MM-dd HH:mm:ss.fff");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Year", "date");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "ValExpos", "exposure");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Cur", "currency");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Share Cl ID", "share_class_id");
                DatabaseConnection.ChangeToDoubleTypeForValueInJObject(jsFile, "exposure");
                DatabaseConnection.AddFieldNameInJObject(jsFile, "fund_id", fundid);
                DatabaseConnection.AddFieldNameInJObject(jsFile, "fund_source", dataSource);
                var testFile = DatabaseConnection.SortPropertiesByName(jsFile);

                // JsonDB (Database) - Rename Keys (field name) and change Date format in JObject
                DatabaseConnection.RemoveFieldNameInJObject(jsonDB, "_time_");
                DatabaseConnection.RemoveFieldNameInJObject(jsonDB, "create_at");
                DatabaseConnection.RemoveFieldNameInJObject(jsonDB, "import_log");
                var testDB = DatabaseConnection.SortPropertiesByName(jsonDB);

                // Sort value in JObject
                var fileSort = testFile.OrderBy(p => (string?)p["exposure"]);
                var dbSort = testDB.OrderBy(p => (string?)p["exposure"]);

                // Verify data from csv file was loaded into DB correctly
                verifyPoint = UploadFileAction.Instance.CompareObjects(fileSort, dbSort);
                verifyPoints.Add(summaryTC = "Verify data (Exposure) from csv file was loaded into DB correctly", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                #endregion

                #region Verify all informations from Data Status
                // Click on Kamehameha icon to reload/refersh Workbench
                LoginAction.Instance.ClickKamehamehaLogo(10);

                // Wait for Data Status tab is disappeared
                LoginAction.Instance.WaitForElementInvisible(10, FundDashboardPage.dataStatusTab);

                // Search a Manual Fund (QA Test 04)
                FundDashboardAction.Instance.InputNameToSearchFund(10, inputSearch)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(managerNameManual))
                                            .ClickFundNameReturnOfResults(10, managerNameManual)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(fundNameManual))
                                            .ClickFundNameReturnOfResults(10, fundNameManual)
                                            .WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);

                // Click on Data Status tab
                FundDashboardAction.Instance.ClickDataStatusTab(10);

                // verify all informations from Data Status - Exposure
                verifyPoint = FundDashboardAction.Instance.SectionTitleContentGetText(20, exposure, rowCounts) == exposureNRowCountVal;
                verifyPoints.Add(summaryTC = "Verify Data Status - " + FundDashboardPage.exposure + " (Manual) " + FundDashboardPage.rowCounts + " is shown correctly: " + exposureNRowCountVal + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = FundDashboardAction.Instance.SectionTitleContentGetText(10, exposure, startDate) == exposureStartDateVal;
                verifyPoints.Add(summaryTC = "Verify Data Status - " + FundDashboardPage.exposure + " (Manual) " + FundDashboardPage.startDate + " is shown correctly: " + exposureStartDateVal + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = FundDashboardAction.Instance.SectionTitleContentGetText(10, exposure, endDate) == exposureEndDateVal;
                verifyPoints.Add(summaryTC = "Verify Data Status - " + FundDashboardPage.exposure + " (Manual) " + FundDashboardPage.endDate + " is shown correctly: " + exposureEndDateVal + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
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

        [Test, Category("Regression Testing")]
        public void TC005_UploadFile_ManualFund_CSV_Valid_file_FundAUM()
        {
            // Notes: Please turn on your VPN before running this test cases

            #region Variables declare
            string urlInstance = LoginPage.url;
            string? inputSearch = null;
            string? managerNameManual = null;
            string? fundNameManual = null;
            int? fundid = null;
            string fundAUM = FundDashboardPage.fundAum;
            string rowCounts = FundDashboardPage.rowCounts;
            string startDate = FundDashboardPage.startDate;
            string endDate = FundDashboardPage.endDate;
            const string fileType = "File Type";
            string fileName = "Upload_FundAUM_2rows.csv";
            string filePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Tests\Documents\CSV files\");
            string timestamp = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff");
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenDKS-UploadFileTest - ManualFund - (Fund AUM) Valid CSV File - TC005");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Log into the application
                LoginAction.Instance.LoginSiteNoGodaddy();

                // Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
                string? fundAUMRowCountsVal = null;
                string? fundAUMStartDateVal = null;
                string? fundAUMEndDateVal = null;
                if (urlInstance.Contains("sandbox"))
                {
                    inputSearch = "qa test 04";
                    managerNameManual = "QA Test 04";
                    fundNameManual = "Main Fund of QA Test 04 update 01";
                    fundid = 19;
                    fundAUMRowCountsVal = "2";
                    fundAUMStartDateVal = "03/31/2020";
                    fundAUMEndDateVal = "04/30/2021";
                }
                if (urlInstance.Contains("conceptia"))
                {
                    inputSearch = "qa test 05";
                    managerNameManual = "QA Test 05";
                    fundNameManual = "Main QA Test 05";
                    fundid = 12;
                    fundAUMRowCountsVal = "2";
                    fundAUMStartDateVal = "03/31/2020";
                    fundAUMEndDateVal = "04/30/2021";
                }

                // Search a Manual Fund (QA Test 04)
                FundDashboardAction.Instance.InputNameToSearchFund(10, inputSearch)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(managerNameManual))
                                            .ClickFundNameReturnOfResults(10, managerNameManual)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(fundNameManual))
                                            .ClickFundNameReturnOfResults(10, fundNameManual)
                                            .WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);

                // Verify the fund name (manual fund) is shown correctly after searching
                verifyPoint = fundNameManual == FundDashboardAction.Instance.FundNameTitleGetText(10);
                verifyPoints.Add(summaryTC = "Verify the fund name (Manual Fund) is shown correctly after searching: (" + fundNameManual + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                #region Upload CSV File (Fund AUM)
                // Click on Data Status tab
                FundDashboardAction.Instance.ClickDataStatusTab(10);

                // Click on Upload File button
                UploadFileAction.Instance.ClickUploadFileButton(10);

                // Input data for File Type (dropdown)
                UploadFileAction.Instance.ClickAndSelectItemInDropdown(10, fileType, "Fund AUM");

                // Upload file (csv file - Exposure)
                UploadFileAction.Instance.UploadFileInput(filePath + fileName);

                // Fields Mapping - Select fields for 'Source' and 'Destination'
                UploadFileAction.Instance.ClickDestinationDropdown(10, 1)
                                         .WaitForAllElementsVisible(10, UploadFilePage.overlayDropdown)
                                         .PressDownArrowKeyUntilElementIsVisible(20);

                // Select item for 'Destination' dropdown
                UploadFileAction.Instance.WaitForAllElementsVisible(10, UploadFilePage.destinationDropdownItemSelect(1, UploadFilePage.date))
                                         .SelectItemInDestinationDropdown(10, 1, UploadFilePage.date)
                                         .ClickAndSelectItemInDestinationDropdown(10, 2, UploadFilePage.aum)
                                         .ClickAndSelectItemInDestinationDropdown(10, 3, UploadFilePage.currency)
                                         .ClickAndSelectItemInDestinationDropdown(10, 4, UploadFilePage.share_class_id);

                // Click Replace button
                UploadFileAction.Instance.ClickUploadLabelButton(10, "Replace") //.ClickDoneButton();
                                         .WaitForAllElementsVisible(10, UploadFilePage.dialogWarning)
                                         .ClickButtonInDialog(10, "Replace");

                // Verify a green toast message is shown after uploading a valid file  (File Type: Fund AUM)
                verifyPoint = "Success" == UploadFileAction.Instance.toastMessageAlertGetText(10, 1)
                            && "File uploaded successfully." == UploadFileAction.Instance.toastMessageAlertGetText(10, 2);
                verifyPoints.Add(summaryTC = "Verify a green toast message is shown after uploading a valid file (File Type: Fund AUM): (Success File uploaded successfully.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Wait for toast message is disappeared
                UploadFileAction.Instance.WaitForElementInvisible(10, UploadFilePage.toastMessage(1));
                #endregion

                #region Verify the database value in the table after uploading CSV file into DB
                const string db = "ks_model";
                const string table = "manual_fund_aum";
                const string dataSource = "manual";
                string query = "SELECT * FROM " + db + "." + table + " WHERE fund_id=" + fundid + " AND _time_ >= timestamp '" + timestamp + "'";

                // Wait until all files are loaded
                System.Threading.Thread.Sleep(10000);

                // Send API Requests and Get API Responses (to retrieve data)
                var sendResponse = DatabaseConnection.GetDatalakeANSISQLQueries(db, query);
                List<JObject> jsonDB = JsonConvert.DeserializeObject<List<JObject>>(sendResponse.Content);

                // Convert csv (file) to json object (--> csv from file PATH)
                var convert = DatabaseConnection.ConvertCsvFileToJsonObject(filePath + fileName);
                List<JObject> jsFile = JsonConvert.DeserializeObject<List<JObject>>(convert);

                // Verify number of rows are loaded
                verifyPoint = jsonDB.Count == jsFile.Count;
                verifyPoints.Add(summaryTC = "Verify number of rows (Fund AUM) are loaded = " + jsFile.Count + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // JsonFile (csv file) - Rename Keys (field name) and change Date format in JObject
                DatabaseConnection.ChangeDateFormatInJObject(jsFile, "Year", "yyyy-MM-dd HH:mm:ss.fff");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Year", "date");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Val", "aum");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Cur", "currency");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Share Cl ID", "share_class_id");
                DatabaseConnection.ChangeToIntegerTypeForValueInJObject(jsFile, "aum");
                DatabaseConnection.AddFieldNameInJObject(jsFile, "fund_id", fundid);
                DatabaseConnection.AddFieldNameInJObject(jsFile, "fund_source", dataSource);
                var testFile = DatabaseConnection.SortPropertiesByName(jsFile);

                // JsonDB (Database) - Rename Keys (field name) and change Date format in JObject
                DatabaseConnection.RemoveFieldNameInJObject(jsonDB, "_time_");
                DatabaseConnection.RemoveFieldNameInJObject(jsonDB, "create_at");
                DatabaseConnection.RemoveFieldNameInJObject(jsonDB, "import_log");
                var testDB = DatabaseConnection.SortPropertiesByName(jsonDB);

                // Sort value in JObject
                var fileSort = testFile.OrderBy(p => (string?)p["aum"]);
                var dbSort = testDB.OrderBy(p => (string?)p["aum"]);

                // Verify data from csv file was loaded into DB correctly
                verifyPoint = UploadFileAction.Instance.CompareObjects(fileSort, dbSort);
                verifyPoints.Add(summaryTC = "Verify data (Fund AUM) from csv file was loaded into DB correctly", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                #endregion

                #region Verify all informations from Data Status
                // Click on Kamehameha icon to reload/refersh Workbench
                LoginAction.Instance.ClickKamehamehaLogo(10);

                // Wait for Data Status tab is disappeared
                LoginAction.Instance.WaitForElementInvisible(10, FundDashboardPage.dataStatusTab);

                // Search a Manual Fund (QA Test 04)
                FundDashboardAction.Instance.InputNameToSearchFund(10, inputSearch)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(managerNameManual))
                                            .ClickFundNameReturnOfResults(10, managerNameManual)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(fundNameManual))
                                            .ClickFundNameReturnOfResults(10, fundNameManual)
                                            .WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);

                // Click on Data Status tab
                FundDashboardAction.Instance.ClickDataStatusTab(10);

                // verify all informations from Data Status - Fund AUM
                verifyPoint = FundDashboardAction.Instance.SectionTitleContentGetText(10, fundAUM, rowCounts) == fundAUMRowCountsVal;
                verifyPoints.Add(summaryTC = "Verify Data Status - " + fundAUM + " (Manual) " + rowCounts + " is shown correctly: " + fundAUMRowCountsVal + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = FundDashboardAction.Instance.SectionTitleContentGetText(10, fundAUM, startDate) == fundAUMStartDateVal;
                verifyPoints.Add(summaryTC = "Verify Data Status - " + fundAUM + " (Manual) " + startDate + " is shown correctly: " + fundAUMStartDateVal + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = FundDashboardAction.Instance.SectionTitleContentGetText(10, fundAUM, endDate) == fundAUMEndDateVal;
                verifyPoints.Add(summaryTC = "Verify Data Status - " + fundAUM + " (Manual) " + endDate + " is shown correctly: " + fundAUMEndDateVal + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
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

        [Test, Category("Regression Testing")]
        public void TC006_UploadFile_ManualFund_CSV_Valid_file_FundReturns()
        {
            // Notes: Please turn on your VPN before running this test cases

            #region Variables declare
            string urlInstance = LoginPage.url;
            string? inputSearch = null;
            string? managerNameManual = null;
            string? fundNameManual = null;
            int? fundid = null;
            string fundReturns = FundDashboardPage.fundReturns;
            string rowCounts = FundDashboardPage.rowCounts;
            string startDate = FundDashboardPage.startDate;
            string endDate = FundDashboardPage.endDate;
            const string fileType = "File Type";
            string fileName = "Upload_FundReturns_2rows.csv";
            string filePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Tests\Documents\CSV files\");
            string timestamp = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff");
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenDKS-UploadFileTest - ManualFund - (Fund Returns) Valid CSV File - TC006");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Log into the application
                LoginAction.Instance.LoginSiteNoGodaddy();

                // Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
                string? fundReturnsRowCountsVal = null;
                string? fundReturnsStartDateVal = null;
                string? fundReturnsEndDateVal = null;
                if (urlInstance.Contains("sandbox"))
                {
                    inputSearch = "qa test 04";
                    managerNameManual = "QA Test 04";
                    fundNameManual = "Main Fund of QA Test 04 update 01";
                    fundid = 19;
                    fundReturnsRowCountsVal = "2";
                    fundReturnsStartDateVal = "07/31/2009";
                    fundReturnsEndDateVal = "12/31/2021";
                }
                if (urlInstance.Contains("conceptia"))
                {
                    inputSearch = "qa test 05";
                    managerNameManual = "QA Test 05";
                    fundNameManual = "Main QA Test 05";
                    fundid = 12;
                    fundReturnsRowCountsVal = "2";
                    fundReturnsStartDateVal = "07/31/2009";
                    fundReturnsEndDateVal = "12/31/2021";
                }

                // Search a Manual Fund (QA Test 04)
                FundDashboardAction.Instance.InputNameToSearchFund(10, inputSearch)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(managerNameManual))
                                            .ClickFundNameReturnOfResults(10, managerNameManual)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(fundNameManual))
                                            .ClickFundNameReturnOfResults(10, fundNameManual)
                                            .WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);

                // Verify the fund name (manual fund) is shown correctly after searching
                verifyPoint = fundNameManual == FundDashboardAction.Instance.FundNameTitleGetText(10);
                verifyPoints.Add(summaryTC = "Verify the fund name (Manual Fund) is shown correctly after searching: (" + fundNameManual + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                #region Upload CSV File (Fund Returns)
                // Click on Data Status tab
                FundDashboardAction.Instance.ClickDataStatusTab(10);

                // Click on Upload File button
                UploadFileAction.Instance.ClickUploadFileButton(10);

                // Input data for File Type (dropdown)
                UploadFileAction.Instance.ClickAndSelectItemInDropdown(10, fileType, "Fund Returns");

                // Upload file (csv file - Fund Returns)
                UploadFileAction.Instance.UploadFileInput(filePath + fileName);

                // Fields Mapping - Select fields for 'Source' and 'Destination'
                UploadFileAction.Instance.ClickDestinationDropdown(10, 1)
                                         .WaitForAllElementsVisible(10, UploadFilePage.overlayDropdown)
                                         .PressDownArrowKeyUntilElementIsVisible(20);
                
                // Select item for 'Destination' dropdown
                UploadFileAction.Instance.WaitForAllElementsVisible(10, UploadFilePage.destinationDropdownItemSelect(1, UploadFilePage.date))
                                         .SelectItemInDestinationDropdown(10, 1, UploadFilePage.date)
                                         .ClickAndSelectItemInDestinationDropdown(10, 2, UploadFilePage.net)
                                         .ClickAndSelectItemInDestinationDropdown(10, 3, UploadFilePage.gross)
                                         .ClickAndSelectItemInDestinationDropdown(10, 4, UploadFilePage.currency)
                                         .ClickAndSelectItemInDestinationDropdown(10, 5, UploadFilePage.share_class_id);

                // Click Replace button
                UploadFileAction.Instance.ClickUploadLabelButton(10, "Replace") //.ClickDoneButton();
                                         .WaitForAllElementsVisible(10, UploadFilePage.dialogWarning)
                                         .ClickButtonInDialog(10, "Replace");

                // Verify a green toast message is shown after uploading a valid file  (File Type: Fund Returns)
                verifyPoint = "Success" == UploadFileAction.Instance.toastMessageAlertGetText(10, 1)
                            && "File uploaded successfully." == UploadFileAction.Instance.toastMessageAlertGetText(10, 2);
                verifyPoints.Add(summaryTC = "Verify a green toast message is shown after uploading a valid file (File Type: Fund Returns): (Success File uploaded successfully.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Wait for toast message is disappeared
                UploadFileAction.Instance.WaitForElementInvisible(10, UploadFilePage.toastMessage(1));
                #endregion

                #region Verify the database value in the table after uploading CSV file into DB
                const string db = "ks_model";
                const string table = "manual_fund_ror";
                const string dataSource = "manual";
                string query = "SELECT * FROM " + db + "." + table + " WHERE fund_id=" + fundid + " AND _time_ >= timestamp '" + timestamp + "'";

                // Wait until all files are loaded
                System.Threading.Thread.Sleep(10000);

                // Send API Requests and Get API Responses (to retrieve data)
                var sendResponse = DatabaseConnection.GetDatalakeANSISQLQueries(db, query);
                List<JObject> jsonDB = JsonConvert.DeserializeObject<List<JObject>>(sendResponse.Content);

                // Convert csv (file) to json object (--> csv from file PATH)
                var convert = DatabaseConnection.ConvertCsvFileToJsonObject(filePath + fileName);
                List<JObject> jsFile = JsonConvert.DeserializeObject<List<JObject>>(convert);

                // Verify number of rows are loaded
                verifyPoint = jsonDB.Count == jsFile.Count;
                verifyPoints.Add(summaryTC = "Verify number of rows (Fund Returns) are loaded = " + jsFile.Count + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // JsonFile (csv file) - Rename Keys (field name) and change Date format in JObject
                DatabaseConnection.ChangeDateFormatInJObject(jsFile, "Year", "yyyy-MM-dd HH:mm:ss.fff");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Year", "date");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "ValNet", "net");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "ValGross", "gross");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Cur", "currency");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Share Cl ID", "share_class_id");
                DatabaseConnection.ChangeToDoubleTypeForValueInJObject(jsFile, "net");
                DatabaseConnection.ChangeToDoubleTypeForValueInJObject(jsFile, "gross");
                DatabaseConnection.AddFieldNameInJObject(jsFile, "fund_id", fundid);
                DatabaseConnection.AddFieldNameInJObject(jsFile, "fund_source", dataSource);
                DatabaseConnection.AddFieldNameInJObject(jsFile, "other_amt_1", null);
                DatabaseConnection.AddFieldNameInJObject(jsFile, "other_amt_2", null);
                DatabaseConnection.AddFieldNameInJObject(jsFile, "other_amt_3", null);
                var testFile = DatabaseConnection.SortPropertiesByName(jsFile);

                // JsonDB (Database) - Rename Keys (field name) and change Date format in JObject
                DatabaseConnection.RemoveFieldNameInJObject(jsonDB, "_time_");
                DatabaseConnection.RemoveFieldNameInJObject(jsonDB, "create_at");
                DatabaseConnection.RemoveFieldNameInJObject(jsonDB, "import_log");
                var testDB = DatabaseConnection.SortPropertiesByName(jsonDB);

                // Sort value in JObject
                var fileSort = testFile.OrderBy(p => (string?)p["net"]);
                var dbSort = testDB.OrderBy(p => (string?)p["net"]);

                // Verify data from csv file was loaded into DB correctly
                verifyPoint = UploadFileAction.Instance.CompareObjects(fileSort, dbSort);
                verifyPoints.Add(summaryTC = "Verify data (Fund Returns) from csv file was loaded into DB correctly", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                #endregion

                #region Verify all informations from Data Status
                // Click on Kamehameha icon to reload/refersh Workbench
                LoginAction.Instance.ClickKamehamehaLogo(10);

                // Wait for Data Status tab is disappeared
                LoginAction.Instance.WaitForElementInvisible(10, FundDashboardPage.dataStatusTab);

                // Search a Manual Fund (QA Test 04)
                FundDashboardAction.Instance.InputNameToSearchFund(10, inputSearch)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(managerNameManual))
                                            .ClickFundNameReturnOfResults(10, managerNameManual)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(fundNameManual))
                                            .ClickFundNameReturnOfResults(10, fundNameManual)
                                            .WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);

                // Click on Data Status tab
                FundDashboardAction.Instance.ClickDataStatusTab(10);

                // verify all informations from Data Status - Fund Returns
                verifyPoint = FundDashboardAction.Instance.SectionTitleContentGetText(10, fundReturns, rowCounts) == fundReturnsRowCountsVal;
                verifyPoints.Add(summaryTC = "Verify Data Status - " + FundDashboardPage.fundReturns + " (Manual) " + FundDashboardPage.rowCounts + " is shown correctly: " + fundReturnsRowCountsVal + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = FundDashboardAction.Instance.SectionTitleContentGetText(10, fundReturns, startDate) == fundReturnsStartDateVal;
                verifyPoints.Add(summaryTC = "Verify Data Status - " + FundDashboardPage.fundReturns + " (Manual) " + FundDashboardPage.startDate + " is shown correctly: " + fundReturnsStartDateVal + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = FundDashboardAction.Instance.SectionTitleContentGetText(10, fundReturns, endDate) == fundReturnsEndDateVal;
                verifyPoints.Add(summaryTC = "Verify Data Status - " + FundDashboardPage.fundReturns + " (Manual) " + FundDashboardPage.endDate + " is shown correctly: " + fundReturnsEndDateVal + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
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

        [Test, Category("Regression Testing")]
        public void TC007_UploadFile_Cambridge_Valid_file_DealInformation()
        {
            // Notes: Please turn on your VPN before running this test cases

            #region Variables declare
            string urlInstance = LoginPage.url;
            string? inputSearch = null;
            string? cambridgeFund = null;
            string? managerName = null;
            string? asOfDate = null;
            int? id = null;
            const string sourceIcon = "M";
            const string fileType = "File Type";
            const string currency = "USD";
            string fileName = "deal_information.csv";
            string filePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Tests\Documents\CSV files\");
            string timestamp = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff");
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenDKS-UploadFileTest -Cambridge (Deal Information) Valid CSV File - TC007");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Log into the application
                LoginAction.Instance.LoginSiteNoGodaddy();

                // Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
                if (urlInstance.Contains("sandbox")) 
                {
                    inputSearch = "priman f0";
                    cambridgeFund = "PriMan F03";
                    managerName = "PriMan F03";
                    asOfDate = "09-30-2022";
                    id = 11; // (manager id)
                }
                if (urlInstance.Contains("conceptia")) 
                {
                    inputSearch = "priman";
                    cambridgeFund = "Firm 01 of PriMan FunSta 01";
                    managerName = "PriMan FunSta 01";
                    asOfDate = "09-30-2022";
                    id = 1; // (manager id)
                }

                // Search a Cambridge Fund (KS-455 the Firm should be searched in the search bar)
                FundDashboardAction.Instance.InputNameToSearchFund(10, inputSearch)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResultsWithItemSource(cambridgeFund, sourceIcon))
                                            .ClickFundNameReturnOfResults(10, cambridgeFund, sourceIcon); Thread.Sleep(1000);

                // Check if loading icon shows then wait for it to disappear
                if (FundDashboardAction.Instance.IsElementPresent(FundDashboardPage.loadingIcon))
                {
                    FundDashboardAction.Instance.WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);
                }

                // Verify all informations from Manager Info (tab)
                Thread.Sleep(1000);
                verifyPoint = cambridgeFund == FundDashboardAction.Instance.FundNameTitleGetText(10);
                verifyPoints.Add(summaryTC = "Verify Manager Info - firm (Cambridge Fund - Manual) is shown correctly after searching: '" + cambridgeFund + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                #region Upload CSV File (Deal Information)
                // Click on Data Status tab
                FundDashboardAction.Instance.ClickDataStatusTab(10);

                // Click on Upload File button
                UploadFileAction.Instance.ClickUploadFileButton(10);

                // Input data for File Type (dropdown)
                UploadFileAction.Instance.ClickAndSelectItemInDropdown(10, fileType, "Deal Information")
                                         .UploadFileInput(filePath + fileName) // Upload file (csv file - Deal Information)
                                         .ClickAndSelectItemInDropdown(10, "Currency", currency)
                                         .InputAsOfDate(10, asOfDate);

                // Click 'Destination' dropdown at row (x)
                UploadFileAction.Instance.ClickDestinationDropdown(10, 1)
                                         .WaitForAllElementsVisible(10, UploadFilePage.overlayDropdown)
                                         .PressDownArrowKeyUntilElementIsVisible(20);
                
                // Select item for 'Destination' dropdown
                UploadFileAction.Instance.WaitForAllElementsVisible(10, UploadFilePage.destinationDropdownItemSelect(1, UploadFilePage.company_name))
                                         .SelectItemInDestinationDropdown(10, 1, UploadFilePage.company_name)
                                         .ClickAndSelectItemInDestinationDropdown(10, 2, UploadFilePage.fund)
                                         .ClickAndSelectItemInDestinationDropdown(10, 3, UploadFilePage.entry_date)
                                         .ClickAndSelectItemInDestinationDropdown(10, 4, UploadFilePage.exit_date)
                                         .ClickAndSelectItemInDestinationDropdown(10, 5, UploadFilePage.status)
                                         .ClickAndSelectItemInDestinationDropdown(10, 6, UploadFilePage.gross_irr)
                                         .ClickAndSelectItemInDestinationDropdown(10, 7, UploadFilePage.invested_capital)
                                         .ClickAndSelectItemInDestinationDropdown(10, 8, UploadFilePage.realized_capital)
                                         .ClickAndSelectItemInDestinationDropdown(10, 9, UploadFilePage.unrealized_fmv)
                                         .ClickAndSelectItemInDestinationDropdown(10, 10, UploadFilePage.attribution_category_1)
                                         .ClickAndSelectItemInDestinationDropdown(10, 11, UploadFilePage.attribution_category_2)
                                         .ClickAndSelectItemInDestinationDropdown(10, 12, UploadFilePage.attribution_category_3)
                                         .ClickAndSelectItemInDestinationDropdown(10, 13, UploadFilePage.custom_weight);
                //.ClickAndSelectItemInDestinationDropdown(10, 14, "manager") // --> No map
                //.ClickAndSelectItemInDestinationDropdown(10, 15, "currency") // --> No map
                //.ClickAndSelectItemInDestinationDropdown(10, 16, "data_as_of_date") // --> No map

                // Click Done button
                UploadFileAction.Instance.ClickDoneButton();

                // Verify a green toast message is shown after uploading a valid file
                verifyPoint = "Success" == UploadFileAction.Instance.toastMessageAlertGetText(10, 1)
                            && "File uploaded successfully." == UploadFileAction.Instance.toastMessageAlertGetText(10, 2);
                verifyPoints.Add(summaryTC = "Verify a green toast message is shown after uploading a valid file (File Type: Deal Information): (Success File uploaded successfully.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Wait for toast message is disappeared
                UploadFileAction.Instance.WaitForElementInvisible(10, UploadFilePage.toastMessage(1));
                #endregion

                #region Verify the database value in the table after uploading CSV file into DB
                const string db = "ks_model";
                const string table = "private_deal_information";
                DateTime convert_as_of_date = DateTime.ParseExact(asOfDate, "MM-dd-yyyy", null);
                string data_as_of_date = convert_as_of_date.ToString("yyyy-MM-dd HH:mm:ss.fff");
                string query = "SELECT * FROM " + db + "." + table + " WHERE manager_id='" + id + "' AND _time_>= timestamp '" + timestamp + "'";

                // Wait until all files are loaded
                System.Threading.Thread.Sleep(23000);

                // Send API Requests and Get API Responses (to retrieve data)
                var sendResponse = DatabaseConnection.GetDatalakeANSISQLQueries(db, query);
                List<JObject> jsonDB = JsonConvert.DeserializeObject<List<JObject>>(sendResponse.Content);

                // Convert csv (file) to json object (--> csv from file PATH)
                var convert = DatabaseConnection.ConvertCsvFileToJsonObject(filePath + fileName);
                List<JObject> jsFile = JsonConvert.DeserializeObject<List<JObject>>(convert);

                // Verify number of rows are loaded
                verifyPoint = jsonDB.Count == jsFile.Count;
                verifyPoints.Add(summaryTC = "Verify number of rows (Deal Information) are loaded = " + jsFile.Count + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // JsonFile (csv file) - Rename Keys (field name) and change Date format in JObject
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Attribution Category 1", "attribution_category_1");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Attribution Category 2", "attribution_category_2");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Attribution Category 3", "attribution_category_3");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Company Name", "company_name");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Currency", "currency");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Custom Weight", "custom_weight");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Data as of Date", "data_as_of_date");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Entry Date", "entry_date");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Exit Date", "exit_date");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Fund", "fund");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Gross IRR", "gross_irr");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Invested Capital", "invested_capital");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Manager", "manager");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Realized Capital", "realized_capital");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Status", "status");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Unrealized FMV", "unrealized_fmv");
                DatabaseConnection.RemoveWhiteSpaceBeginAndEndValuesInJObject(jsFile, "attribution_category_1");
                DatabaseConnection.RemoveWhiteSpaceBeginAndEndValuesInJObject(jsFile, "company_name");
                DatabaseConnection.ReplaceTextValueInJObject(jsFile, "custom_weight", "", null);
                DatabaseConnection.ReplaceTextValueInJObject(jsFile, "data_as_of_date", "6/1/2020", data_as_of_date);
                DatabaseConnection.ChangeDateFormatInJObject(jsFile, "data_as_of_date", "yyyy-MM-dd HH:mm:ss.fff");
                DatabaseConnection.ChangeDateFormatWithNullInJObject(jsFile, "entry_date", "yyyy-MM-dd HH:mm:ss.fff", null, null);
                DatabaseConnection.ChangeDateFormatWithNullInJObject(jsFile, "exit_date", "yyyy-MM-dd HH:mm:ss.fff", null, null);
                DatabaseConnection.ChangeValuesInJObject(jsFile, "manager", managerName); // cambridgeFund
                DatabaseConnection.ReplaceContainTextValueInJObject(jsFile, "gross_irr", "%", ""); // --> Remove % sign by replacing with ""
                DatabaseConnection.ChangeToFormatDoubleAndIntegerWithNullValueInJObject(jsFile, "gross_irr", "0.##", null, null);
                DatabaseConnection.ChangeToIntegerTypeForValueInJObject(jsFile, "invested_capital");
                DatabaseConnection.ChangeToFormatDoubleAndIntegerWithNullValueInJObject(jsFile, "realized_capital", "0.##", " -   ", null);
                DatabaseConnection.ChangeToFormatDoubleAndIntegerWithNullValueInJObject(jsFile, "unrealized_fmv", "0.##", " -   ", null);
                DatabaseConnection.AddFieldNameInJObject(jsFile, "manager_id", "" + id + "");
                var testFile = DatabaseConnection.SortPropertiesByName(jsFile);

                // JsonDB (Database) - Rename Keys (field name) and change Date format in JObject
                DatabaseConnection.RemoveFieldNameInJObject(jsonDB, "manager_source");
                DatabaseConnection.RemoveFieldNameInJObject(jsonDB, "_time_");
                DatabaseConnection.RemoveFieldNameInJObject(jsonDB, "import_log");
                var testDB = DatabaseConnection.SortPropertiesByName(jsonDB);

                // Sort value in JObject
                var fileSort = testFile.OrderBy(p => (string?)p["company_name"]);
                var dbSort = testDB.OrderBy(p => (string?)p["company_name"]);

                // Verify data from csv file was loaded into DB correctly
                verifyPoint = UploadFileAction.Instance.CompareJObjectsToString(fileSort, dbSort);
                verifyPoints.Add(summaryTC = "Verify data (Deal Information) from csv file was loaded into DB correctly", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                #endregion

                #region Verify all informations from Data Status
                // Click on Kamehameha icon to reload/refersh Workbench
                LoginAction.Instance.ClickKamehamehaLogo(10);

                // Wait for Data Status tab is disappeared
                LoginAction.Instance.WaitForElementInvisible(10, FundDashboardPage.dataStatusTab);

                // Search a Cambridge Fund
                FundDashboardAction.Instance.InputNameToSearchFund(10, inputSearch)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResultsWithItemSource(cambridgeFund, sourceIcon))
                                            .ClickFundNameReturnOfResults(10, cambridgeFund, sourceIcon);
                                            //.WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);

                // Click on Data Status tab
                FundDashboardAction.Instance.ClickDataStatusTab(10);

                // verify all informations from Data Status - Deal Information
                verifyPoint = FundDashboardAction.Instance.SectionTitleContentGetText(20, FundDashboardPage.dealInformation, FundDashboardPage.rowCount) == jsFile.Count.ToString();
                verifyPoints.Add(summaryTC = "Verify Data Status - " + FundDashboardPage.dealInformation + " (Manual) " + FundDashboardPage.rowCount + " is shown correctly: " + jsFile.Count + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = FundDashboardAction.Instance.SectionTitleContentGetText(10, FundDashboardPage.dealInformation, FundDashboardPage.asOfDate) == asOfDate.Replace("-", "/");
                verifyPoints.Add(summaryTC = "Verify Data Status - " + FundDashboardPage.dealInformation + " (Manual) " + FundDashboardPage.asOfDate + " is shown correctly: " + asOfDate.Replace("-", "/") + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
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

        [Test, Category("Regression Testing")]
        public void TC008_UploadFile_Cambridge_Valid_file_FundInformation()
        {
            // Notes: Please turn on your VPN before running this test cases

            #region Variables declare
            string urlInstance = LoginPage.url;
            string? inputSearch = null;
            string? cambridgeFund = null;
            string? managerName = null;
            string? asOfDate = null;
            int? id = null;
            const string sourceIcon = "M";
            const string fileType = "File Type";
            const string currency = "USD";
            string fileName = "fund_information.csv";
            string filePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Tests\Documents\CSV files\");
            string timestamp = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff");
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenDKS-UploadFileTest - Cambridge (Fund Information) Valid CSV File - TC008");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Log into the application
                LoginAction.Instance.LoginSiteNoGodaddy();

                // Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
                if (urlInstance.Contains("sandbox"))
                {
                    inputSearch = "priman f0";
                    cambridgeFund = "PriMan F03";
                    managerName = "PriMan F03";
                    asOfDate = "09-30-2022";
                    id = 11; // (manager id)
                }
                if (urlInstance.Contains("conceptia"))
                {
                    inputSearch = "priman";
                    cambridgeFund = "Firm 01 of PriMan FunSta 01";
                    managerName = "PriMan FunSta 01";
                    asOfDate = "09-30-2022";
                    id = 1; // (manager id)
                }

                // Search a Cambridge Fund (KS-455 the Firm should be searched in the search bar)
                FundDashboardAction.Instance.InputNameToSearchFund(10, inputSearch)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResultsWithItemSource(cambridgeFund, sourceIcon))
                                            .ClickFundNameReturnOfResults(10, cambridgeFund, sourceIcon); Thread.Sleep(1000);

                // Check if loading icon shows then wait for it to disappear
                if (FundDashboardAction.Instance.IsElementPresent(FundDashboardPage.loadingIcon))
                {
                    FundDashboardAction.Instance.WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);
                }

                // Verify all informations from "Manager Info" tab
                Thread.Sleep(1000);
                verifyPoint = cambridgeFund == FundDashboardAction.Instance.FundNameTitleGetText(10);
                verifyPoints.Add(summaryTC = "Verify Manager Info - firm (Cambridge Fund - Manual) is shown correctly after searching: '" + cambridgeFund + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                #region Upload CSV File (Fund Information)
                // Click on Data Status tab
                FundDashboardAction.Instance.ClickDataStatusTab(10);

                // Click on Upload File button
                UploadFileAction.Instance.ClickUploadFileButton(10);

                // Input data for File Type (dropdown)
                UploadFileAction.Instance.ClickAndSelectItemInDropdown(10, fileType, "Fund Information")
                                         .UploadFileInput(filePath + fileName) // Upload file (csv file - Deal Information)
                                         .ClickAndSelectItemInDropdown(10, "Currency", currency)
                                         .InputAsOfDate(10, asOfDate);

                // Click 'Destination' dropdown at row (x)
                UploadFileAction.Instance.ClickDestinationDropdown(10, 1)
                                         .WaitForAllElementsVisible(10, UploadFilePage.overlayDropdown)
                                         .PressDownArrowKeyUntilElementIsVisible(20);
                
                // Select item for 'Destination' dropdown
                UploadFileAction.Instance.WaitForAllElementsVisible(10, UploadFilePage.destinationDropdownItemSelect(1, UploadFilePage.fund_name))
                                         .SelectItemInDestinationDropdown(10, 1, UploadFilePage.fund_name)
                                         .ClickAndSelectItemInDestinationDropdown(10, 2, UploadFilePage.fund_size)
                                         .ClickAndSelectItemInDestinationDropdown(10, 3, UploadFilePage.vintage_year)
                                         .ClickAndSelectItemInDestinationDropdown(10, 4, UploadFilePage.invested_capital)
                                         .ClickAndSelectItemInDestinationDropdown(10, 5, UploadFilePage.realized)
                                         .ClickAndSelectItemInDestinationDropdown(10, 6, UploadFilePage.unrealized_current_nav)
                                         .ClickAndSelectItemInDestinationDropdown(10, 7, UploadFilePage.gross_irr)
                                         .ClickAndSelectItemInDestinationDropdown(10, 8, UploadFilePage.net_irr)
                                         .ClickAndSelectItemInDestinationDropdown(10, 9, UploadFilePage.gross_tvpi)
                                         .ClickAndSelectItemInDestinationDropdown(10, 10, UploadFilePage.net_tvpi);
                                         //.ClickAndSelectItemInDestinationDropdown(10, 11, "manager") // --> No map
                                         //.ClickAndSelectItemInDestinationDropdown(10, 12, "Currency") // --> No map
                                         //.ClickAndSelectItemInDestinationDropdown(10, 13, "data_as_of_date") // --> No map

                // Click Done button
                UploadFileAction.Instance.ClickDoneButton();

                // Verify a green toast message is shown after uploading a valid file
                verifyPoint = "Success" == UploadFileAction.Instance.toastMessageAlertGetText(10, 1)
                            && "File uploaded successfully." == UploadFileAction.Instance.toastMessageAlertGetText(10, 2);
                verifyPoints.Add(summaryTC = "Verify a green toast message is shown after uploading a valid file (File Type: Fund Information): (Success File uploaded successfully.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Wait for toast message is disappeared
                UploadFileAction.Instance.WaitForElementInvisible(10, UploadFilePage.toastMessage(1));
                #endregion

                #region Verify the database value in the table after uploading CSV file into DB
                const string db = "ks_model";
                const string table = "private_fund_information";
                DateTime convert_as_of_date = DateTime.ParseExact(asOfDate, "MM-dd-yyyy", null);
                string data_as_of_date = convert_as_of_date.ToString("yyyy-MM-dd HH:mm:ss.fff");
                string query = "SELECT * FROM " + db + "." + table + " WHERE manager_id='" + id + "' AND _time_>= timestamp '" + timestamp + "'";

                // Wait until all files are loaded
                System.Threading.Thread.Sleep(10000);

                // Send API Requests and Get API Responses (to retrieve data)
                var sendResponse = DatabaseConnection.GetDatalakeANSISQLQueries(db, query);
                List<JObject> jsonDB = JsonConvert.DeserializeObject<List<JObject>>(sendResponse.Content);

                // Convert csv (file) to json object (--> csv from file PATH)
                var convert = DatabaseConnection.ConvertCsvFileToJsonObject(filePath + fileName);
                List<JObject> jsFile = JsonConvert.DeserializeObject<List<JObject>>(convert);

                // Verify number of rows are loaded into DB
                verifyPoint = jsonDB.Count == jsFile.Count;
                verifyPoints.Add(summaryTC = "Verify number of rows (Fund Information) are loaded = " + jsFile.Count + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // JsonFile (csv file) - Rename Keys (field name) and change Date format in JObject
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Fund Name", UploadFilePage.fund_name);
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Fund Size", UploadFilePage.fund_size);
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Vintage Year", UploadFilePage.vintage_year);
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Invested Capital", UploadFilePage.invested_capital);
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Realized", UploadFilePage.realized);
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Unrealized (Current NAV)", UploadFilePage.unrealized_current_nav);
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Gross IRR", UploadFilePage.gross_irr);
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Net IRR", UploadFilePage.net_irr);
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Gross TVPI", UploadFilePage.gross_tvpi);
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Net TVPI", UploadFilePage.net_tvpi);
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Manager", "manager");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Currency", UploadFilePage.currency);
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Data as of Date", UploadFilePage.data_as_of_date);
                DatabaseConnection.ChangeToFormatDoubleAndIntegerWithNullValueInJObject(jsFile, UploadFilePage.fund_size, "0.##", null, null);
                DatabaseConnection.ChangeToFormatDoubleAndIntegerWithNullValueInJObject(jsFile, UploadFilePage.vintage_year, "0.##", null, null);
                DatabaseConnection.ChangeToFormatDoubleAndIntegerWithNullValueInJObject(jsFile, UploadFilePage.invested_capital, "0.##", null, null);
                DatabaseConnection.ChangeToFormatDoubleAndIntegerWithNullValueInJObject(jsFile, UploadFilePage.realized, "0.##", null, null);
                DatabaseConnection.ChangeToFormatDoubleAndIntegerWithNullValueInJObject(jsFile, UploadFilePage.unrealized_current_nav, "0.##", null, null);
                DatabaseConnection.ReplaceContainTextValueInJObject(jsFile, UploadFilePage.gross_irr, "%", ""); // --> Remove % sign by replacing with ""
                DatabaseConnection.ReplaceContainTextValueInJObject(jsFile, UploadFilePage.net_irr, "%", ""); // --> Remove % sign by replacing with ""
                DatabaseConnection.ChangeToFormatDoubleAndIntegerWithNullValueInJObject(jsFile, UploadFilePage.gross_irr, "0.##", null, null);
                DatabaseConnection.ChangeToFormatDoubleAndIntegerWithNullValueInJObject(jsFile, UploadFilePage.net_irr, "0.##", null, null);
                DatabaseConnection.ChangeToFormatDoubleAndIntegerWithNullValueInJObject(jsFile, UploadFilePage.gross_tvpi, "0.##", null, null);
                DatabaseConnection.ChangeToFormatDoubleAndIntegerWithNullValueInJObject(jsFile, UploadFilePage.net_tvpi, "0.##", null, null);
                DatabaseConnection.ChangeValuesInJObject(jsFile, "manager", managerName); // cambridgeFund
                DatabaseConnection.ReplaceTextValueInJObject(jsFile, UploadFilePage.currency, "USD", currency);
                DatabaseConnection.ReplaceTextValueInJObject(jsFile, UploadFilePage.data_as_of_date, "6/1/2020", data_as_of_date);
                DatabaseConnection.AddFieldNameInJObject(jsFile, "manager_id", "" + id + "");
                var testFile = DatabaseConnection.SortPropertiesByName(jsFile);

                // JsonDB (Database) - Rename Keys (field name) and change Date format in JObject
                DatabaseConnection.RemoveFieldNameInJObject(jsonDB, "_time_");
                DatabaseConnection.RemoveFieldNameInJObject(jsonDB, "import_log");
                DatabaseConnection.RemoveFieldNameInJObject(jsonDB, "dpi");
                DatabaseConnection.RemoveFieldNameInJObject(jsonDB, "effective_date");
                var testDB = DatabaseConnection.SortPropertiesByName(jsonDB);

                // Sort value in JObject
                DatabaseConnection.RemoveFieldNameInJObject(jsonDB, "manager_source");
                var fileSort = testFile.OrderBy(p => (string?)p["fund_name"]);
                var dbSort = testDB.OrderBy(p => (string?)p["fund_name"]);

                // Verify data from csv file was loaded into DB correctly
                verifyPoint = UploadFileAction.Instance.CompareJObjectsToString(fileSort, dbSort);
                verifyPoints.Add(summaryTC = "Verify data (Fund Information) from csv file was loaded into DB correctly", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                #endregion

                #region Verify all informations from Data Status
                // Click on Kamehameha icon to reload/refersh Workbench
                LoginAction.Instance.ClickKamehamehaLogo(10);

                // Wait for Data Status tab is disappeared
                LoginAction.Instance.WaitForElementInvisible(10, FundDashboardPage.dataStatusTab);

                // Search a Cambridge Fund
                FundDashboardAction.Instance.InputNameToSearchFund(10, inputSearch)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResultsWithItemSource(cambridgeFund, sourceIcon))
                                            .ClickFundNameReturnOfResults(10, cambridgeFund, sourceIcon);
                                            //.WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);

                // Click on Data Status tab
                FundDashboardAction.Instance.ClickDataStatusTab(10);

                // verify all informations from Data Status - Fund Information
                verifyPoint = FundDashboardAction.Instance.SectionTitleContentGetText(20, FundDashboardPage.fundInformation, FundDashboardPage.rowCount) == jsFile.Count.ToString();
                verifyPoints.Add(summaryTC = "Verify Data Status - " + FundDashboardPage.fundInformation + " (Manual) " + FundDashboardPage.rowCount + " is shown correctly: " + jsFile.Count + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = FundDashboardAction.Instance.SectionTitleContentGetText(10, FundDashboardPage.fundInformation, FundDashboardPage.asOfDate) == asOfDate.Replace("-", "/");
                verifyPoints.Add(summaryTC = "Verify Data Status - " + FundDashboardPage.fundInformation + " (Manual) " + FundDashboardPage.asOfDate + " is shown correctly: " + asOfDate.Replace("-", "/") + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
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

        [Test, Category("Regression Testing")]
        public void TC009_UploadFile_Cambridge_Valid_file_CashFlowInformation()
        {
            // Notes: Please turn on your VPN before running this test cases

            #region Variables declare
            string urlInstance = LoginPage.url;
            string? inputSearch = null;
            string? cambridgeFund = null;
            string? managerName = null;
            string? asOfDate = null;
            int? id = null;
            const string sourceIcon = "M";
            const string fileType = "File Type";
            const string currency = "USD";
            string fileName = "cashflow_information.csv";
            string filePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Tests\Documents\CSV files\");
            string timestamp = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss.fff");
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenDKS-UploadFileTest - Cambridge (Cash Flow Information) Valid CSV File - TC009");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Log into the application
                LoginAction.Instance.LoginSiteNoGodaddy();

                // Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
                if (urlInstance.Contains("sandbox"))
                {
                    inputSearch = "priman f0";
                    cambridgeFund = "PriMan F03";
                    managerName = "PriMan F03";
                    asOfDate = "09-30-2022";
                    id = 11; // (manager id)
                }
                if (urlInstance.Contains("conceptia"))
                {
                    inputSearch = "priman";
                    cambridgeFund = "Firm 01 of PriMan FunSta 01";
                    managerName = "PriMan FunSta 01";
                    asOfDate = "09-30-2022";
                    id = 1; // (manager id)
                }

                // Search a Cambridge Fund (KS-455 the Firm should be searched in the search bar)
                FundDashboardAction.Instance.InputNameToSearchFund(10, inputSearch)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResultsWithItemSource(cambridgeFund, sourceIcon))
                                            .ClickFundNameReturnOfResults(10, cambridgeFund, sourceIcon); Thread.Sleep(1000);

                // Check if loading icon shows then wait for it to disappear
                if (FundDashboardAction.Instance.IsElementPresent(FundDashboardPage.loadingIcon))
                {
                    FundDashboardAction.Instance.WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);
                }

                // Verify all informations from "Manager Info" tab'
                Thread.Sleep(1000);
                verifyPoint = cambridgeFund == FundDashboardAction.Instance.FundNameTitleGetText(10);
                verifyPoints.Add(summaryTC = "Verify Manager Info - firm (Cambridge Fund - Manual) is shown correctly after searching: '" + cambridgeFund + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                #region Upload CSV File (Fund Information)
                // Click on Data Status tab
                FundDashboardAction.Instance.ClickDataStatusTab(10);

                // Click on Upload File button
                UploadFileAction.Instance.ClickUploadFileButton(10);

                // Input data for File Type (dropdown)
                UploadFileAction.Instance.ClickAndSelectItemInDropdown(10, fileType, "Cash Flow Information")
                                         .UploadFileInput(filePath + fileName) // Upload file (csv file - Deal Information)
                                         .ClickAndSelectItemInDropdown(10, "Currency", currency)
                                         .InputAsOfDate(10, asOfDate);

                // Click 'Destination' dropdown at row (x)
                UploadFileAction.Instance.ClickDestinationDropdown(10, 1)
                                         .WaitForAllElementsVisible(10, UploadFilePage.overlayDropdown)
                                         .PressDownArrowKeyUntilElementIsVisible(20);
                
                // Select item for 'Destination' dropdown
                UploadFileAction.Instance.WaitForAllElementsVisible(10, UploadFilePage.destinationDropdownItemSelect(1, UploadFilePage.date))
                                         .SelectItemInDestinationDropdown(10, 1, UploadFilePage.date)
                                         .ClickAndSelectItemInDestinationDropdown(10, 2, UploadFilePage.fund)
                                         .ClickAndSelectItemInDestinationDropdown(10, 3, UploadFilePage.contribution)
                                         .ClickAndSelectItemInDestinationDropdown(10, 4, UploadFilePage.distribution);
                                         //.ClickAndSelectItemInDestinationDropdown(10, 5, "manager") // --> No map
                                         //.ClickAndSelectItemInDestinationDropdown(10, 6, "currency") // --> No map
                                         //.ClickAndSelectItemInDestinationDropdown(10, 7, "data_as_of_date") // --> No map

                // Click Done button
                UploadFileAction.Instance.ClickDoneButton();

                // Verify a green toast message is shown after uploading a valid file
                verifyPoint = "Success" == UploadFileAction.Instance.toastMessageAlertGetText(10, 1)
                            && "File uploaded successfully." == UploadFileAction.Instance.toastMessageAlertGetText(10, 2);
                verifyPoints.Add(summaryTC = "Verify a green toast message is shown after uploading a valid file (File Type: Cash Flow Information): (Success File uploaded successfully.)", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Wait for toast message is disappeared
                UploadFileAction.Instance.WaitForElementInvisible(10, UploadFilePage.toastMessage(1));
                #endregion

                #region Verify the database value in the table after uploading CSV file into DB
                const string db = "ks_model";
                const string table = "private_cash_flow_information";
                DateTime convert_as_of_date = DateTime.ParseExact(asOfDate, "MM-dd-yyyy", null);
                string data_as_of_date = convert_as_of_date.ToString("yyyy-MM-dd HH:mm:ss.fff");
                string query = "SELECT * FROM " + db + "." + table + " WHERE manager_id='" + id + "' AND _time_>= timestamp '" + timestamp + "'";

                // Wait until all files are loaded
                System.Threading.Thread.Sleep(10000);

                // Send API Requests and Get API Responses (to retrieve data)
                var sendResponse = DatabaseConnection.GetDatalakeANSISQLQueries(db, query);
                List<JObject> jsonDB = JsonConvert.DeserializeObject<List<JObject>>(sendResponse.Content); // JObject

                // Convert csv (file) to json object (--> csv from file PATH)
                var convert = DatabaseConnection.ConvertCsvFileToJsonObjectContainsQuoted(filePath + fileName);
                List<JObject> jsFile = JsonConvert.DeserializeObject<List<JObject>>(convert);

                // Verify number of rows are loaded into DB
                verifyPoint = jsonDB.Count == jsFile.Count;
                verifyPoints.Add(summaryTC = "Verify number of rows (Cash Flow Information) are loaded = " + jsFile.Count + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // JsonFile (csv file) - Rename Keys (field name) and change Date format in JObject
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Date", UploadFilePage.date);
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Fund", UploadFilePage.fund);
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Contribution", UploadFilePage.contribution);
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Distribution", UploadFilePage.distribution);
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Manager", "manager");
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Currency", UploadFilePage.currency);
                DatabaseConnection.ChangeFieldNameInJObject(jsFile, "Data as of Date", UploadFilePage.data_as_of_date);

                DatabaseConnection.ChangeDateFormatWithNullInJObject(jsFile, UploadFilePage.date, "yyyy-MM-dd HH:mm:ss.fff", null, null);
                DatabaseConnection.RemoveWhiteSpaceBeginAndEndValuesInJObject(jsFile, UploadFilePage.fund);

                DatabaseConnection.ReplaceContainTextValueInJObject(jsFile, UploadFilePage.contribution, ",", ""); // --> Remove , sign by replacing with ""
                DatabaseConnection.ReplaceContainTextValueInJObject(jsFile, UploadFilePage.contribution, @"""", ""); // --> Remove "" sign (double quotes) by replacing with ""
                DatabaseConnection.ChangeToFormatDoubleAndIntegerWithNullValueInJObject(jsFile, UploadFilePage.contribution, "0.##", " -   ", null);
                DatabaseConnection.ReplaceContainTextValueInJObject(jsFile, UploadFilePage.distribution, ",", ""); // --> Remove , sign by replacing with ""
                DatabaseConnection.ReplaceContainTextValueInJObject(jsFile, UploadFilePage.distribution, @"""", ""); // --> Remove "" sign (double quotes) by replacing with ""
                DatabaseConnection.ChangeToFormatDoubleAndIntegerWithNullValueInJObject(jsFile, UploadFilePage.distribution, "0.##", " -   ", null);

                DatabaseConnection.ChangeValuesInJObject(jsFile, "manager", managerName); // cambridgeFund
                DatabaseConnection.ChangeValuesInJObject(jsFile, UploadFilePage.currency, currency);
                DatabaseConnection.ChangeValuesInJObject(jsFile, UploadFilePage.data_as_of_date, data_as_of_date);
                DatabaseConnection.AddFieldNameInJObject(jsFile, "manager_id", "" + id + "");
                var testFile = DatabaseConnection.SortPropertiesByName(jsFile);

                // JsonDB (Database) - Rename Keys (field name) and change Date format in JObject
                DatabaseConnection.RemoveFieldNameInJObject(jsonDB, "manager_source");
                DatabaseConnection.RemoveFieldNameInJObject(jsonDB, "_time_");
                DatabaseConnection.RemoveFieldNameInJObject(jsonDB, "import_log");
                var testDB = DatabaseConnection.SortPropertiesByName(jsonDB);

                // Sort value in JObject
                var fileSort = testFile.OrderBy(p => (string?)p[UploadFilePage.date]).OrderBy(o => (string?)o.Property(UploadFilePage.contribution).Value == null).ThenBy(o => (string?)o.Property(UploadFilePage.contribution).Value);
                var dbSort = testDB.OrderBy(p => (string?)p[UploadFilePage.date]).OrderBy(o => (string?)o.Property(UploadFilePage.contribution).Value == null).ThenBy(o => (string?)o.Property(UploadFilePage.contribution).Value);
                
                // Verify data from csv file was loaded into DB correctly
                verifyPoint = UploadFileAction.Instance.CompareJObjectsToString(fileSort, dbSort); // CompareObjects CompareJObjectsToString
                verifyPoints.Add(summaryTC = "Verify data (Cash Flow Information) from csv file was loaded into DB correctly", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                #endregion

                #region Verify all informations from Data Status
                // Click on Kamehameha icon to reload/refersh Workbench
                LoginAction.Instance.ClickKamehamehaLogo(10);

                // Wait for Data Status tab is disappeared
                LoginAction.Instance.WaitForElementInvisible(10, FundDashboardPage.dataStatusTab);

                // Search a Cambridge Fund
                FundDashboardAction.Instance.InputNameToSearchFund(10, inputSearch)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResultsWithItemSource(cambridgeFund, sourceIcon))
                                            .ClickFundNameReturnOfResults(10, cambridgeFund, sourceIcon);
                                            //.WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);

                // Click on Data Status tab
                FundDashboardAction.Instance.ClickDataStatusTab(10);

                // verify all informations from Data Status - Cash Flow Information
                verifyPoint = FundDashboardAction.Instance.SectionTitleContentGetText(20, FundDashboardPage.cashFlowInformation, FundDashboardPage.rowCount) == jsFile.Count.ToString();
                verifyPoints.Add(summaryTC = "Verify Data Status - " + FundDashboardPage.cashFlowInformation + " (Manual) " + FundDashboardPage.rowCount + " is shown correctly: " + jsFile.Count + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = FundDashboardAction.Instance.SectionTitleContentGetText(10, FundDashboardPage.cashFlowInformation, FundDashboardPage.asOfDate) == asOfDate.Replace("-", "/");
                verifyPoints.Add(summaryTC = "Verify Data Status - " + FundDashboardPage.cashFlowInformation + " (Manual) " + FundDashboardPage.asOfDate + " is shown correctly: " + asOfDate.Replace("-", "/") + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
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
