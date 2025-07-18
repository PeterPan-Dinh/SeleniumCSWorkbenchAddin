using AventStack.ExtentReports;
using NUnit.Framework;
using OpenQA.Selenium;
using SeleniumGendKS.Core.BaseTestCase;
using SeleniumGendKS.Core.Selenium;
using SeleniumGendKS.Pages;

namespace SeleniumGendKS.Tests.Features_Testing.Regression_Testing
{
    [TestFixture, Order(5)]
    internal class ExcelOnlineMultipleInstancesTest : BaseTestCase
    {
        // Variables declare
        [Obsolete]
        readonly ExtentReports rep = ExtReportgetInstance();
        private string? summaryTC;

        [Test, Category("Regression Testing")]
        public void TC001_ExcelOnlineInstances_Login()
        {
            #region Variables declare
            string urlInstance = LoginPage.url;
            const string notificationAddInExcelOnline = @"Add Microsoft Graph data to your document!";
            string fileName2 = @"manifest (1).xml"; // Name1 = @"manifest.xml";
            string manifestFilePath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\";
            int iconImageExcelOnlineInstance2 = 3; // (Instance1 = 1;)
            string iconImageExcelOnlineSandbox = @"https://workbench-sandbox.lab.gend.vn/assets/images/QA.png";
            string iconImageExcelOnlineStaging = @"https://workbench.conceptia.com/assets/images/Conceptia.png";
            string managerNameAlbourne = FundDashboardPage.managerNameAlbourne;
            string fundNameAlbourne = FundDashboardPage.fundNameAlbourne;
            int buttonFADAddInEOInstance2 = 1; // --> if No use acc without Godaddy then = 2
            #endregion

            #region Workflow scenario
            // Delete Manifest file
            LoginAction.Instance.DeleteAllFilesManifest();

            // check if the 1st Workbench instance is the Sandbox site, then login the Sandbox site first
            if (urlInstance.Contains("sandbox"))
            {
                #region Add the 1st FAD Add-in Instance
                SandboxInstanceLogin();
                #endregion

                #region Add the 2nd FAD Add-in Instance
                try
                {
                    // Add Workbench Excel Online for instance 2 --> By uploading manifest.xml (for the 2nd instance)
                    NavigationAction.Instance.UploadManifestExcelOnlineNoGodaddy(manifestFilePath, fileName2); //UploadManifestExcelOnline();

                    // Switch to the frame of Office Document Body (to interact with menu/button of MS Excel Online)
                    Driver.Browser.SwitchTo().DefaultContent();
                    NavigationAction.Instance.SwitchToFrameWithWaitMethod(40, NavigationPage.frameOfficeExcelOnline);

                    // Wait for dialog notification (Excel Online - Add Microsoft Graph data to your document!) is displayed
                    NavigationAction.Instance.WaitForElementVisible(10, NavigationPage.dialogNotificationExcelOnline(notificationAddInExcelOnline));

                    // Click on Close button of 'Notification' dialog
                    NavigationAction.Instance.ClickNotificationCloseButtonExcelOnline(10);

                    // Verify image icon of FAD Add-ins (instance 2) is displayed
                    verifyPoint = iconImageExcelOnlineStaging == NavigationAction.Instance.IsImageIconExcelOnlineInstanceShown(10, iconImageExcelOnlineInstance2);
                    verifyPoints.Add(summaryTC = "Verify image icon of FAD Add-ins Excel Online (instance 2) is displayed", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    // Click on the 2nd "FAD Add-in" button
                    NavigationAction.Instance.ClickFADAddInButtonInstanceExcelOnline(10, buttonFADAddInEOInstance2);

                    // Wait for WorkBench Excel Online Load Done
                    NavigationAction.Instance.WaitingElementLoadDone(10, NavigationPage.workBenchLeftPaneExcelOnline);
                    System.Threading.Thread.Sleep(3000);

                    // Switch to the FAD Add-in frame (to interact with Workbench for Excel Online)
                    NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameIdFADAddInCurrentInstance("sandbox")); // old: conceptia

                    // Check if the Login (With MS) button still being shown then click Login button & then click Allow button
                    if (LoginAction.Instance.IsElementPresent(LoginPage.loginWithMSAccountBtn))
                    {
                        LoginAction.Instance.ClickLogin(10)
                                            .ClickbuttonAllowDialogFADAddInEO(10);
                    }

                    // Wait for buttons is shown
                    LoginAction.Instance.WaitForElementVisible(15, LoginPage.logOutDropdown)
                                        .WaitForElementVisible(15, LoginPage.uploadPDFFilesIcon)
                                        .WaitForElementVisible(15, LoginPage.ownedByKSCkb);

                    // Search a Fund - Source = Albourne (at instance 1)
                    FundDashboardAction.Instance.InputNameToSearchFund(10, "cita")
                                                .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(managerNameAlbourne))
                                                .ClickFundNameReturnOfResults(10, managerNameAlbourne)
                                                .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(fundNameAlbourne))
                                                .ClickFundNameReturnOfResults(10, fundNameAlbourne)
                                                .WaitForElementVisible(10, FundDashboardPage.fundNameTitle);

                    verifyPoint = fundNameAlbourne == FundDashboardAction.Instance.FundNameTitleGetText(10);
                    verifyPoints.Add(summaryTC = "Verify the Fund Name (Albourne) (Instance 2) is shown correctly after searching: '" + fundNameAlbourne + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);
                }
                catch (Exception exception)
                {
                    // Print exception
                    System.Console.WriteLine(exception);

                    // Delete Manifest file
                    LoginAction.Instance.DeleteAllFilesManifest();

                    // Warning
                    ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                    Assert.Inconclusive("Something wrong! Please check console log.");
                }
                #endregion
            }

            // check if the 1st Workbench instance is the Staging site, then login the Staging site first
            if (urlInstance.Contains("conceptia"))
            {
                #region Add the 1st FAD Add-in Instance
                StagingInstanceLogin();
                #endregion

                #region Add the 2nd FAD Add-in Instance
                try
                {
                    // Add Workbench Excel Online for instance 2 --> By uploading manifest.xml (for the 2nd instance)
                    NavigationAction.Instance.UploadManifestExcelOnlineNoGodaddy(manifestFilePath, fileName2); // UploadManifestExcelOnline(manifestFilePath, fileName2);

                    // Switch to the frame of Office Document Body (to interact with menu/button of MS Excel Online)
                    Driver.Browser.SwitchTo().DefaultContent();
                    NavigationAction.Instance.SwitchToFrameWithWaitMethod(40, NavigationPage.frameOfficeExcelOnline);

                    // Wait for dialog notification (Excel Online - Add Microsoft Graph data to your document!) is displayed
                    NavigationAction.Instance.WaitForElementVisible(10, NavigationPage.dialogNotificationExcelOnline(notificationAddInExcelOnline));

                    // Click on Close button of 'Notification' dialog
                    NavigationAction.Instance.ClickNotificationCloseButtonExcelOnline(10);

                    // Verify image icon of FAD Add-ins (instance 2) is displayed
                    verifyPoint = iconImageExcelOnlineSandbox == NavigationAction.Instance.IsImageIconExcelOnlineInstanceShown(10, iconImageExcelOnlineInstance2);
                    verifyPoints.Add(summaryTC = "Verify image icon of FAD Add-ins Excel Online (instance 2) is displayed", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);

                    // Click on the 2nd "FAD Add-in" button
                    NavigationAction.Instance.ClickFADAddInButtonInstanceExcelOnline(10, buttonFADAddInEOInstance2);

                    // Wait for WorkBench Excel Online Load Done
                    NavigationAction.Instance.WaitingElementLoadDone(10, NavigationPage.workBenchLeftPaneExcelOnline);
                    System.Threading.Thread.Sleep(3000);
                    //System.Windows.Forms.SendKeys.SendWait(@"{DOWN}");

                    // Switch to the FAD Add-in frame (to interact with Workbench for Excel Online)
                    NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameIdFADAddInCurrentInstance("conceptia")); // old: sandbox

                    // Check if the Login (With MS) button still being shown then click Login button & then click Allow button
                    if (LoginAction.Instance.IsElementPresent(LoginPage.loginWithMSAccountBtn))
                    {
                        LoginAction.Instance.ClickLogin(10)
                                            .ClickbuttonAllowDialogFADAddInEO(10);
                    }

                    // Wait for buttons is shown
                    LoginAction.Instance.WaitForElementVisible(10, LoginPage.logOutDropdown)
                                        .WaitForElementVisible(10, LoginPage.uploadPDFFilesIcon)
                                        .WaitForElementVisible(10, LoginPage.ownedByKSCkb);

                    // Search a Fund - Source = Albourne (at instance 1)
                    FundDashboardAction.Instance.InputNameToSearchFund(10, "cita")
                                                .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(managerNameAlbourne))
                                                .ClickFundNameReturnOfResults(10, managerNameAlbourne)
                                                .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(fundNameAlbourne))
                                                .ClickFundNameReturnOfResults(10, fundNameAlbourne)
                                                .WaitForElementVisible(10, FundDashboardPage.fundNameTitle);

                    verifyPoint = fundNameAlbourne == FundDashboardAction.Instance.FundNameTitleGetText(10);
                    verifyPoints.Add(summaryTC = "Verify the Fund Name (Albourne) (Instance 2) is shown correctly after searching: '" + fundNameAlbourne + "'", verifyPoint);
                    ExtReportResult(verifyPoint, summaryTC);
                }
                catch (Exception exception)
                {
                    // Print exception
                    System.Console.WriteLine(exception);

                    // Delete Manifest file
                    LoginAction.Instance.DeleteAllFilesManifest();

                    // Warning
                    ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                    Assert.Inconclusive("Something wrong! Please check console log.");
                }
                #endregion
            }

            // Delete Manifest file
            LoginAction.Instance.DeleteAllFilesManifest();
            #endregion
        }

        internal void SandboxInstanceLogin()
        {
            #region Variables declare
            string urlInstance = LoginPage.url;
            string urlInstance2 = @"https://workbench.conceptia.com";
            const string notificationAddInExcelOnline = @"Add Microsoft Graph data to your document!";
            string fileName = @"manifest.xml";
            int iconImageExcelOnlineInstance1 = 1;
            string iconImageExcelOnlineSandbox = @"https://workbench-sandbox.lab.gend.vn/assets/images/QA.png";
            string managerNameAlbourne = FundDashboardPage.managerNameAlbourne;
            string fundNameAlbourne = FundDashboardPage.fundNameAlbourne;
            By shareMenuPopup = By.XPath(@"//div[@role='heading' and .='Use Teams to collaborate on files']");
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Excel Online Multiple Instances Test");
            try
            {
                // Login for the 1st Workbench instance
                LoginAction.Instance.LoginSiteNoGodaddy(urlInstance); // LoginSite(urlInstance);

                // Donwload Manifest file
                LoginAction.Instance.NavigateSite(urlInstance + @"/#/manifest");

                // Check manifest downloaded (timeout = 9s)
                string userProfileDownloadPath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + "Downloads";
                verifyPoint = LoginAction.Instance.CheckFileDownloadIsComplete(9, userProfileDownloadPath, fileName);
                verifyPoints.Add(summaryTC = "The " + fileName + " file is downloaded successful", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                #region Method 1 --> Close and then re-open browser
                /*// Close browser
                Driver.StopBrowser();

                // Open browser        
                Driver.StartBrowser();

                // Login for the 2nd Workbench instance
                LoginAction.Instance.LoginSiteNoGodaddy(urlInstance2); */
                #endregion

                #region Method 2 --> Open a new tab and switch to new tab
                ((IJavaScriptExecutor)Driver.Browser).ExecuteScript("window.open('" + urlInstance2 + "');");  // Open new tab
                System.Threading.Thread.Sleep(1000);
                Driver.Browser.SwitchTo().Window(Driver.Browser.WindowHandles.Last());
                System.Threading.Thread.Sleep(1000);

                // Login for the 2nd Workbench instance
                LoginAction.Instance.ClickLogin(10) //.LoginSiteNoGodaddy(urlInstance2); // LoginSite(urlInstance2);
                                    .WaitForElementVisible(10, LoginPage.logOutDropdown);
                #endregion

                // Donwload Manifest file
                LoginAction.Instance.NavigateSite(urlInstance2 + @"/#/manifest");

                // Check manifest downloaded (timeout = 9s)
                verifyPoint = LoginAction.Instance.CheckFileDownloadIsComplete(9, userProfileDownloadPath, @"manifest (1).xml");
                verifyPoints.Add(summaryTC = "The manifest (1).xml file is downloaded successful", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Go to Excel Online
                LoginAction.Instance.NavigateSite(LoginPage.urlExcelOnline);

                // Upload manifest.xml (the 1st instance)
                string manifestFilePath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\";
                NavigationAction.Instance.UploadManifestExcelOnlineNoGodaddy(manifestFilePath, fileName); // UploadManifestExcelOnline();

                // Switch to the frame of Office Document Body (to interact with menu/button of MS Excel Online)
                Driver.Browser.SwitchTo().DefaultContent();
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameOfficeExcelOnline);

                // Wait for dialog notification (Excel Online - Add Microsoft Graph data to your document!) is displayed
                NavigationAction.Instance.WaitForElementVisible(10, NavigationPage.dialogNotificationExcelOnline(notificationAddInExcelOnline));

                // Click on Close button of 'Notification' dialog
                NavigationAction.Instance.ClickNotificationCloseButtonExcelOnline(10);

                // Verify image icon of FAD Add-ins (instance 1) is displayed
                verifyPoint = iconImageExcelOnlineSandbox == NavigationAction.Instance.IsImageIconExcelOnlineInstanceShown(10, iconImageExcelOnlineInstance1);
                verifyPoints.Add(summaryTC = "Verify image icon of FAD Add-ins Excel Online (instance 1) is displayed", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click on "FAD Add-in" button
                NavigationAction.Instance.ClickFADAddInButtonExcelOnline(10);

                // Wait for WorkBench Excel Online Load Done
                NavigationAction.Instance.WaitingElementLoadDone(10, NavigationPage.workBenchLeftPaneExcelOnline);
                System.Threading.Thread.Sleep(3000);

                // Check if 'Share' menu show popup 'Use Teams to collaborate on files' then click 'Got it' button
                if (NavigationAction.Instance.IsElementPresent(shareMenuPopup))
                {
                    // CLick 'Got it' button
                    NavigationAction.Instance.MouseClick(3, NavigationPage.gotItBtnExcelOnline);
                }

                // Switch to the FAD Add-in frame (to interact with Workbench for Excel Online)
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameIdFADAddInCurrentInstance(LoginPage.instanceName));

                // Check if the Login (With MS) button still being shown then click Login button & then click Allow button
                LoginAction.Instance.WaitForLoginButtonIsShown(urlInstance, 20);

                // Wait for buttons is shown
                LoginAction.Instance.WaitForLogOutButtonIsShown(20);

                // Search a Fund - Source = Albourne (at instance 1)
                FundDashboardAction.Instance.InputNameToSearchFund(10, "cita")
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(managerNameAlbourne))
                                            .ClickFundNameReturnOfResults(10, managerNameAlbourne)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(fundNameAlbourne))
                                            .ClickFundNameReturnOfResults(10, fundNameAlbourne)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameTitle);

                verifyPoint = fundNameAlbourne == FundDashboardAction.Instance.FundNameTitleGetText(10);
                verifyPoints.Add(summaryTC = "Verify the Fund Name (Albourne) (Instance 1) is shown correctly after searching: '" + fundNameAlbourne + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
            }
            catch (Exception exception)
            {
                // Print exception
                System.Console.WriteLine(exception);

                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Warning
                ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong! Please check console log.");
            }
            #endregion
        }

        internal void StagingInstanceLogin()
        {
            #region Variables declare
            string urlInstance = LoginPage.url;
            string urlInstance2 = @"https://workbench-sandbox.lab.gend.vn";
            const string notificationAddInExcelOnline = @"Add Microsoft Graph data to your document!";
            string fileName = @"manifest.xml";
            int iconImageExcelOnlineInstance1 = 1;
            string iconImageExcelOnlineStaging = @"https://workbench.conceptia.com/assets/images/Conceptia.png";
            string managerNameAlbourne = FundDashboardPage.managerNameAlbourne;
            string fundNameAlbourne = FundDashboardPage.fundNameAlbourne;
            By shareMenuPopup = By.XPath(@"//div[@role='heading' and .='Use Teams to collaborate on files']");
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Excel Online Multiple Instances Test");
            try
            {
                // Login for the 1st Workbench instance
                LoginAction.Instance.LoginSiteNoGodaddy(urlInstance); //LoginSite(urlInstance);

                // Donwload Manifest file
                LoginAction.Instance.NavigateSite(urlInstance + @"/#/manifest");

                // Check manifest downloaded (timeout = 9s)
                string userProfileDownloadPath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + "Downloads";
                verifyPoint = LoginAction.Instance.CheckFileDownloadIsComplete(9, userProfileDownloadPath, fileName);
                verifyPoints.Add(summaryTC = "The " + fileName + " file is downloaded successful", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                #region Method 1 --> Close and then re-open browser
                /*// Close browser
                Driver.StopBrowser();

                // Open browser        
                Driver.StartBrowser();
                
                // Login for the 2nd Workbench instance
                LoginAction.Instance.LoginSiteNoGodaddy(urlInstance2); // LoginSite(urlInstance2); */
                #endregion

                #region Method 2 --> Open a new tab and switch to new tab
                ((IJavaScriptExecutor)Driver.Browser).ExecuteScript("window.open('"+ urlInstance2 + "');");  // Open new tab
                System.Threading.Thread.Sleep(1000);
                Driver.Browser.SwitchTo().Window(Driver.Browser.WindowHandles.Last());
                System.Threading.Thread.Sleep(1000);
                
                // Login for the 2nd Workbench instance
                LoginAction.Instance.ClickLogin(10) //.LoginSiteNoGodaddy(urlInstance2); // LoginSite(urlInstance2);
                                    .WaitForElementVisible(10, LoginPage.logOutDropdown);
                #endregion

                // Donwload Manifest file
                LoginAction.Instance.NavigateSite(urlInstance2 + @"/#/manifest");

                // Check manifest downloaded (timeout = 9s)
                verifyPoint = LoginAction.Instance.CheckFileDownloadIsComplete(9, userProfileDownloadPath, @"manifest (1).xml");
                verifyPoints.Add(summaryTC = "The manifest (1).xml file is downloaded successful", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Go to Excel Online
                LoginAction.Instance.NavigateSite(LoginPage.urlExcelOnline);

                // Upload manifest.xml (the 1st instance)
                string manifestFilePath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\";
                NavigationAction.Instance.UploadManifestExcelOnlineNoGodaddy(manifestFilePath, fileName); //UploadManifestExcelOnline();

                // Switch to the frame of Office Document Body (to interact with menu/button of MS Excel Online)
                Driver.Browser.SwitchTo().DefaultContent();
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameOfficeExcelOnline);

                // Wait for dialog notification (Excel Online - Add Microsoft Graph data to your document!) is displayed
                NavigationAction.Instance.WaitForElementVisible(10, NavigationPage.dialogNotificationExcelOnline(notificationAddInExcelOnline));

                // Click on Close button of 'Notification' dialog
                NavigationAction.Instance.ClickNotificationCloseButtonExcelOnline(10);

                // Verify image icon of FAD Add-ins (instance 1) is displayed
                verifyPoint = iconImageExcelOnlineStaging == NavigationAction.Instance.IsImageIconExcelOnlineInstanceShown(10, iconImageExcelOnlineInstance1);
                verifyPoints.Add(summaryTC = "Verify image icon of FAD Add-ins Excel Online (instance 1) is displayed", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click on "FAD Add-in" button
                NavigationAction.Instance.ClickFADAddInButtonExcelOnline(10);

                // Wait for WorkBench Excel Online Load Done
                NavigationAction.Instance.WaitingElementLoadDone(10, NavigationPage.workBenchLeftPaneExcelOnline);
                System.Threading.Thread.Sleep(3000); // 15000

                // Check if 'Share' menu show popup 'Use Teams to collaborate on files' then click 'Got it' button
                if (NavigationAction.Instance.IsElementPresent(shareMenuPopup))
                {
                    // CLick 'Got it' button
                    NavigationAction.Instance.MouseClick(3, NavigationPage.gotItBtnExcelOnline);
                }

                // Switch to the FAD Add-in frame (to interact with Workbench for Excel Online)
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameIdFADAddInCurrentInstance(LoginPage.instanceName));

                // Check if the Login (With MS) button still being shown then click Login button & then click Allow button
                LoginAction.Instance.WaitForLoginButtonIsShown(urlInstance, 15);

                // Wait for buttons is shown
                LoginAction.Instance.WaitForLogOutButtonIsShown(30);

                // Search a Fund - Source = Albourne (at instance 1)
                FundDashboardAction.Instance.InputNameToSearchFund(10, "cita")
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(managerNameAlbourne))
                                            .ClickFundNameReturnOfResults(10, managerNameAlbourne)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(fundNameAlbourne))
                                            .ClickFundNameReturnOfResults(10, fundNameAlbourne)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameTitle);

                verifyPoint = fundNameAlbourne == FundDashboardAction.Instance.FundNameTitleGetText(10);
                verifyPoints.Add(summaryTC = "Verify the Fund Name (Albourne) (Instance 1) is shown correctly after searching: '" + fundNameAlbourne + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
            }
            catch (Exception exception)
            {
                // Print exception
                System.Console.WriteLine(exception);

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
