using AventStack.ExtentReports;
using NUnit.Framework;
using OpenQA.Selenium;
using SeleniumGendKS.Core.BaseTestCase;
using SeleniumGendKS.Core.Selenium;
using SeleniumGendKS.Pages;
using System;
using System.Security.Policy;

namespace SeleniumGendKS.Tests.Features_Testing.Regression_Testing
{
    [TestFixture, Order(4)]
    internal class ExcelOnlineLoginTest : BaseTestCase
    {
        // Variables declare
        [Obsolete]
        readonly ExtentReports rep = ExtReportgetInstance();
        private string? summaryTC;

        [Test, Category("Regression Testing")]
        [Obsolete]
        public void TC001_ExcelOnline_Login()
        {
            #region Variables declare
            string url = LoginPage.url;
            const string notificationAddInExcelOnline = @"Add Microsoft Graph data to your document!";
            const string dialogChangeInReviewMenu = @"Discover new changes";
            string iconImageExcelOnline = LoginPage.iconImageExcelOnline;
            string usernameAddin = LoginPage.usernameAddin;
            By dropdownEditingPencilExcelOnline = By.XPath(@"//div[@id='ModeSwitcher-container']");
            By bookSavedTitleLoadDone = By.XPath(@"//i[contains(@data-icon-name,'CloudCheckmark') and @aria-hidden='true']");
            By shareMenuPopup = By.XPath(@"//div[@role='heading' and .='Use Teams to collaborate on files']");
            const string ownedByKSText = "Owned by KS";
            #endregion

            #region Workflow scenario
            /*
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Excel Online Login Test");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Login Excel Online
                LoginAction.Instance.LoginExcelOnlineSite();

                // Switch to the frame of Office Document Body (to interact with menu/button of MS Excel Online)
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(40, NavigationPage.frameOfficeExcelOnline);

                // Wait for Menu/Button of MS Excel Online load Done
                NavigationAction.Instance.WaitForElementVisible(30, menuAutomateExcelOnline)
                                         .WaitForElementVisible(30, dropdownEditingPencilExcelOnline)
                                         .WaitForElementVisible(30, buttonAnalyzeDataExcelOnline);

                // Check if Review menu show dialog "Discover new changes" then click "Don't show" button
                if (NavigationAction.Instance.IsElementPresent(NavigationPage.dialogNotificationExcelOnline(dialogChangeInReviewMenu)))
                {
                    NavigationAction.Instance.ClickDontShowButtonInReviewMenuExcelOnline(10);
                }

                // Click on Insert menu of MS Excel Online
                NavigationAction.Instance.ClickInsertMenuExcelOnline(30)
                                         .ClickOfficeAddInsButtonExcelOnline(10);

                // Switch to the frame (of Office Add In ExcelOnline popup)
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(30, NavigationPage.frameDialogAddinExcelOnline);

                // Click on Upload My Add In link
                NavigationAction.Instance.ClickUploadMyAddInLinkExcelOnline(10);

                // Upload the manifest file
                //string manifestFilePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Tests\Documents\Manifest files\");
                string userProfileDownloadPath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\";
                string fileName = @"manifest.xml";

                NavigationAction.Instance.SelectBrowserFileMyAddInExcelOnline(10, userProfileDownloadPath + fileName)
                                         .ClickUploadButtonUploadAddInExcelOnline(10);

                // Switch to the frame of Office Document Body (to interact with menu/button of MS Excel Online)
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(20, NavigationPage.frameOfficeExcelOnline);

                // Wait for dialog notification (Excel Online - Add Microsoft Graph data to your document!) is displayed and then disappeared
                NavigationAction.Instance.WaitForElementOnFormToDisappear(10, NavigationPage.dialogNotificationExcelOnline(notificationAddInExcelOnline));

                // Verify image icon of FAD Add-ins is displayed
                verifyPoint = iconImageExcelOnline == LoginAction.Instance.imageIconInExcelOnlineGetIcon(10);
                verifyPoints.Add(summaryTC = "Verify image icon of FAD Add-ins in Excel Online is displayed", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click on "FAD Add-in" button
                NavigationAction.Instance.ClickFADAddInButtonExcelOnline(10);

                // Wait for WorkBench Excel Online Load Done
                NavigationAction.Instance.WaitingElementLoadDone(10, NavigationPage.workBenchLeftPaneExcelOnline);
                System.Threading.Thread.Sleep(3000);

                // Switch to the FAD Add-in frame (to interact with Workbench for Excel Online)
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameIdFADAddInCurrentInstance(LoginPage.instanceName));

                // Check if the Login (With MS) button still being shown then click Login button & then click Allow button
                if (LoginAction.Instance.IsElementPresent(LoginPage.loginWithMSAccountBtn))
                {
                    LoginAction.Instance.ClickLogin(10)
                                        .ClickbuttonAllowDialogFADAddInEO(10);
                }

                // Wait for Logout dropdown is shown
                LoginAction.Instance.WaitForElementVisible(10, LoginPage.logOutDropdown)
                                    .WaitForElementVisible(10, LoginPage.uploadPDFFilesIcon)
                                    .WaitForElementVisible(10, LoginPage.ownedByKSCkb);
                  
                // Verify username, Search Fund, 'Owned by KS' checkbox and Upload icon is shown after logging successfully
                verifyPoint = LoginAction.Instance.AccountUserNameAddinBtnGetText(10) == usernameAddin;
                verifyPoints.Add(summaryTC = "Verify 'Username' button is shown after logging successfully", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = LoginAction.Instance.IsSearchFundTxtShown();
                verifyPoints.Add(summaryTC = "Verify search fund textbox is shown after logging successfully", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = LoginAction.Instance.CkbOwnedByKSGetText(10) == ownedByKSText;
                verifyPoints.Add(summaryTC = "Verify checkbox 'Owned By KS' is shown after logging successfully", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = LoginAction.Instance.IsUploadPDFFilesIconShown(10);
                verifyPoints.Add(summaryTC = "Verify Upload PDF icon is shown after logging successfully", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click logout
                LoginAction.Instance.ClicklogOut();

                // Verify the login button is shown (Login With Microsoft Account)
                verifyPoint = LoginAction.Instance.IsLoginWithMSAccountBtnShown(10);
                verifyPoints.Add(summaryTC = "Verify Login button (Login With Microsoft Account) is displayed after logout", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();
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
            */
            #endregion

            #region Workflow scenario (Login with an account without Godaddy)
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Excel Online Login Test");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Login Excel Online
                LoginAction.Instance.LoginExcelOnlineSiteNoGodaddy();

                // Switch to the frame of Office Document Body (to interact with menu/button of MS Excel Online)
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(40, NavigationPage.frameOfficeExcelOnline);

                // Wait for Menu/Button of MS Excel Online load Done
                NavigationAction.Instance.WaitForElementVisible(30, dropdownEditingPencilExcelOnline)
                                         .WaitForElementVisible(30, bookSavedTitleLoadDone);

                // Check if Review menu show dialog "Discover new changes" then click "Don't show" button
                if (NavigationAction.Instance.IsElementPresent(NavigationPage.dialogNotificationExcelOnline(dialogChangeInReviewMenu)))
                {
                    NavigationAction.Instance.ClickDontShowButtonInReviewMenuExcelOnline(10);
                }

                // Click on Home menu of MS Excel Online
                NavigationAction.Instance.ClickHomeMenuExcelOnline(30);
                NavigationAction.Instance.ClickAddinsButtonAndWaitPopupExcelOnline(30); //.ClickAddinsButtonExcelOnline(30);
                NavigationAction.Instance.WaitForElementVisible(30, NavigationPage.popularAddinsPopup);
                NavigationAction.Instance.ClickAdvancedButtonExcelOnline(30) //.ClickMoreAddinsButtonExcelOnline(30)

                // Wait for frame(x) to load Done and then Switch to the frame (of Office Add In ExcelOnline popup)
                .SwitchToFrameWithWaitMethod(20, NavigationPage.frameDialogAddinExcelOnline);

                //// Click on 'MY ADD INS' tab 
                //.ClickMYADDINSTabExcelOnline(30);

                // Click on Upload My Add In link
                NavigationAction.Instance.ClickUploadMyAddInLinkExcelOnline(10);

                // Upload the manifest file
                string userProfileDownloadPath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\";
                string fileName = @"manifest.xml";

                NavigationAction.Instance.SelectBrowserFileMyAddInExcelOnline(10, userProfileDownloadPath + fileName)
                                         .ClickUploadButtonUploadAddInExcelOnline(10);

                // Switch to the frame of Office Document Body (to interact with menu/button of MS Excel Online)
                Driver.Browser.SwitchTo().DefaultContent();
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(20, NavigationPage.frameOfficeExcelOnline);

                // Wait for dialog notification (Excel Online - Add Microsoft Graph data to your document!) is displayed
                NavigationAction.Instance.WaitForElementVisible(10, NavigationPage.dialogNotificationExcelOnline(notificationAddInExcelOnline));

                // Click on Close button of 'Notification' dialog
                NavigationAction.Instance.ClickNotificationCloseButtonExcelOnline(10);

                // Verify image icon of FAD Add-ins is displayed
                verifyPoint = iconImageExcelOnline == LoginAction.Instance.imageIconInExcelOnlineGetIcon(10);
                verifyPoints.Add(summaryTC = "Verify image icon of FAD Add-ins in Excel Online is displayed", verifyPoint);
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
                LoginAction.Instance.WaitForLoginButtonIsShown(url, 20);

                // Wait for Logout dropdown is shown
                LoginAction.Instance.WaitForElementVisible(10, LoginPage.logOutDropdown)
                                    .WaitForElementVisible(10, LoginPage.uploadPDFFilesIcon)
                                    .WaitForElementVisible(10, LoginPage.ownedByKSCkb);
                  
                // Verify username, Search Fund, 'Owned by KS' checkbox and Upload icon is shown after logging successfully
                verifyPoint = LoginAction.Instance.AccountUserNameAddinBtnGetText(10) == usernameAddin;
                verifyPoints.Add(summaryTC = "Verify 'Username' button is shown after logging successfully", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = LoginAction.Instance.IsSearchFundTxtShown();
                verifyPoints.Add(summaryTC = "Verify search fund textbox is shown after logging successfully", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = LoginAction.Instance.CkbOwnedByKSGetText(10) == ownedByKSText;
                verifyPoints.Add(summaryTC = "Verify checkbox 'Owned By KS' is shown after logging successfully", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = LoginAction.Instance.IsUploadPDFFilesIconShown(10);
                verifyPoints.Add(summaryTC = "Verify Upload PDF icon is shown after logging successfully", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click logout
                LoginAction.Instance.ClicklogOut();

                // Verify the login button is shown (Login With Microsoft Account)
                verifyPoint = LoginAction.Instance.IsLoginWithMSAccountBtnShown(10);
                verifyPoints.Add(summaryTC = "Verify Login button (Login With Microsoft Account) is displayed after logout", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();
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
