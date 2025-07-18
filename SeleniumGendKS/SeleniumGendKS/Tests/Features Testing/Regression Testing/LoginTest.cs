using AventStack.ExtentReports;
using NUnit.Framework;
using SeleniumGendKS.Core.BaseTestCase;
using SeleniumGendKS.Core.Selenium;
using SeleniumGendKS.Pages;
using System.Xml.Linq;
using System.Xml.XPath;

namespace SeleniumGendKS.Tests.Features_Testing.Regression_Testing
{
    [TestFixture, Order(1)]
    internal class LoginTest : BaseTestCase
    {
        // Variables declare
        private static readonly XDocument xdoc = LoginPage.xdoc;
        [Obsolete]
        readonly ExtentReports rep = ExtReportgetInstance();
        private string? summaryTC;

        [SetUp]
        public override void SetupTest()
        {
            //Data-driven for site testing
            string url = xdoc.XPathSelectElement("config/site").Value;
            verifyPoints.Clear();
            Driver.StartBrowser();
            LoginAction.Instance.NavigateSite(url);
        }

        [Test, Category("Regression Testing")]
        [Obsolete]
        public void TC001_login_with_valid_account()
        {
            #region Variables declare
            string email = LoginPage.email;
            string password = LoginPage.password;
            string username = LoginPage.username;
            const string ownedByKSText = "Owned by KS";
            #endregion

            #region Workflow scenario
            /*
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Login Test");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Store the current window handle (main window - Workbench)
                string winHandleBefore = Driver.Browser.CurrentWindowHandle;

                // Wait for Login button is shown
                LoginAction.Instance.WaitForElementVisible(10, LoginPage.loginWithMSAccountBtn);

                // Verify the login button is shown (Login With Microsoft Account)
                verifyPoint = LoginAction.Instance.IsLoginWithMSAccountBtnShown(10);
                verifyPoints.Add(summaryTC = "Verify Login button (Login With Microsoft Account) is displayed", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click Login button
                LoginAction.Instance.ClickLogin(10);

                // Wait for the new window is opened
                LoginAction.Instance.WaitForNewWindowOpen(15);

                // Store all the opened window into the 'list'
                string winHandleLast = Driver.Browser.WindowHandles.Last();
                List<string> listWindow = Driver.Browser.WindowHandles.ToList();
                string lastWindowHandle = "";
                foreach (var handle in listWindow)
                {
                    //Switch to the desired window first and then execute commands using driver
                    Driver.Browser.SwitchTo().Window(handle);
                    lastWindowHandle = handle;
                }

                // Switch to new window opened
                Driver.Browser.SwitchTo().Window(lastWindowHandle);

                // Wait for the element of new window is opened
                LoginAction.Instance.WaitForElementVisible(10, LoginPage.signInEmailInputTxt);

                // Enter Email and then click Next button
                LoginAction.Instance.EnterEmail(email).ClickNext(10);

                // Switch to new window opened
                Driver.Browser.SwitchTo().Window(winHandleLast);
                LoginAction.Instance.WaitForElementVisible(10, LoginPage.signInBtn);

                // Sign In - Enter password (godaddy - email is auto filled by Browser)
                LoginAction.Instance.EnterPasswordGodaddy(password)
                                    .ClickSignInGodaddy(10)
                                    .WaitForElementVisible(10, LoginPage.btnBackGodaddy)
                                    .ClickNext(10);

                // Check if the popup MS still being shown, then switch to them main window to click login (with MS) button again
                LoginAction.Instance.CheckIfMSLoginPopupStillShown(5);

                // Switch to the main window
                Driver.Browser.SwitchTo().Window(winHandleBefore); // or WindowHandles[0] (main windows KS Workbench)
                System.Threading.Thread.Sleep(2000);

                // Wait for buttons is shown
                LoginAction.Instance.WaitForElementVisible(10, LoginPage.logOutDropdown)
                                        .WaitForElementVisible(10, LoginPage.uploadPDFFilesIcon)
                                        .WaitForElementVisible(10, LoginPage.ownedByKSCkb);

                // Donwload Manifest file
                LoginAction.Instance.NavigateSite(LoginPage.url + @"/#/manifest");

                // Check manifest downloaded (timeout = 9s)
                string userProfileDownloadPath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + "Downloads";
                string fileName = @"manifest.xml";
                LoginAction.Instance.CheckFileDownloadIsComplete(9, userProfileDownloadPath, fileName);

                // Verify username, Search Fund, 'Owned by KS' checkbox and Upload icon is shown after logging successfully
                verifyPoint = LoginAction.Instance.AccountUserNameBtnGetText(10) == username;
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
                verifyPoint = LoginAction.Instance.IsLoginWithMSAccountBtnShown(1);
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

            #region Workflow scenario (Login with an account without Godaddy-loggin)
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Login Test");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Store the current window handle (main window - Workbench)
                string winHandleBefore = Driver.Browser.CurrentWindowHandle;

                // Wait for Login button is shown
                LoginAction.Instance.WaitForElementVisible(10, LoginPage.loginWithMSAccountBtn);

                // Verify the login button is shown (Login With Microsoft Account)
                verifyPoint = LoginAction.Instance.IsLoginWithMSAccountBtnShown(10);
                verifyPoints.Add(summaryTC = "Verify Login button (Login With Microsoft Account) is displayed", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click Login button
                LoginAction.Instance.ClickLogin(10);

                // Wait for the new window is opened
                LoginAction.Instance.WaitForNewWindowOpen(15);

                // Store all the opened window into the 'list'
                string winHandleLast = Driver.Browser.WindowHandles.Last();
                List<string> listWindow = Driver.Browser.WindowHandles.ToList();
                string lastWindowHandle = "";
                foreach (var handle in listWindow)
                {
                    //Switch to the desired window first and then execute commands using driver
                    Driver.Browser.SwitchTo().Window(handle);
                    lastWindowHandle = handle;
                }

                // Switch to new window opened
                Driver.Browser.SwitchTo().Window(lastWindowHandle);

                // Wait for the element of new window is opened
                LoginAction.Instance.WaitForElementVisible(10, LoginPage.signInEmailInputTxt);

                // Enter Email and then click Next button
                LoginAction.Instance.EnterEmail(email).ClickNext(10);

                // Switch to new window opened
                Driver.Browser.SwitchTo().Window(winHandleLast);
                System.Threading.Thread.Sleep(3000); //LoginAction.Instance.WaitForElementVisible(10, LoginPage.signInBtn);

                // Enter Password and then Click on SignIn button ...
                LoginAction.Instance.EnterPassword(10, password)
                                    .ClickSignIn(10);
                                    //.ClickNext(10); --> MS Login updated: Skip thip step 

                // Check if the popup MS still being shown, then switch to them main window to click login (with MS) button again
                LoginAction.Instance.CheckIfMSLoginPopupStillShown(20);

                // Wait for buttons is shown
                LoginAction.Instance.WaitForElementVisible(10, LoginPage.logOutDropdown)
                                        .WaitForElementVisible(10, LoginPage.uploadPDFFilesIcon)
                                        .WaitForElementVisible(10, LoginPage.ownedByKSCkb);

                // Donwload Manifest file
                LoginAction.Instance.NavigateSite(LoginPage.url + @"/#/manifest");

                // Check manifest downloaded (timeout = 9s)
                string userProfileDownloadPath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + "Downloads";
                string fileName = @"manifest.xml";
                verifyPoint = LoginAction.Instance.CheckFileDownloadIsComplete(9, userProfileDownloadPath, fileName);
                verifyPoints.Add(summaryTC = "The "+ fileName + " file is downloaded successful", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Verify username, Search Fund, 'Owned by KS' checkbox and Upload icon is shown after logging successfully
                verifyPoint = LoginAction.Instance.AccountUserNameBtnGetText(10) == username;
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
                verifyPoint = LoginAction.Instance.IsLoginWithMSAccountBtnShown(1);
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

