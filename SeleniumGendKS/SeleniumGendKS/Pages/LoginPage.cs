using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using SeleniumGendKS.Core.Selenium;
using SeleniumGendKS.Core.BaseClass;
using System.Xml.Linq;
using System.Xml.XPath;
using NUnit.Framework;
using SeleniumGendKS.Core.BaseTestCase;
using System.Security.Policy;
using static Microsoft.AspNetCore.Razor.Language.TagHelperMetadata;

namespace SeleniumGendKS.Pages
{
    internal class LoginPage : BasePageElementMap
    {
        internal static readonly XDocument xdoc = XDocument.Load(@"Config\Config.xml");
        internal static string url = xdoc.XPathSelectElement("config/site").Value;
        internal static string urlExcelOnline = xdoc.XPathSelectElement("config/account/excelOnline").Attribute("url").Value;
        internal static string downloadedBook = xdoc.XPathSelectElement("config/account/excelOnline").Attribute("downloadedBook").Value;
        internal static string instanceName = xdoc.XPathSelectElement("config/instanceName").Value;
        internal static string iconImageExcelOnline = xdoc.XPathSelectElement("config/iconImg").Value;
        internal static string manifestFile = xdoc.XPathSelectElement("config/manifestFile").Attribute("file").Value;
        internal static string email = xdoc.XPathSelectElement("config/account/valid").Attribute("email").Value;
        internal static string password = xdoc.XPathSelectElement("config/account/valid").Attribute("password").Value;
        internal static string username = xdoc.XPathSelectElement("config/account/valid").Attribute("username").Value;
        internal static string clientId = xdoc.XPathSelectElement("config/account/msal").Attribute("clientId").Value;
        internal static string tenantId = xdoc.XPathSelectElement("config/account/msal").Attribute("tenantId").Value;
        internal static string redirectUri = xdoc.XPathSelectElement("config/account/msal").Attribute("redirectUri").Value;
        internal static string usernameAddin = xdoc.XPathSelectElement("config/account/valid").Attribute("usernameAddin").Value;
        internal static WebDriverWait? wait;

        // Initiate the By objects for elements
        internal static By loginWithMSAccountBtn = By.XPath(@"//button[.='Login with Microsoft Account']");
        internal static By imageIconInExcelOnline = By.XPath(@"//button[@id='AddinControl1']/span/i/img");
        internal static By signInEmailInputTxt = By.Name("loginfmt");
        internal static By nextBtn = By.Id("idSIButton9");
        internal static By passwordInputTxt = By.XPath(@"//input[@type='password']");
        internal static By signInBtn = By.Id("idSIButton9");
        internal static By emailGodaddyInputTxt = By.Id("username");
        internal static By passwordGoddadyInputTxt = By.Id("password");
        internal static By signInGodaddyBtn = By.Id("submitBtn");
        internal static By btnBackGodaddy = By.Id("idBtn_Back");
        internal static By logoKamehameha = By.XPath(@"//span[.='K']");
        internal static By uploadPDFFilesIcon = By.XPath(@"//i[@class='pi pi-upload']");
        internal static By searchFundInputTxt = By.XPath(@"//input[@role='searchbox']");
        internal static By ownedByKSCkb = By.XPath(@"//label[.='Owned by KS']");
        internal static By accountUsernamebtn = By.XPath(@"//button[.='" + username + "']");
        internal static By accountUsernameAddinbtn = By.XPath(@"//button[.='" + usernameAddin + "']");
        internal static By allowButtonDialogFADAddInEO = By.XPath(@"//input[@value='Allow']");
        internal static By logOutDropdown = By.XPath(@"//div[contains(@class,'account-options')]//button[2]");
        internal static By sessionExpiredPopup = By.XPath(@"//span[.='Session Expired']");
        internal static By loginSessionExpiredBtn = By.XPath(@"//span[.='Login']/ancestor::button");
        internal static By logOutBtn = By.XPath(@"//span[.='Log Out']");

        // Initiate the elements
        public IWebElement btnLogin(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(loginWithMSAccountBtn));
        }
        public IWebElement iconImageInExcelOnline(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(imageIconInExcelOnline));
        }
        public IWebElement txtEmail => Driver.Browser.FindElement(signInEmailInputTxt);
        public IWebElement btnNext(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(d => d.FindElement(nextBtn));
        }
        public IWebElement txtPassword(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(d => d.FindElement(passwordInputTxt));
        }
        public IWebElement btnSignIn(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(signInBtn));
        }
        public IWebElement txtEmailGodaddy => Driver.Browser.FindElement(emailGodaddyInputTxt);
        public IWebElement txtPasswordGodaddy => Driver.Browser.FindElement(passwordGoddadyInputTxt);
        public IWebElement btnLoginGodaddy(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(d => d.FindElement(signInGodaddyBtn));
        }
        public IWebElement kamehamehaLogo(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(d => d.FindElement(logoKamehameha));
        }
        public IWebElement iconUploadPDFFiles(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(d => d.FindElement(uploadPDFFilesIcon));
        }
        public IWebElement txtSearchFund => Driver.Browser.FindElement(searchFundInputTxt);
        public IWebElement ckbOwnedByKS(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(d => d.FindElement(ownedByKSCkb));
        }
        public IWebElement btnAccountUsername(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(d => d.FindElement(accountUsernamebtn));
        }
        public IWebElement btnAccountUsernameAddin(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(d=>d.FindElement(accountUsernameAddinbtn));
        }
        public IWebElement buttonAllowDialogFADAddInEO(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(allowButtonDialogFADAddInEO));
        }
        public IWebElement dropdownLogOut(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(logOutDropdown));
        }
        public IWebElement buttonLoginSessionExpired(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(ExpectedConditions.ElementIsVisible(loginSessionExpiredBtn));
        }
        public IWebElement btnLogOut(int timeoutInSeconds)
        {
            wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
            return wait.Until(d => d.FindElement(logOutBtn));
        }
    }

    internal sealed class LoginAction : BasePage<LoginAction, LoginPage>
    {
        #region LoginAction constructor
        private LoginAction() { }
        #endregion

        #region Login Items Action
        // Wait for element visible
        public LoginAction WaitForElementVisible(int timeoutInSeconds, By element)
        {
            if (timeoutInSeconds > 0)
            {
                WebDriverWait wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
                wait.Until(ExpectedConditions.ElementIsVisible(element));
            }
            return this;
        }

        // Wait for the new window is opened
        public LoginAction WaitForNewWindowOpen(int timeoutInSeconds)
        {
            if (timeoutInSeconds > 0)
            {
                WebDriverWait wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
                int previousWinCount = Driver.Browser.WindowHandles.Count;
                wait.Until(d => d.WindowHandles.Count == (previousWinCount + 1));
            }
            return this;
        }

        // Wait for the popup Window closed
        public LoginAction WaitForPopupWindowClosed(int timeoutInSeconds)
        {
            if (timeoutInSeconds > 0)
            {
                WebDriverWait wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
                int previousWinCount = Driver.Browser.WindowHandles.Count;
                wait.Until(d => d.WindowHandles.Count < 2);
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

        // (wait method 2) Wait for element load Done (can apply for performance)
        public LoginAction WaitingElementLoadDone(int timeout, By by)
        {
            int times = 0;
            while ((!IsElementPresent(by)) && (timeout < times))
            {
                times++;
                System.Threading.Thread.Sleep(1000);
            }

            if (times == timeout)
            {
                Console.WriteLine("The Element failed to load in " + timeout + " seconds");
                return this;
            }
            else
            {
                //Console.WriteLine("The Element Load Done in " + times + " seconds");
                return this;
            }
        }

        // Check if the popup MS still being shown, then switch to them main window to refresh browser
        public void CheckIfMSLoginPopupStillShown(int timeout)
        {
            int winCount = Driver.Browser.WindowHandles.Count;
            string winHandleBefore = Driver.Browser.WindowHandles[0];
            Driver.Browser.SwitchTo().Window(winHandleBefore); // --> Switch to the main window
            System.Threading.Thread.Sleep(3000);

            // Check if Microsoft Login popup still show (although the user is able to login successfully in Workbench)
            if (winCount == 2 && !IsElementPresent(LoginPage.loginWithMSAccountBtn))
            {
                // Switch to the main window
                Driver.Browser.SwitchTo().Window(winHandleBefore);
            }

            // Check if Microsoft Login popup still show (after clicking login button), then click login until successfully
            if (winCount == 2 && IsElementPresent(LoginPage.loginWithMSAccountBtn))
            {
                // Click login until successfully
                for (int times = 1; times <= timeout; times++)
                {
                    // Switch to the main window
                    Driver.Browser.SwitchTo().Window(winHandleBefore);
                    System.Threading.Thread.Sleep(500);

                    // Check if "login With Microsoft Account" Button is shown then click on it
                    if (IsElementPresent(LoginPage.loginWithMSAccountBtn))
                    {
                        ClickLogin(timeout);
                        System.Threading.Thread.Sleep(500);
                    }

                    winCount = Driver.Browser.WindowHandles.Count;
                    if (winCount == 1 || IsElementPresent(LoginPage.logOutDropdown))
                    {
                        break;
                    }
                    if (winCount == 2 && times == timeout)
                    {
                        Console.WriteLine("The element 'Login with MS Account' button failed to click in " + timeout + " seconds");
                        break;
                    }
                }
            }

            // Switch to the main window
            Driver.Browser.SwitchTo().Window(winHandleBefore);
            System.Threading.Thread.Sleep(1000);
        }
        
        // Wait for the Login button (ExcelOnline) is shown
        public void WaitForLoginButtonIsShown(string url, int timeout)
        {
            // Wait for 'Log out' button is shown
            int times = 0;
            while (!IsElementPresent(LoginPage.loginWithMSAccountBtn) && times < timeout)
            {
                times++;
                System.Threading.Thread.Sleep(1000);

                if (IsElementPresent(LoginPage.logOutDropdown))
                {
                    break;
                }
            }

            if (!IsElementPresent(LoginPage.loginWithMSAccountBtn) && times == timeout)
            {
                // This will shift focus back to main (default) content in which frame 'one' lies
                Driver.Browser.SwitchTo().DefaultContent();

                // Switch to the frame of Office Document Body (to interact with menu/button of MS Excel Online)
                System.Threading.Thread.Sleep(500);
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameOfficeExcelOnline);

                // Click on "FAD Add-in" button
                NavigationAction.Instance.ClickFADAddInButtonExcelOnline(10);

                // Wait for WorkBench Excel Online is displayed
                NavigationAction.Instance.WaitForElementVisible(10, NavigationPage.workBenchLeftPaneExcelOnline);

                // Switch to the FAD Add-in frame (to interact with Workbench for Excel Online)
                string? instanceName = null;
                if (url.Contains("sandbox"))
                {
                    instanceName = "sandbox";
                }
                if (url.Contains("conceptia"))
                {
                    instanceName = "conceptia";
                }
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameIdFADAddInCurrentInstance(instanceName));

                //Console.WriteLine("Unable to locate element: " + LoginPage.logOutDropdown + "");
                //BaseTestCase.ExtReportResult(false, "Unable to locate element: " + LoginPage.logOutDropdown + "");
            }

            // click Login button & then click Allow button
            if (IsElementPresent(LoginPage.loginWithMSAccountBtn))
            {
                ClickLogin(10);
                ClickbuttonAllowDialogFADAddInEO(10);
            }
        }

        // Wait for the Logout button (ExcelOnline) is shown
        public void WaitForLogOutButtonIsShown(int timeout)
        {
            // Wait for 'Log out' button is shown
            int times = 0;
            while (!IsElementPresent(LoginPage.logOutDropdown) && times < timeout)
            {
                times++;
                System.Threading.Thread.Sleep(1000);
            }

            if (!IsElementPresent(LoginPage.logOutDropdown) && times == timeout)
            {
                Console.WriteLine("Unable to locate element: " + LoginPage.logOutDropdown + "");
                BaseTestCase.ExtReportResult(false, "Unable to locate element: " + LoginPage.logOutDropdown + "");
            }

            // Wait for buttons is shown
            WaitForElementVisible(10, LoginPage.uploadPDFFilesIcon);
            WaitForElementVisible(10, LoginPage.ownedByKSCkb);
        }

        public bool CheckFileDownloadIsComplete(int timeout, string downloadPath, string fileName)
        {
            int times = 0;
            while (!File.Exists(downloadPath + @"\" + fileName) && times < timeout)
            {
                times++;
                System.Threading.Thread.Sleep(1000);
            }

            if (times == timeout)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool IsElementVisible(By searchElementBy)
        {
            try
            {
                return Driver.Browser.FindElement(searchElementBy).Displayed;

            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }

        // Wait for element Invisible
        public LoginAction WaitForElementInvisible(int timeoutInSeconds, By element)
        {
            if (timeoutInSeconds > 0)
            {
                WebDriverWait wait = new WebDriverWait(Driver.Browser, TimeSpan.FromSeconds(timeoutInSeconds));
                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(element));
            }
            return this;
        }

        public LoginAction NavigateSite(string url)
        {
            Driver.Browser.Navigate().GoToUrl(String.Concat(url));
            return this;
        }

        public LoginAction ClickLogin(int timeoutInSeconds)
        {
            this.Map.btnLogin(timeoutInSeconds).Click();
            return this;
        }

        public bool IsLoginWithMSAccountBtnShown(int timeoutInSeconds)
        {
            return this.Map.btnLogin(timeoutInSeconds).Displayed;
        }

        public string imageIconInExcelOnlineGetIcon(int timeoutInSeconds)
        {
            return this.Map.iconImageInExcelOnline(timeoutInSeconds).GetAttribute("src");
        }

        public bool IsUploadPDFFilesIconShown(int timeoutInSeconds)
        {
            return this.Map.iconUploadPDFFiles(timeoutInSeconds).Displayed;
        }

        public bool IsSearchFundTxtShown()
        {
            return this.Map.txtSearchFund.Displayed;
        }

        public bool IsAccountUserNameBtnShown(int timeoutInSeconds)
        {
            return this.Map.btnAccountUsername(timeoutInSeconds).Displayed;
        }

        public string AccountUserNameBtnGetText(int timeoutInSeconds)
        {
            return this.Map.btnAccountUsername(timeoutInSeconds).Text;
        }

        public string AccountUserNameAddinBtnGetText(int timeoutInSeconds)
        {
            return this.Map.btnAccountUsernameAddin(timeoutInSeconds).Text;
        }

        public string CkbOwnedByKSGetText(int timeoutInSeconds)
        {
            return this.Map.ckbOwnedByKS(timeoutInSeconds).Text;
        }

        public LoginAction ClickKamehamehaLogo(int timeoutInSeconds)
        {
            //this.Map.kamehamehaLogo(timeoutInSeconds).Click(); // --> Element Click Intercepted Exception

            // Try with javascript if Element Click Intercepted Exception
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].click();", this.Map.kamehamehaLogo(timeoutInSeconds));
            return this;
        }

        public LoginAction ClickUploadPDFFilesIcon(int timeoutInSeconds)
        {
            this.Map.iconUploadPDFFiles(timeoutInSeconds).Click();
            return this;
        }

        // Email - Sign in to your account (Microsoft site)
        public LoginAction EnterEmail(string txt)
        {
            this.Map.txtEmail.SendKeys(txt);
            return this;
        }

        // Next button (Microsoft site - popup 1st)
        public LoginAction ClickNext(int timeoutInSeconds)
        {
            this.Map.btnNext(timeoutInSeconds).Click();
            return this;
        }

        public LoginAction EnterPassword(int timeoutInSeconds, string txt)
        {
            this.Map.txtPassword(timeoutInSeconds).SendKeys(txt);
            return this;
        }

        public LoginAction ClickSignIn(int timeoutInSeconds)
        {
            //this.Map.btnSignIn(timeoutInSeconds).Click(); // --> Element Click Intercepted Exception

            // Try with javascript if Element Click Intercepted Exception
            IJavaScriptExecutor je = (IJavaScriptExecutor)Driver.Browser;
            je.ExecuteScript("arguments[0].click();", this.Map.btnSignIn(timeoutInSeconds));
            return this;
        }

        // Email & Password and Sign In - Godaddy (Microsoft)
        public LoginAction EnterEmailGodaddy(string txt)
        {
            this.Map.txtEmailGodaddy.SendKeys(txt);
            return this;
        }

        public LoginAction EnterPasswordGodaddy(string txt)
        {
            this.Map.txtPasswordGodaddy.SendKeys(txt);
            return this;
        }

        public LoginAction ClickbuttonAllowDialogFADAddInEO(int timeoutInSeconds)
        {
            this.Map.buttonAllowDialogFADAddInEO(timeoutInSeconds).Click();
            return this;
        }

        public LoginAction ClickSignInGodaddy(int timeoutInSeconds)
        {
            this.Map.btnLoginGodaddy(timeoutInSeconds).Click();
            return this;
        }

        public LoginAction ClickButtonLoginSessionExpired(int timeoutInSeconds)
        {
            this.Map.buttonLoginSessionExpired(timeoutInSeconds).Click();
            return this;
        }
        #endregion

        #region Login Built-in Actions
        public void LoginSite(string? url = null)
        {
            // Check if user login with a specific url, if no then will get url from config.xml (by default)
            if (url == null) url = LoginPage.url;

            // Variables declare
            string email = LoginPage.email;
            string password = LoginPage.password;
            try
            {
                NavigateSite(url);
                WaitForElementVisible(60, LoginPage.loginWithMSAccountBtn);

                // Store the current window handle (main window - Workbench)
                string winHandleBefore = Driver.Browser.CurrentWindowHandle;

                ClickLogin(10);
                WaitForNewWindowOpen(15);

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

                // Wait for the element of new window is opened (Email text field)
                WaitForElementVisible(10, LoginPage.signInEmailInputTxt);

                // Enter Email and then click Next button
                EnterEmail(email).ClickNext(10);

                // Switch to new window opened
                Driver.Browser.SwitchTo().Window(winHandleLast);
                WaitForElementVisible(10, LoginPage.signInGodaddyBtn);

                // Sign In - Enter password (godaddy - email is auto filled by Browser)
                EnterPasswordGodaddy(password)
                .ClickSignInGodaddy(10)
                .WaitForElementVisible(10, LoginPage.btnBackGodaddy)
                .ClickNext(10);

                // Check if the popup MS still being shown, then switch to them main window to click Login button again
                CheckIfMSLoginPopupStillShown(20);

                // Wait for buttons is shown
                WaitForElementVisible(10, LoginPage.logOutDropdown)
                .WaitForElementVisible(10, LoginPage.uploadPDFFilesIcon)
                .WaitForElementVisible(10, LoginPage.ownedByKSCkb);
            }
            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                // Warning
                BaseTestCase.ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong with Microsoft Login! Please check console log.");
            }
        }
        public void LoginExcelOnlineSite(string? url = null)
        {
            // Check if user login with a specific url, if no then will get url from config.xml (by default)
            if (url == null) url = LoginPage.url;

            // Log into the application (Workbench website)
            LoginSite(url);

            // Donwload Manifest file
            NavigateSite(LoginPage.url + @"/#/manifest");

            // Check File (manifest) is downloaded successful or not(timeout = 9s)
            string userProfileDownloadPath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + "Downloads";
            string fileName = @"manifest.xml";
            if (CheckFileDownloadIsComplete(9, userProfileDownloadPath, fileName) == true)
            {
                BaseTestCase.ExtReportResult(true, "The " + fileName + " file is downloaded successful");
                Console.WriteLine("The " + fileName + " file is downloaded successful: Pass");
            }
            else
            {
                BaseTestCase.ExtReportResult(false, "The " + fileName + " file is downloaded successful");
                Console.WriteLine("The " + fileName + " file is downloaded successful: Fail");
            }

            try
            {
                // Go to Excel Online
                string urlExcelOnline = LoginPage.urlExcelOnline;
                NavigateSite(urlExcelOnline);
            }
            catch (Exception exception)
            {
                // Print exception
                System.Console.WriteLine(exception);

                // Warning
                BaseTestCase.ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong with Microsoft url! Please check console log.");
            }
        }
        public void LoginExcelOnlineSiteAndAddIn(string? url = null, string? manifest = null)
        {
            // Check if user login with a specific url, if no then will get url from config.xml (by default)
            if (url == null) url = LoginPage.url;
            if (manifest == null) manifest = "manifest";

            // Variables declare
            const string notificationAddInExcelOnline = @"Add Microsoft Graph data to your document!";

            // Log into the application (Workbench website)
            LoginSite(url);

            // Download Manifest file
            NavigateSite(url + @"/#/" + manifest + "");

            // Check File (manifest) is downloaded successful or not(timeout = 9s)
            string userProfileDownloadPath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + "Downloads";
            string fileName = @"manifest.xml";
            if (CheckFileDownloadIsComplete(9, userProfileDownloadPath, fileName) == true)
            {
                BaseTestCase.ExtReportResult(true, "The " + fileName + " file is downloaded successful");
                Console.WriteLine("The " + fileName + " file is downloaded successful: Pass");
            }
            else
            {
                BaseTestCase.ExtReportResult(false, "The " + fileName + " file is downloaded successful");
                Console.WriteLine("The " + fileName + " file is downloaded successful: Fail");
            }

            try
            {
                // Go to Excel Online
                string urlExcelOnline = LoginPage.urlExcelOnline;
                NavigateSite(urlExcelOnline);

                // Upload manifest.xml
                string manifestFilePath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\";
                NavigationAction.Instance.UploadManifestExcelOnline(manifestFilePath, fileName);

                // Switch to the frame of Office Document Body (to interact with menu/button of MS Excel Online)
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameOfficeExcelOnline);

                // Wait for dialog notification (Excel Online - Add Microsoft Graph data to your document!) is displayed and then disappeared
                NavigationAction.Instance.WaitForElementOnFormToDisappear(10, NavigationPage.dialogNotificationExcelOnline(notificationAddInExcelOnline));

                // Click on "FAD Add-in" button
                NavigationAction.Instance.ClickFADAddInButtonExcelOnline(10);

                // Wait for WorkBench Excel Online Load Done
                NavigationAction.Instance.WaitingElementLoadDone(10, NavigationPage.workBenchLeftPaneExcelOnline);
                System.Threading.Thread.Sleep(3000);

                // Switch to the FAD Add-in frame (to interact with Workbench for Excel Online)
                string? instanceName = null;
                if (url.Contains("sandbox"))
                {
                    instanceName = "sandbox";
                }
                if (url.Contains("conceptia"))
                {
                    instanceName = "conceptia";
                }
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameIdFADAddInCurrentInstance(instanceName));

                // Check if the Login (With MS) button still being shown then click Login button & then click Allow button
                if (IsElementPresent(LoginPage.loginWithMSAccountBtn))
                {
                    ClickLogin(10)
                    .ClickbuttonAllowDialogFADAddInEO(10);
                }

                // Wait for Logout dropdown is shown
                WaitForElementVisible(10, LoginPage.logOutDropdown)
                .WaitForElementVisible(10, LoginPage.uploadPDFFilesIcon)
                .WaitForElementVisible(10, LoginPage.ownedByKSCkb);
            }
            catch (Exception exception)
            {
                // Print exception
                System.Console.WriteLine(exception);

                // Warning
                BaseTestCase.ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong with Microsoft url! Please check console log.");
            }
        }
        public LoginAction ClicklogOut()
        {
            this.Map.dropdownLogOut(10).Click();
            this.Map.btnLogOut(10).Click();
            WaitForElementVisible(3, LoginPage.loginWithMSAccountBtn);
            return this;
        }
        public void DeleteAllFilesManifest()
        {
            string manifestFilePath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\";
            string[] fileNames = Directory.GetFiles(manifestFilePath);
            foreach (string file in fileNames)
            {
                if (file.Contains("manifest"))
                {
                    File.Delete(file);
                }
            }
        }
        public void DeleteAllFilesWorkBookExcel()
        {
            string manifestFilePath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\";
            string[] fileNames = Directory.GetFiles(manifestFilePath);
            foreach (string file in fileNames)
            {
                if (file.Contains("Book"))
                {
                    File.Delete(file);
                }
            }
        }
        public void DeleteFilePath(string path, string fileName)
        {
            string[] fileNames = Directory.GetFiles(path);
            foreach (string file in fileNames)
            {
                if (file.Contains(fileName))
                {
                    File.Delete(file);
                }
            }
        }
        #endregion

        #region Login Built-in Actions (login an acc without Godaddy-login)
        public void LoginSiteNoGodaddy(string? url = null)
        {
            // Check if user login with a specific url, if no then will get url from config.xml (by default)
            if (url == null) url = LoginPage.url;

            // Variables declare
            string email = LoginPage.email;
            string password = LoginPage.password;
            try
            {
                NavigateSite(url);
                WaitForElementVisible(60, LoginPage.loginWithMSAccountBtn);

                // Store the current window handle (main window - Workbench)
                string winHandleBefore = Driver.Browser.CurrentWindowHandle;

                ClickLogin(10);
                WaitForNewWindowOpen(15);

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

                // Wait for the element of new window is opened (Email text field)
                WaitForElementVisible(10, LoginPage.signInEmailInputTxt);

                // Enter Email and then click Next button
                EnterEmail(email).ClickNext(10);

                // Switch to new window opened
                Driver.Browser.SwitchTo().Window(winHandleLast);
                System.Threading.Thread.Sleep(3000); //WaitForElementVisible(10, LoginPage.signInBtn);

                // Enter Password and then Click on SignIn button ...
                EnterPassword(10, password)
                .ClickSignIn(10);
                //.ClickNext(10); --> MS Login updated: Skip thip step 

                // Check if the popup MS still being shown, then switch to them main window to click Login button again
                CheckIfMSLoginPopupStillShown(20);

                // Wait for buttons is shown
                WaitForElementVisible(10, LoginPage.logOutDropdown)
                .WaitForElementVisible(10, LoginPage.uploadPDFFilesIcon)
                .WaitForElementVisible(10, LoginPage.ownedByKSCkb);
            }
            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                // Warning
                BaseTestCase.ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong with Microsoft Login! Please check console log.");
            }
        }
        public void LoginExcelOnlineSiteNoGodaddy(string? url = null)
        {
            // Check if user login with a specific url, if no then will get url from config.xml (by default)
            if (url == null) url = LoginPage.url;

            // Log into the application (Workbench website)
            LoginSiteNoGodaddy();

            // Donwload Manifest file
            NavigateSite(LoginPage.url + @"/#/manifest");

            // Check File (manifest) is downloaded successful or not(timeout = 9s)
            string userProfileDownloadPath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + "Downloads";
            string fileName = @"manifest.xml";
            if (CheckFileDownloadIsComplete(9, userProfileDownloadPath, fileName) == true)
            {
                BaseTestCase.ExtReportResult(true, "The " + fileName + " file is downloaded successful");
                Console.WriteLine("The " + fileName + " file is downloaded successful: Pass");
            }
            else 
            {
                BaseTestCase.ExtReportResult(false, "The " + fileName + " file is downloaded successful");
                Console.WriteLine("The " + fileName + " file is downloaded successful: Fail"); 
            }

            // Excel Online - Steps
            try
            {
                // Go to Excel Online
                NavigateSite(LoginPage.urlExcelOnline);
            }
            catch (Exception exception)
            {
                // Print exception
                System.Console.WriteLine(exception);

                // Warning
                BaseTestCase.ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong with Microsoft url! Please check console log.");
            }
        }
        public void LoginExcelOnlineSiteAndAddInNoGodaddy(string? url = null, string? manifest = null)
        {
            // Variables declare
            string? instanceName = null;
            const string notificationAddInExcelOnline = @"Add Microsoft Graph data to your document!";
            //By bookSavedTitleLoadDone = By.XPath(@"//span[@data-unique-id='DocumentTitleSaveStatus' and .='Saved']"); //'Saving...']
            By bookSavedTitleLoadDone = By.XPath(@"//i[contains(@data-icon-name,'CloudCheckmark') and @aria-hidden='true']");
            By bookSaving = By.XPath(@"//span[@data-unique-id='DocumentTitleSaveStatus' and .='Saving...']");
            By shareMenuPopup = By.XPath(@"//div[@role='heading' and .='Use Teams to collaborate on files']");

            // Check if user login with a specific url, if no then will get url from config.xml (by default)
            if (url == null) url = LoginPage.url;
            if (manifest == null) manifest = "manifest";

            // Log into the application (Workbench website)
            LoginSiteNoGodaddy(url);

            // Download Manifest file
            NavigateSite(url + @"/#/" + manifest + "");

            // Check File (manifest) is downloaded successful or not(timeout = 9s)
            string userProfileDownloadPath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + "Downloads";
            string fileName = "" + manifest + ".xml";
            if (CheckFileDownloadIsComplete(9, userProfileDownloadPath, fileName) == true) 
            {
                BaseTestCase.ExtReportResult(true, "The " + fileName + " file is downloaded successful");
                Console.WriteLine("The " + fileName + " file is downloaded successful: Pass");
            }
            else
            {
                BaseTestCase.ExtReportResult(false, "The " + fileName + " file is downloaded successful");
                Console.WriteLine("The " + fileName + " file is downloaded successful: Fail");
            }

            // Excel Online - Steps
            try
            {
                // Go to Excel Online
                NavigateSite(LoginPage.urlExcelOnline);

                // Upload manifest.xml
                string manifestFilePath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\";
                NavigationAction.Instance.UploadManifestExcelOnlineNoGodaddy(manifestFilePath, fileName); // UploadManifestExcelOnline()

                // Switch to the frame of Office Document Body (to interact with menu/button of MS Excel Online)
                //System.Threading.Thread.Sleep(3000);
                Driver.Browser.SwitchTo().DefaultContent();
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameOfficeExcelOnline);

                // Wait for dialog notification (Excel Online - Add Microsoft Graph data to your document!) is displayed
                NavigationAction.Instance.WaitForElementVisible(10, NavigationPage.dialogNotificationExcelOnline(notificationAddInExcelOnline));

                // Click on Close button of 'Notification' dialog
                NavigationAction.Instance.ClickNotificationCloseButtonExcelOnline(10);

                // Click on "FAD Add-in" button
                NavigationAction.Instance.ClickFADAddInButtonExcelOnline(10); Thread.Sleep(3000);

                // Check if book Saved Title is not changed then wait for it is changed
                //if (IsElementPresent(bookSavedTitleLoadDone)) WaitForElementVisible(10, bookSavedTitleLoadDone); --> not good
                //else WaitForElementVisible(10, bookSavedTitleLoadDone); --> not good

                // Refresh page if Book still at 'Saving' status
                int time = 0;
                while (IsElementPresent(bookSaving) == true && time < 6)
                {
                    if (IsElementPresent(bookSaving) == false) { break; }
                    if (IsElementPresent(bookSaving) == true && time == 5) 
                    { 
                        Driver.Browser.Navigate().Refresh();
                        // Switch to the FAD Add-in frame (to interact with Workbench for Excel Online)
                        if (url.Contains("sandbox"))
                        {
                            instanceName = "sandbox";
                        }
                        if (url.Contains("conceptia"))
                        {
                            instanceName = "conceptia";
                        }
                        NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameIdFADAddInCurrentInstance(instanceName));
                    }
                    Thread.Sleep(1000);
                    time++;
                }

                // Wait for WorkBench Excel Online is displayed
                NavigationAction.Instance.WaitForElementVisible(10, NavigationPage.workBenchLeftPaneExcelOnline);

                // Check if 'Share' menu show popup 'Use Teams to collaborate on files' then click 'Got it' button
                if (IsElementPresent(shareMenuPopup))
                {
                    // CLick 'Got it' button
                    NavigationAction.Instance.MouseClick(3, NavigationPage.gotItBtnExcelOnline);
                }

                // Switch to the FAD Add-in frame (to interact with Workbench for Excel Online)
                //string? instanceName = null;
                if (url.Contains("sandbox"))
                {
                    instanceName = "sandbox";
                }
                if (url.Contains("conceptia"))
                {
                    instanceName = "conceptia";
                }
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameIdFADAddInCurrentInstance(instanceName));

                // Check if the Login (With MS) button still being shown then click Login button & then click Allow button
                WaitForLoginButtonIsShown(url, 15);
                //WaitForLogOutButtonIsShown(30);

                // Wait for buttons is shown
                WaitForElementVisible(15, LoginPage.logOutDropdown);
                WaitForElementVisible(15, LoginPage.uploadPDFFilesIcon);
                WaitForElementVisible(15, LoginPage.ownedByKSCkb);
            }
            catch (Exception exception)
            {
                // Print exception
                System.Console.WriteLine(exception);

                // Warning
                BaseTestCase.ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong with Microsoft url! Please check console log.");
            }
        }
        public void LoginExcelOnlineSiteAfterRefresh(string? url = null, string? manifest = null)
        {
            // Variables declare
            string? instanceName = null;
            const string notificationAddInExcelOnline = @"Add Microsoft Graph data to your document!";
            By bookSavedTitleLoadDone = By.XPath(@"//i[contains(@data-icon-name,'CloudCheckmark') and @aria-hidden='true']");
            By bookSaving = By.XPath(@"//span[@data-unique-id='DocumentTitleSaveStatus' and .='Saving...']");
            By shareMenuPopup = By.XPath(@"//div[@role='heading' and .='Use Teams to collaborate on files']");

            // Check if user login with a specific url, if no then will get url from config.xml (by default)
            if (url == null) url = LoginPage.url;
            if (manifest == null) manifest = "manifest";

            // Check File (manifest) is downloaded successful or not(timeout = 9s)
            string userProfileDownloadPath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + "Downloads";
            string fileName = "" + manifest + ".xml";
            if (CheckFileDownloadIsComplete(9, userProfileDownloadPath, fileName) == true)
            {
                BaseTestCase.ExtReportResult(true, "The " + fileName + " file is downloaded successful");
                Console.WriteLine("The " + fileName + " file is downloaded successful: Pass");
            }
            else
            {
                BaseTestCase.ExtReportResult(false, "The " + fileName + " file is downloaded successful");
                Console.WriteLine("The " + fileName + " file is downloaded successful: Fail");
            }

            // Excel Online - Steps
            try
            {
                // Upload manifest.xml
                string manifestFilePath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\";
                NavigationAction.Instance.UploadManifestExcelOnlineNoGodaddy(manifestFilePath, fileName);

                // Switch to the frame of Office Document Body (to interact with menu/button of MS Excel Online)
                Driver.Browser.SwitchTo().DefaultContent();
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameOfficeExcelOnline);

                // Wait for dialog notification (Excel Online - Add Microsoft Graph data to your document!) is displayed
                NavigationAction.Instance.WaitForElementVisible(10, NavigationPage.dialogNotificationExcelOnline(notificationAddInExcelOnline));

                // Click on Close button of 'Notification' dialog
                NavigationAction.Instance.ClickNotificationCloseButtonExcelOnline(10);

                // Click on "FAD Add-in" button
                NavigationAction.Instance.ClickFADAddInButtonExcelOnline(10); Thread.Sleep(3000);

                // Refresh page if Book still at 'Saving' status
                int time = 0;
                while (IsElementPresent(bookSaving) == true && time < 6)
                {
                    if (IsElementPresent(bookSaving) == false) { break; }
                    if (IsElementPresent(bookSaving) == true && time == 5)
                    {
                        Driver.Browser.Navigate().Refresh();
                        // Switch to the FAD Add-in frame (to interact with Workbench for Excel Online)
                        if (url.Contains("sandbox"))
                        {
                            instanceName = "sandbox";
                        }
                        if (url.Contains("conceptia"))
                        {
                            instanceName = "conceptia";
                        }
                        NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameIdFADAddInCurrentInstance(instanceName));
                    }
                    Thread.Sleep(1000);
                    time++;
                }

                // Wait for WorkBench Excel Online is displayed
                NavigationAction.Instance.WaitForElementVisible(10, NavigationPage.workBenchLeftPaneExcelOnline);

                // Check if 'Share' menu show popup 'Use Teams to collaborate on files' then click 'Got it' button
                if (IsElementPresent(shareMenuPopup))
                {
                    // CLick 'Got it' button
                    NavigationAction.Instance.MouseClick(3, NavigationPage.gotItBtnExcelOnline);
                }

                // Switch to the FAD Add-in frame (to interact with Workbench for Excel Online)
                if (url.Contains("sandbox"))
                {
                    instanceName = "sandbox";
                }
                if (url.Contains("conceptia"))
                {
                    instanceName = "conceptia";
                }
                NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameIdFADAddInCurrentInstance(instanceName));

                // Check if the Login (With MS) button still being shown then click Login button & then click Allow button
                WaitForLoginButtonIsShown(url, 15);

                // Wait for buttons is shown
                WaitForElementVisible(15, LoginPage.logOutDropdown);
                WaitForElementVisible(15, LoginPage.uploadPDFFilesIcon);
                WaitForElementVisible(15, LoginPage.ownedByKSCkb);
            }
            catch (Exception exception)
            {
                // Print exception
                System.Console.WriteLine(exception);

                // Warning
                BaseTestCase.ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong with Microsoft url! Please check console log.");
            }
        }
        #endregion

        #region Login Built-in Actions with Solution 2: Bypass Godaddy issue (Opened chrome browser) --> not used
        /*
        public void LoginSiteBypassGodaddy(string? url = null)
        {
            // Check if user login with a specific url, if no then will get url from config.xml (by default)
            if (url == null) url = LoginPage.url;

            try
            {
                // Go to Workbench site
                NavigateSite(url);

                //// Check if Session Expired Popup is shown then click on Login button
                //if(IsElementPresent(LoginPage.sessionExpiredPopup))
                //{
                //    // Click on Login button of Session Expired popup
                //    ClickButtonLoginSessionExpired(10);

                //    // Check if the popup MS still being shown, then switch to them main window to click Login button again
                //    CheckIfMSLoginPopupStillShown(5);
                //}

                //// Check if the popup MS is being shown
                //if (IsElementPresent(LoginPage.loginWithMSAccountBtn))
                //{
                //    // Click "Login With Microsoft Account" button
                //    ClickLogin(10);

                //    // Check if the popup MS still being shown, then switch to them main window to click Login button again
                //    CheckIfMSLoginPopupStillShown(5);
                //}

                // Wait for buttons is shown
                WaitForElementVisible(10, LoginPage.logOutDropdown)
                .WaitForElementVisible(10, LoginPage.uploadPDFFilesIcon)
                .WaitForElementVisible(10, LoginPage.ownedByKSCkb);
            }
            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                // Warning
                BaseTestCase.ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong with Microsoft Login! Please check console log.");
            }
        }
        public void LoginExcelOnlineSiteBypassGodaddy(string? url = null)
        {
            // Check if user login with a specific url, if no then will get url from config.xml (by default)
            if (url == null) url = LoginPage.url;

            try
            {
                // Go to Workbench site
                NavigateSite(url);

                // Wait for buttons is shown
                WaitForElementVisible(10, LoginPage.logOutDropdown)
                .WaitForElementVisible(10, LoginPage.uploadPDFFilesIcon)
                .WaitForElementVisible(10, LoginPage.ownedByKSCkb);

                // Go to Excel Online
                string urlExcelOnline = LoginPage.urlExcelOnline;
                NavigateSite(urlExcelOnline);
            }
            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                // Warning
                BaseTestCase.ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong with Microsoft Login! Please check console log.");
            }
        }
        */
        #endregion
    }
}
