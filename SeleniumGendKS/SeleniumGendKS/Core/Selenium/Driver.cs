﻿using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.IE;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.XPath;
using OpenQA.Selenium.Edge;
using System.Diagnostics;

namespace SeleniumGendKS.Core.Selenium
{
    public static class Driver
    {
        private static Process proc = new Process();
        private static WebDriverWait? browserWait;

        private static IWebDriver? browser;

        public static IWebDriver? Browser
        {
            get
            {
                if (browser == null)
                {
                    throw new NullReferenceException("The WebDriver browser instance was not initialized. You should first call the method class");
                }
                return browser;
            }
            private set
            {
                browser = value;
            }
        }

        public static WebDriverWait? BrowserWait
        {
            get
            {
                if (browserWait == null || browser == null)
                {
                    throw new NullReferenceException("The WebDriver browser instance was not initialized. You should first call the method class");
                }
                return browserWait;
            }
            private set
            {
                browserWait = value;
            }
        }

        public static void ClearBrowserCache()
        {
            Browser.Manage().Cookies.DeleteAllCookies();
            //System.Threading.Thread.Sleep(7000); // wait 7 seconds to clear cookies
        }

        public static void StartBrowser(string? browserType = null, int defaultTimeOut = 30)
        {
            XDocument xdoc = XDocument.Load(@"Config\Config.xml");
            var browser = xdoc.XPathSelectElement("config/browser").Value;

            // Using Chrome browser by default, if no then get default browser at Config.xml
            if (browserType == null) browserType = browser; //"chrome";

            // Open browser type (chrome; firefox; ie ...)
            string webDriverPath = @"Core\WebDrivers\";
            switch (browserType)
            {
                case "CHROME": case "Chrome": case "chrome":
                    ChromeOptions options = new ChromeOptions();
                    options.AddUserProfilePreference("download.prompt_for_download", false);
                    options.AddUserProfilePreference("disable-popup-blocking", "true");
                    options.AddUserProfilePreference("safebrowsing.enabled", "true");
                    Browser = new ChromeDriver(Path.Combine(Environment.CurrentDirectory, webDriverPath), options);

                    /*
                    // (Solution 1) Fix (By pass) Godaddy Your browser is a bit unusual ... --> Not Working
                    options.BinaryLocation = @"C:\Program Files\Google\Chrome\Application\chrome.exe";
                    options.AddArguments(@"user-data-dir=E:\Chrome_Profile"); // Use custom profile (also called user data directory)
                    //options.AddArguments(@"user-data-dir=E:\Selenium\Chrome_Test_Profile");
                    //options.AddArguments("--incognito");
                    //options.AddArgument("--enable-javascript");
                    Browser = new ChromeDriver(Path.Combine(Environment.CurrentDirectory, webDriverPath), options);
                    //Browser.Manage().Cookies.DeleteAllCookies();
                    //Browser.Navigate().GoToUrl("chrome://settings/clearBrowserData");
                    //Browser.FindElement(By.XPath("//settings-ui")).SendKeys(OpenQA.Selenium.Keys.Enter);
                    //System.Threading.Thread.Sleep(5000);
                    */

                    //// (Solution 2) Fix (By pass) Godaddy Your browser is a bit unusual ...
                    ////Process proc = new Process();
                    //proc.StartInfo.FileName = @"C:\Program Files\Google\Chrome\Application\chrome.exe";
                    //proc.StartInfo.Arguments = "-remote-debugging-port=9222 --user-data-dir=" + @"E:\Selenium\Chrome_Test_Profile"; // + " --incognito"
                    //proc.Start();
                    //ChromeOptions options = new ChromeOptions();
                    //options.DebuggerAddress = "localhost: 9222"; // or 9222 / 9014 //options.AddArgument("–proxy - server = 8.8.8.8:8080");
                    //Browser = new ChromeDriver(Path.Combine(Environment.CurrentDirectory, webDriverPath), options);
                    Browser.Manage().Window.Maximize();
                    break;

                case "FIREFOX": case "FireFox": case "Firefox": case "firefox":
                    //XDocument xdoc = XDocument.Load(@"Config\Config.xml");
                    string firefoxBinaryPath = xdoc.XPathSelectElement("config/gecko").Attribute("path").Value;
                    Browser = new FirefoxDriver(new string(firefoxBinaryPath), new FirefoxOptions());
                    ClearBrowserCache();
                    break;

                case "IE": case "ie":
                    Browser = new InternetExplorerDriver(Path.Combine(Environment.CurrentDirectory, webDriverPath));
                    Browser.Manage().Window.Maximize();
                    //ClearBrowserCache();
                    break;

                case "Edge": case "edge":
                    Browser = new EdgeDriver(Path.Combine(Environment.CurrentDirectory, webDriverPath));
                    Browser.Manage().Window.Maximize();
                    //ClearBrowserCache();
                    break;

                default:
                    break;
            }
            BrowserWait = new WebDriverWait(Browser, TimeSpan.FromSeconds(defaultTimeOut));
        }

        public static void StopBrowser()
        {
            Browser.Quit();
            //proc.CloseMainWindow(); // --> apply for Godaddy Your browser is a bit unusual ...
            Browser = null;
            BrowserWait = null;
        }

        public static void TakeScreenShot(string screenShotName)
        {
            //XDocument xdoc = XDocument.Load(@"Config\Config.xml");
            //string screenshotPath = xdoc.XPathSelectElement("config/screenshot/screenshotPath").Value;
            //string screenshotFormat = xdoc.XPathSelectElement("config/screenshot/screenshotFormat").Value;

            //Screenshot ss = ((ITakesScreenshot)Driver.browser).GetScreenshot();
            //ss.SaveAsFile(screenshotPath + screenShotName + screenshotFormat, ScreenshotImageFormat.Jpeg);
        }
    }
}