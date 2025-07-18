using AventStack.ExtentReports;
using NUnit.Framework;
using SeleniumGendKS.Core.BaseTestCase;
using SeleniumGendKS.Core.Selenium;
using SeleniumGendKS.Pages;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.XPath;

namespace SeleniumGendKS.Tests.Performance_Testing
{
    [TestFixture]
    internal class PerformanceTestSuite : BaseTestCase
    {
        #region Variables declare
        //private static readonly XDocument xdoc = LoginPage.xdoc;
        [Obsolete]
        readonly ExtentReports rep = ExtReportgetInstance();
        private string? summaryTC;

        ////bool verifypoint;
        //List<bool> verifypoints = new List<bool>();
        #endregion

        //[SetUp]
        //public override void SetupTest()
        //{
        //    //Data-driven for site testing
        //    string url = xdoc.XPathSelectElement("config/site").Value;
        //    //verifyPoints.Clear();
        //    Driver.StartBrowser();
        //    LoginAction.Instance.NavigateSite(url);
        //}

        //[Test, Category("PerformanceTestSuite")]
        [Ignore("Ignore this")]
        public void Measure_DxD_Report()
        {
            #region Variables declare
            string urlInstance = LoginPage.url;
            string cambridgeFund = "Blackstone Group, The";
            const string sourceIcon = "C";
            //verifypoints.Clear();
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Performance DxD Report Excel Online");
            try
            {
                // Delete Manifest, WorkBook files
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "Book");

                // Login Excel Online and then add (manifest.xml) FAD Add-in
                LoginAction.Instance.LoginExcelOnlineSiteAndAddInNoGodaddy(urlInstance);

                // Search a Cambridge Fund (Blackstone)
                FundDashboardAction.Instance.InputNameToSearchFund(10, "Blackstone ")
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResultsWithItemSource(cambridgeFund, sourceIcon))
                                            .ClickFundNameReturnOfResults(10, cambridgeFund, sourceIcon)
                                            .WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);

                /// User Input
                // Select Report Type = Deal by Deal
                FADAddInAction.Instance.ClickToCheckTheCheckbox("Deal by Deal")
                                       .WaitForElementVisible(10, FADAddInPage.labelDropdown(FADAddInPage.weightingMethodology))
                                       .InputTxtDatePickerTitle(10, FADAddInPage.dateSelection, "2022", "Mar", "31")
                                       .InputTxtNameCRBMRow(10, 1, "s&p") // input 1st CRBM
                                       .WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("S&P 500 INDEX"))
                                       .ClickNameCRBMReturnOfResults(10, "S&P 500 INDEX").WaitForElementInvisible(10, AddEditFundPage.overlayDropdown)
                                       .InputNumberBetaCRBMRow(10, 1, "1")
                                       .InputNumberGrossExposureCRBMRow(10, 1, "80")
                                       .ClickAndSelectItemInDropdown(10, FADAddInPage.weightingMethodology, "Percent Fund Final")
                                       .ClickAndSelectItemInDropdown(10, FADAddInPage.attributionCategory1, "% of Fund")
                                       .ClickAndSelectItemInDropdown(10, FADAddInPage.attributionCategory2, "Equal Weight")
                                       .ClickAndSelectItemInDropdown(10, FADAddInPage.attributionCategory3, "Invested Capital");

                int numberOfTimes = 3;
                for (int i = 1; i <= numberOfTimes; i++)
                {
                    // Click Run button to run Report
                    ClickRunButton(cambridgeFund);

                    // Switch to the FAD Add-in frame (to interact with Workbench for Excel Online)
                    NavigationAction.Instance.SwitchToFrameWithWaitMethod(10, NavigationPage.frameIdFADAddInCurrentInstance(LoginPage.instanceName));
                }

                // Delete Manifest and WorkBook file
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "Book");

                
            }
            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                // Delete Manifest, WorkBook files 
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "manifest");
                LoginAction.Instance.DeleteFilePath(Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads\", "Book");

                // Warning
                ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong! Please check console log.");
            }
            #endregion

            #region Result
            ////Passed or Failed
            //bool result = true;
            //foreach (bool verify in verifypoints)
            //{
            //    System.Console.WriteLine(verify);
            //    result = result && verify;
            //}
            //Assert.IsTrue(result);
            #endregion
        }

        public void ClickRunButton(string fundType)
        {
            // Variables declare
            Stopwatch stopwatch = new Stopwatch();

            // Click Run button to run Report
            FADAddInAction.Instance.ClickRunButton(10);
            stopwatch.Start();
            FADAddInAction.Instance.WaitForLoadingIconToDisappear(189, FundDashboardPage.loadingIcon)
                                   .WaitforExcelWebDxDOutputRenderDone(60);
            stopwatch.Stop();

            // Write result to log/report 
            Console.WriteLine(summaryTC = "Running DxD Report '" + fundType + "' in " + stopwatch.ElapsedMilliseconds + "ms");
            test.Log(Status.Info, summaryTC);
        }
    }
}
