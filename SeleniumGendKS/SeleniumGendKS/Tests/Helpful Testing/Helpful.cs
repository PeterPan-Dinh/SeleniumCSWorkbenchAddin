using AventStack.ExtentReports;
using NUnit.Framework;
using SeleniumGendKS.Core.BaseTestCase;
using SeleniumGendKS.Core.ExcelComparision;
using SeleniumGendKS.Core.Selenium;
using SeleniumGendKS.Pages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.XPath;

namespace SeleniumGendKS.Tests.Helpful_Testing
{
    [TestFixture]
    internal class Helpful : BaseTestCase
    {
        // Variables declare
        private static readonly XDocument xdoc = LoginPage.xdoc;

        [SetUp]
        public override void SetupTest()
        {
            //Data-driven for site testing
            string url = xdoc.XPathSelectElement("config/site").Value;
            verifyPoints.Clear();
            //Driver.StartBrowser();
            //LoginAction.Instance.NavigateSite(url);
        }

        //[Test, Category("Helpful")]
        [Ignore("Ignore this")]
        //[Obsolete]
        public void HelpfulTesting()
        {
            // Variables declare
            string ExcelFileDownloadedPath = Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + @"Downloads";
            string ExcelFileBaselinePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Tests\Documents\Excel files\SinglePublicReport_Baseline\");
            string fileNameBaseline = @"Book_OasisJapan.xlsx";
            string fundNameAlbourne = "Oasis Japan Strategic Fund";
            string sheetName;

            //Verify content of the downloaded file by doing compare with the baseline excel files.
            sheetName = "Raw Data";
            verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + @"Book.xlsx", ExcelFileBaselinePath + fileNameBaseline, sheetName);
            verifyPoints.Add("Verify the information of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundNameAlbourne + ")", verifyPoint);
            sheetName = "WorkBench";
            verifyPoint = ExcelComparision.IsSheetsEqual(ExcelFileDownloadedPath + @"\" + @"Book.xlsx", ExcelFileBaselinePath + fileNameBaseline, sheetName);
            verifyPoints.Add("Verify the information of Single Public Report Excel (sheet: " + sheetName + ") is generated correctly. (" + fundNameAlbourne + ")", verifyPoint);

            // FAD ADDIN add CRBM by using System.Windows.Forms
            // Using System.Windows.Forms to BYPASS issue interact with FAD ADD-IN (Run report show 'cannot retrieve...' although crbm had already been added)
            //FADAddInAction.Instance.InputTxtNameCRBMRow(10, 1, "s&p");
            //System.Windows.Forms.SendKeys.SendWait(@"^a"); // --> Ctrl + a
            //System.Windows.Forms.SendKeys.SendWait(@"S&P 500 INDEX");
            //FADAddInAction.Instance.WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("S&P 500 INDEX"));
            //System.Windows.Forms.SendKeys.SendWait(@"{DOWN}"); System.Windows.Forms.SendKeys.SendWait(@"{ENTER}");
            //System.Windows.Forms.SendKeys.SendWait(@"{TAB}"); System.Windows.Forms.SendKeys.SendWait(@"{TAB}"); //System.Windows.Forms.SendKeys.SendWait(@"^a");
            ////System.Windows.Forms.SendKeys.SendWait(@"1"); System.Windows.Forms.SendKeys.SendWait(@"{TAB}"); System.Windows.Forms.SendKeys.SendWait(@"^a");
            //System.Windows.Forms.SendKeys.SendWait(@"{TAB}");
            //System.Windows.Forms.SendKeys.SendWait(@"80"); System.Windows.Forms.SendKeys.SendWait(@"{TAB}");

            //FADAddInAction.Instance.ClickCRBMAddButton(10); // Add 2nd CRBM
            //FADAddInAction.Instance.InputTxtNameCRBMRow(10, 2, "s&p");
            //System.Windows.Forms.SendKeys.SendWait(@"^a");
            //System.Windows.Forms.SendKeys.SendWait(@"MSCI EM LATIN AMERICA ex BRAZIL Net Return USD Index");
            //FADAddInAction.Instance.WaitForElementVisible(20, FADAddInPage.nameCRBMReturnOfResults("MSCI EM LATIN AMERICA ex BRAZIL Net Return USD Index"));
            //System.Windows.Forms.SendKeys.SendWait(@"{DOWN}"); System.Windows.Forms.SendKeys.SendWait(@"{ENTER}");
            //System.Windows.Forms.SendKeys.SendWait(@"{TAB}"); System.Windows.Forms.SendKeys.SendWait(@"{TAB}"); //System.Windows.Forms.SendKeys.SendWait(@"^a");
            ////System.Windows.Forms.SendKeys.SendWait(@"1"); System.Windows.Forms.SendKeys.SendWait(@"{TAB}"); System.Windows.Forms.SendKeys.SendWait(@"^a");
            //System.Windows.Forms.SendKeys.SendWait(@"{TAB}");
            //System.Windows.Forms.SendKeys.SendWait(@"-4.5"); System.Windows.Forms.SendKeys.SendWait(@"{TAB}");
        }

        [TearDown]
        public override void TeardownTest() // virtual void TeardownTest
        {
            //Driver.StopBrowser();

            #region Result (Passed or Failed)
            bool result = true;
            foreach (var verify in verifyPoints)
            {
                string step_result = verify.Value ? "Pass" : "Fail";
                Console.WriteLine(verify.Key + " : " + step_result);
                result = result && verify.Value;
            }
            Assert.That(result);
            #endregion
        }
    }
}
