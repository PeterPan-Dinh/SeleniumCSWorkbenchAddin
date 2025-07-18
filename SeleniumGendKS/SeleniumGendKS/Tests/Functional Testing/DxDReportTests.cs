using Newtonsoft.Json.Linq;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using SeleniumGendKS.FADAddInApi;
using SeleniumGendKS.FADAddInApi.PredefinedScenarios;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.XPath;

// Notes: Please turn on your VPN and make sure update latest msalIdtoken (at BaseFunctionTest) before running API Testing

namespace SeleniumGendKS.Tests.Functional_Testing
{
    [TestFixture]
    internal class DxDReportTests : BaseFunctionTest
    {
        #region Initiate variables
        internal static string workbenchApi = xdoc.XPathSelectElement("config/webApis").Attribute("WorkbenchApi").Value;
        internal static string filePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"FADAddInApi\Documents\");
        internal static string? fileName = null;
        #endregion

        #region TestMethod
        [Test, Category("API Smoke Tests")]
        public void ST001_DxDReport()
        {
            #region Variables declare
            const string currency = "USD";
            const string effective_date = "2021-06-30";
            const string main_weight = "Custom Weight";
            const string attribution_category_1 = "Equal Weight";
            const string attribution_category_2 = "% of Fund";
            const string attribution_category_3 = "% of Fund";
            const string custom_risk_benchmarks = "[{" 
                                                  + "\n" + "\"benchmark_id\"" + " : " + "\"" + "SPTR Index.USD" + "\","
                                                  + "\n" + "\"beta\"" + " : " + "1,"
                                                  + "\n" + "\"exposure\"" + " : " + "0.8,"
                                                  + "\n" + "\"return_type\"" + " : " + "\"\"" + "\n" +
                                                  "}]";
            const string manager = "VGO Capital Partners";
            const string funds = "[" + "\n" +
                                    "{" + "\n" + "\"fund_name\"" + " : " + "\"" + "VGO Special Situations Fund I" + "\","
                                        + "\n" + "\"fund_size\"" + " : " + "0" + "\n" 
                                  + "}," +"\n" +
                                    "{" + "\n" + "\"fund_name\"" + " : " + "\"" + "VGO Special Situations Fund II" + "\","
                                        + "\n" + "\"fund_size\"" + " : " + "288" + "\n"
                                  + "}" + "\n" +
                                 "]";
            var body = "{" + "\n" + "\"currency\"" + " : " + "\"" + currency + "\","
                           + "\n" + "\"effective_date\"" + " : " + "\"" + effective_date + "\","
                           + "\n" + "\"main_weight\"" + " : " + "" + "\"" + main_weight + "\","
                           + "\n" + "\"attribution_category_1\"" + " : " + "\"" + attribution_category_1 + "\","
                           + "\n" + "\"attribution_category_2\"" + " : " + "\"" + attribution_category_2 + "\","
                           + "\n" + "\"attribution_category_3\"" + " : " + "\"" + attribution_category_3 + "\","
                           + "\n" + "\"custom_risk_benchmarks\"" + " : " + custom_risk_benchmarks + ","
                           + "\n" + "\"manager\"" + " : " + "\"" + manager + "\","
                           + "\n" + "\"funds\"" + " : " + funds + "\n" +
                       "}";
            #endregion

            #region Check if api of Sandbox or Staging then get data (on that site)
            if (workbenchApi.Contains("sandbox"))
            {
                fileName = "DxDReportOutput.json";
            }
            if (workbenchApi.Contains("conceptia"))
            {
                fileName = "DxDReportStagingOutput.json";
            }
            #endregion

            #region Run Tests
            // Get DxD Report
            var dxDReport = WorkbenchApi.DxDReport(body, msalIdtoken);
            Assert.That(dxDReport.StatusCode, Is.EqualTo(HttpStatusCode.OK));

            // Parse IRestResponse to JObject
            JObject dxDReportJs = JObject.Parse(dxDReport.Content);
            ClassicAssert.AreEqual("200", dxDReportJs["statusCode"].ToString());
            ClassicAssert.AreEqual("application/json", dxDReportJs["headers"]["content-type"].ToString());
            //Assert.AreEqual("0.4.6", dxDReportJs["headers"]["version"].ToString());
            JObject dxDReportJsBL = JObject.Parse(File.ReadAllText(filePath + fileName));
            dxDReportJsBL["headers"]["date"] = dxDReportJs["headers"]["date"];
            ClassicAssert.AreEqual(dxDReportJs["headers"]["date"], dxDReportJsBL["headers"]["date"]);
            ClassicAssert.AreEqual(dxDReportJs["body"]["pme"], dxDReportJsBL["body"]["pme"]);
            ClassicAssert.AreEqual(dxDReportJs["body"]["gross_total_alpha"], dxDReportJsBL["body"]["gross_total_alpha"]);
            ClassicAssert.AreEqual(dxDReportJs["body"]["fund_table"], dxDReportJsBL["body"]["fund_table"]);
            ClassicAssert.AreEqual(dxDReportJs["body"]["results_table"], dxDReportJsBL["body"]["results_table"]);
            ClassicAssert.AreEqual(dxDReportJs["body"]["gta_table"], dxDReportJsBL["body"]["gta_table"]);
            //Assert.AreEqual(dxDReportJs["body"]["attr1_table"], dxDReportJsBL["body"]["attr1_table"]);
            //Assert.AreEqual(dxDReportJs["body"]["base"], dxDReportJsBL["body"]["base"]);
            ClassicAssert.AreEqual(dxDReportJs["body"]["final_row"], dxDReportJsBL["body"]["final_row"]);

            //var sortReportJs = new JObject(dxDReportJs.Properties().OrderBy(p => (string?)p.Name));
            //var sortReportJsBL = new JObject(dxDReportJsBL.Properties().OrderBy(p => (string?)p.Name));
            //Assert.AreEqual(sortReportJs, sortReportJsBL);
            #endregion
        }
        #endregion
    }
}
