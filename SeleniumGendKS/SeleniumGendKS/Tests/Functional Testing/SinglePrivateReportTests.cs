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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;

// Notes: Please turn on your VPN and make sure update latest msalIdtoken (at BaseFunctionTest) before running API Testing

namespace SeleniumGendKS.Tests.Functional_Testing
{
    [TestFixture]
    internal class SinglePrivateReportTests : BaseFunctionTest
    {
        #region Initiate variables
        internal static string workbenchApi = xdoc.XPathSelectElement("config/webApis").Attribute("WorkbenchApi").Value;
        internal static string filePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"FADAddInApi\Documents\");
        internal static string? fileNameDSourceManual = null;
        internal static string? fileNameDSourceCambridge = null;
        #endregion

        #region TestMethod
        [Test, Category("API Smoke Tests")]
        public void ST001_SinglePrivateReportSourceManual()
        {
            #region Variables declare
            const string custom_risk_benchmarks = "[{"
                                                  + "\n" + "\"benchmark_id\"" + " : " + "\"" + "SPTR Index.USD" + "\","
                                                  + "\n" + "\"beta\"" + " : " + "1,"
                                                  + "\n" + "\"exposure\"" + " : " + "0.8,"
                                                  + "\n" + "\"return_type\"" + " : " + "\"Gross\"" + "\n" +
                                                  "}]";
            const string effective_date = "2021-06-30";
            const string ca_asset_class = "Buyout";
            const string ca_geo = "Africa";
            const string manager = "VGO Capital Partners";
            const string data_source = "Manual";
            var body = "{" + "\n" + "\"custom_risk_benchmarks\"" + " : " + custom_risk_benchmarks + ","
                           + "\n" + "\"effective_date\"" + " : " + "\"" + effective_date + "\","
                           + "\n" + "\"ca_asset_class\"" + " : " + "\"" + ca_asset_class + "\","
                           + "\n" + "\"ca_geo\"" + " : " + "\"" + ca_geo + "\","
                           + "\n" + "\"manager\"" + " : " + "\"" + manager + "\","
                           + "\n" + "\"data_source\"" + " : " + "\"" + data_source + "\"" + "\n" +
                       "}";
            #endregion

            #region Check if api of Sandbox or Staging then get data (on that site)
            if (workbenchApi.Contains("sandbox"))
            {
                fileNameDSourceManual = "SinglePrivateReportManagerOutput.json";
            }
            if (workbenchApi.Contains("conceptia"))
            {
                fileNameDSourceManual = "SinglePrivateReportManagerStagingOutput.json";
            }
            #endregion

            #region Run Tests
            // Get Single Private Report
            var singlePrivateReport = WorkbenchApi.SinglePrivateReport(body, msalIdtoken);
            Assert.That(singlePrivateReport.StatusCode, Is.EqualTo(HttpStatusCode.OK));

            // Parse IRestResponse to JObject
            JObject singlePrivateReportJs = JObject.Parse(singlePrivateReport.Content);
            ClassicAssert.AreEqual("200", singlePrivateReportJs["statusCode"].ToString());
            ClassicAssert.AreEqual("application/json", singlePrivateReportJs["headers"]["content-type"].ToString());
            ClassicAssert.AreEqual(2, singlePrivateReportJs["body"]["base"].Count());
            //JObject singlePrivateReportJsBL = JObject.Parse(File.ReadAllText(filePath + fileNameDSourceManual));
            //singlePrivateReportJsBL["headers"]["date"] = singlePrivateReportJs["headers"]["date"];
            //singlePrivateReportJsBL.SelectToken("headers.version")?.Parent.Remove();
            //singlePrivateReportJs.SelectToken("headers.version")?.Parent.Remove();
            //var sortReportJs = new JObject(singlePrivateReportJs.Properties().OrderBy(p => (string?)p.Name));
            //var sortReportJsBL = new JObject(singlePrivateReportJsBL.Properties().OrderBy(p => (string?)p.Name));
            //Assert.AreEqual(sortReportJs, sortReportJsBL);
            #endregion
        }

        [Test, Category("API Smoke Tests")]
        public void ST002_SinglePrivateReportSourceCambridge()
        {
            #region Variables declare
            const string custom_risk_benchmarks = "[{"
                                                  + "\n" + "\"benchmark_id\"" + " : " + "\"" + "SPTR Index.USD" + "\","
                                                  + "\n" + "\"beta\"" + " : " + "1,"
                                                  + "\n" + "\"exposure\"" + " : " + "0.8,"
                                                  + "\n" + "\"return_type\"" + " : " + "\"Gross\"" + "\n" +
                                                  "}]";
            const string effective_date = "2021-06-30";
            const string ca_asset_class = "Venture";
            const string ca_geo = "United States";
            const string manager = "GSR Ventures";
            const string data_source = "Cambridge";
            var body = "{" + "\n" + "\"custom_risk_benchmarks\"" + " : " + custom_risk_benchmarks + ","
                           + "\n" + "\"effective_date\"" + " : " + "\"" + effective_date + "\","
                           + "\n" + "\"ca_asset_class\"" + " : " + "\"" + ca_asset_class + "\","
                           + "\n" + "\"ca_geo\"" + " : " + "\"" + ca_geo + "\","
                           + "\n" + "\"manager\"" + " : " + "\"" + manager + "\","
                           + "\n" + "\"data_source\"" + " : " + "\"" + data_source + "\"" + "\n" +
                       "}";
            #endregion

            #region Check if api of Sandbox or Staging then get data (on that site)
            if (workbenchApi.Contains("sandbox"))
            {
                fileNameDSourceCambridge = "SinglePrivateReportCambridgeOutput.json";
            }
            if (workbenchApi.Contains("conceptia"))
            {
                fileNameDSourceCambridge = "SinglePrivateReportCambridgeStagingOutput.json";
            }
            #endregion

            #region Run Tests
            // Get Single Private Report
            var singlePrivateReport = WorkbenchApi.SinglePrivateReport(body, msalIdtoken);
            Assert.That(singlePrivateReport.StatusCode, Is.EqualTo(HttpStatusCode.OK));

            // Parse IRestResponse to JObject
            JObject singlePrivateReportJs = JObject.Parse(singlePrivateReport.Content);
            ClassicAssert.AreEqual("200", singlePrivateReportJs["statusCode"].ToString());
            ClassicAssert.AreEqual("application/json", singlePrivateReportJs["headers"]["content-type"].ToString());
            ClassicAssert.AreEqual(1, singlePrivateReportJs["body"]["base"].Count());
            //JObject singlePrivateReportJsBL = JObject.Parse(File.ReadAllText(filePath + fileNameDSourceCambridge));
            //singlePrivateReportJsBL["headers"]["date"] = singlePrivateReportJs["headers"]["date"];
            //singlePrivateReportJsBL.SelectToken("headers.version")?.Parent.Remove();
            //singlePrivateReportJs.SelectToken("headers.version")?.Parent.Remove();
            //var sortReportJs = new JObject(singlePrivateReportJs.Properties().OrderBy(p => (string?)p.Name));
            //var sortReportJsBL = new JObject(singlePrivateReportJsBL.Properties().OrderBy(p => (string?)p.Name));
            //Assert.AreEqual(sortReportJs, sortReportJsBL);
            #endregion
        }
        #endregion
    }
}
