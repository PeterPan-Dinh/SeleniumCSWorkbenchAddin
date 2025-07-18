using AventStack.ExtentReports.Gherkin.Model;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using SeleniumGendKS.FADAddInApi;
using SeleniumGendKS.FADAddInApi.PredefinedScenarios;
using SeleniumGendKS.FADAddInApi.UIStepsDataManagement;
using System.Net;
using System.Reflection;
using System.Security.Policy;
using System.Xml.XPath;

// Notes: Please turn on your VPN and make sure update latest msalIdtoken (at BaseFunctionTest) before running API Testing

namespace SeleniumGendKS.Tests.Functional_Testing
{
    [TestFixture]
    internal class SearchFundTests : BaseFunctionTest
    {
        #region Initiate variables
        internal static string workbenchApi = xdoc.XPathSelectElement("config/webApis").Attribute("WorkbenchApi").Value;
        internal static string? managerId;
        internal static int fundId;
        internal static string? dataSource;
        #endregion

        #region TestMethod
        [Test, Category("API Smoke Tests")]
        public void ST001_SearchFund()
        {
            // Variables declare
            const string isOwnedByKS = "false";
            const string search_text = "Citadel Multi Strategy Funds";

            // Search Fund
            var body = "{" + "\n" + "\"isOwnedByKS\"" + " : " + isOwnedByKS + ","
                           + "\n" + "\"search_text\"" + " : " + "\"" + search_text + "\"" + "\n" +
                       "}";
            var searchFund = WorkbenchApi.SearchFund(body, msalIdtoken);
            Assert.That(searchFund.StatusCode, Is.EqualTo(HttpStatusCode.OK));

            // Parse IRestResponse to List JObject
            List<JObject> searchFundJs = JsonConvert.DeserializeObject<List<JObject>>(searchFund.Content);
            Assert.That(searchFundJs.Count, Is.GreaterThanOrEqualTo(1));
            ClassicAssert.AreEqual("14590", searchFundJs[0]["fund_id"].ToString());
            ClassicAssert.AreEqual(search_text, searchFundJs[0]["fund_name"].ToString());
            ClassicAssert.AreEqual("357", searchFundJs[0]["manager_id"].ToString());
            ClassicAssert.AreEqual("Citadel Advisors LLC", searchFundJs[0]["manager_name"].ToString());
            ClassicAssert.AreEqual("ALB", searchFundJs[0]["source"].ToString());
        }

        [Test, Category("API Smoke Tests")]
        public void ST002_SearchFaceset()
        {
            // Variables declare
            string facesetInput = "S&P 500 Index"; // --> search "SP50" shall get 2 facesets

            // Search Faceset
            var searchFaceset = WorkbenchApi.SearchFaceset(facesetInput, msalIdtoken);
            Assert.That(searchFaceset.StatusCode, Is.EqualTo(HttpStatusCode.OK));

            // Parse IRestResponse to List JObject
            List<JObject> searchFacesetJs = JsonConvert.DeserializeObject<List<JObject>>(searchFaceset.Content);

            // Check if api of Sandbox or Staging then get data (on that site)
            int? count = null;
            if (workbenchApi.Contains("sandbox")) { count = 1; }
            if (workbenchApi.Contains("conceptia")) { count = 1; }
            ClassicAssert.AreEqual(count, searchFacesetJs.Count);

            // Sort Properties and values
            var FacesetJsSortKey = WorkbenchApi.SortPropertiesByName(searchFacesetJs);
            var facesetJsSortValue = FacesetJsSortKey.OrderBy(p => (string?)p["name"]).ToList();
            ClassicAssert.AreEqual("S&P 500 Index", facesetJsSortValue[0]["name"].ToString());
        }

        [Test, Category("API Smoke Tests")]
        public void ST003_GetFundDataALBById()
        {
            // Variables declare
            fundId = 14590;
            dataSource = "ALB";

            // Get Fund Data by Id (source=ALB)
            var fundData = WorkbenchApi.GetFundDataById(fundId, dataSource, msalIdtoken);
            Assert.That(fundData.StatusCode, Is.EqualTo(HttpStatusCode.OK));

            // Parse IRestResponse to JObject
            JObject fundDataJs = JObject.Parse(fundData.Content);

            // Check if api of Sandbox or Staging then get data (on that site)
            int? count = null;
            if (workbenchApi.Contains("sandbox")) { count = 166; }
            if (workbenchApi.Contains("conceptia")) { count = 166; }
            ClassicAssert.AreEqual(count, fundDataJs.Count);
            ClassicAssert.AreEqual(fundId.ToString(), fundDataJs["fund_id"].ToString());
            ClassicAssert.AreEqual("ALBOURNE", fundDataJs["data_source"].ToString());
            ClassicAssert.AreEqual("ALB", fundDataJs["source"].ToString());
            ClassicAssert.AreEqual("False", fundDataJs["data_catch_up"].ToString());
        }

        [Test, Category("API Smoke Tests")]
        public void ST004_GetFundStatusALBById()
        {
            // Variables declare
            fundId = 14590;
            dataSource = "ALB";

            // Get Fund Status by Id (source=ALB)
            var fundStatus = WorkbenchApi.GetFundStatusById(fundId, dataSource, msalIdtoken);
            Assert.That(fundStatus.StatusCode, Is.EqualTo(HttpStatusCode.OK));

            // Parse IRestResponse to JObject
            JObject fundStatusJs = JObject.Parse(fundStatus.Content);
            // Check if api of Sandbox or Staging then get data (on that site)
            string? first_aum_date = null;
            if (workbenchApi.Contains("sandbox")) { first_aum_date = "1998-01-31 00:00:00.000"; }
            if (workbenchApi.Contains("conceptia")) { first_aum_date = ""; }
            ClassicAssert.AreEqual(2, fundStatusJs.Count); // Update base on KS-517
            ClassicAssert.AreEqual(fundId.ToString(), fundStatusJs["manual"]["fund_id"].ToString());
            ClassicAssert.AreEqual(fundId.ToString(), fundStatusJs["ALB"]["fund_id"].ToString());
            ClassicAssert.AreEqual(first_aum_date, fundStatusJs["manual"]["first_aum_date"].ToString());
            ClassicAssert.AreEqual("1995-07-31 00:00:00.000", fundStatusJs["manual"]["first_ror_date"].ToString());
            ClassicAssert.AreEqual(dataSource, fundStatusJs["manual"]["fund_source"].ToString());
            ClassicAssert.AreEqual("manual", fundStatusJs["manual"]["data_source"].ToString());
            ClassicAssert.AreEqual("ALBOURNE", fundStatusJs["ALB"]["data_source"].ToString());
        }

        [Test, Category("API Smoke Tests")]
        public void ST005_GetFundDataPrivateById()
        {
            // Variables declare
            managerId = "VGO%20Capital%20Partners";

            // Get Fund Data by Id (source=ALB)
            var fundData = WorkbenchApi.GetFundDataPrivateById(managerId, msalIdtoken);
            Assert.That(fundData.StatusCode, Is.EqualTo(HttpStatusCode.OK));

            // Parse IRestResponse to List JObject
            List<JObject> fundDataJs = JsonConvert.DeserializeObject<List<JObject>>(fundData.Content);
            ClassicAssert.AreEqual(2, fundDataJs.Count);
            ClassicAssert.AreEqual("VGO Capital Partners", fundDataJs[0]["firm"].ToString());
            ClassicAssert.AreEqual("VGO Capital Partners", fundDataJs[1]["firm"].ToString());
            ClassicAssert.AreEqual("European Mezzanine", fundDataJs[0]["strategy"].ToString());
            ClassicAssert.AreEqual("European Mezzanine", fundDataJs[1]["strategy"].ToString());
            ClassicAssert.AreEqual("VGO Special Situations Fund I", fundDataJs[0]["fund_name"].ToString());
            ClassicAssert.AreEqual("VGO Special Situations Fund II", fundDataJs[1]["fund_name"].ToString());
            ClassicAssert.AreEqual("", fundDataJs[0]["fund_size_expected_size_m"].ToString());
            ClassicAssert.AreEqual("288", fundDataJs[1]["fund_size_expected_size_m"].ToString());
        }

        [Test, Category("API Smoke Tests")]
        public void ST006_GetFundStatusPrivateById()
        {
            // Variables declare
            managerId = "VGO%20Capital%20Partners";
            dataSource = "cambridge"; // old: ALB

            // Get Fund Status by Id (source=cambridge)
            var fundStatus = WorkbenchApi.GetFundStatusPrivateById(managerId, dataSource, msalIdtoken);
            Assert.That(fundStatus.StatusCode, Is.EqualTo(HttpStatusCode.OK));

            // Parse IRestResponse to JObject
            JObject fundStatusJs = JObject.Parse(fundStatus.Content);
            // Check if api of Sandbox or Staging then get data (on that site)
            string? fundInforTotalRecordsCount = null;
            if (workbenchApi.Contains("sandbox")) { fundInforTotalRecordsCount = "15"; }
            if (workbenchApi.Contains("conceptia")) { fundInforTotalRecordsCount = "2"; }
            ClassicAssert.AreEqual(2, fundStatusJs.Count);
            ClassicAssert.AreEqual("VGO Capital Partners", fundStatusJs["dealInfor"]["manager_id"].ToString());
            ClassicAssert.AreEqual("23", fundStatusJs["dealInfor"]["total_records"].ToString());
            ClassicAssert.AreEqual("VGO Capital Partners", fundStatusJs["fundInfor"]["manager_id"].ToString());
            ClassicAssert.AreEqual(fundInforTotalRecordsCount, fundStatusJs["fundInfor"]["total_records"].ToString());
        }

        [Test, Category("API Smoke Tests")]
        public void ST007_GetFundDataSolovisById()
        {
            // Variables declare
            fundId = 565;
            dataSource = "solovis";

            // Get Fund Data by Id (source=Solovis)
            var fundData = WorkbenchApi.GetFundDataById(fundId, dataSource, msalIdtoken);
            Assert.That(fundData.StatusCode, Is.EqualTo(HttpStatusCode.OK));

            // Parse IRestResponse to JObject
            JObject fundDataJs = JObject.Parse(fundData.Content);
            ClassicAssert.AreEqual(47, fundDataJs.Count);
            ClassicAssert.AreEqual(fundId.ToString(), fundDataJs["fund_id"].ToString());
            ClassicAssert.AreEqual("Laurion Capital Ltd.", fundDataJs["fund_name"].ToString());
            ClassicAssert.AreEqual("Absolute Return", fundDataJs["asset_class_0"].ToString());
            //ClassicAssert.AreEqual("Relative Value", fundDataJs["asset_class_1"].ToString());
            ClassicAssert.AreEqual("Relative Value", fundDataJs["asset_class_2"].ToString());
            ClassicAssert.AreEqual("Relative Value", fundDataJs["asset_class_3"].ToString());
            ClassicAssert.AreEqual("Laurion Capital Management LP", fundDataJs["manager_name"].ToString());
            ClassicAssert.AreEqual("", fundDataJs["manager_sub_strategy"].ToString());
            ClassicAssert.AreEqual("", fundDataJs["data_fund_status"].ToString());
            ClassicAssert.AreEqual("2020-12-29", fundDataJs["ks_date"].ToString());
            ClassicAssert.AreEqual("2005-09-30", fundDataJs["data_inception_date"].ToString());
            ClassicAssert.AreEqual("2005-09-30", fundDataJs["ror_start_date"].ToString());
            ClassicAssert.AreEqual("0.12", fundDataJs["tracking_error"].ToString());
            ClassicAssert.AreEqual("Accounting Fund", fundDataJs["data_fund_type"].ToString());
            ClassicAssert.AreEqual("KS", fundDataJs["fund_owner"].ToString());
            ClassicAssert.AreEqual("Absolute Return", fundDataJs["strategy"].ToString());
        }

        [Test, Category("API Smoke Tests")]
        public void ST008_GetFundStatusSolovisById()
        {
            // Variables declare
            fundId = 565;
            dataSource = "solovis";

            // Get Fund Status by Id (source=Solovis)
            var fundStatus = WorkbenchApi.GetFundStatusById(fundId, dataSource, msalIdtoken);
            Assert.That(fundStatus.StatusCode, Is.EqualTo(HttpStatusCode.OK));

            // Parse IRestResponse to JObject
            JObject fundStatusJs = JObject.Parse(fundStatus.Content);
            int? count = null;
            if (workbenchApi.Contains("sandbox")) { count = 2; }
            if (workbenchApi.Contains("conceptia")) { count = 1; }
            ClassicAssert.AreEqual(count, fundStatusJs.Count); // Update base on KS-517
            ClassicAssert.AreEqual(fundId.ToString(), fundStatusJs[dataSource]["fund_id"].ToString());
        }

        [Test, Category("API Smoke Tests")]
        public void ST009_GetFundDataAlternativesEvestmentById()
        {
            // Variables declare
            fundId = 632785;
            dataSource = "aevest";

            // Get Fund Data by Id (source=Evestment)
            var fundData = WorkbenchApi.GetFundDataById(fundId, dataSource, msalIdtoken);
            Assert.That(fundData.StatusCode, Is.EqualTo(HttpStatusCode.OK));

            // Parse IRestResponse to JObject
            JObject fundDataJs = JObject.Parse(fundData.Content);
            //Assert.AreEqual(26, fundDataJs.Count);
            ClassicAssert.AreEqual(fundId.ToString(), fundDataJs["fund_id"].ToString());
            ClassicAssert.AreEqual("Advent Global Partners", fundDataJs["fund_name"].ToString());
            ClassicAssert.AreEqual("Advent Capital Management, LLC", fundDataJs["manager_name"].ToString());
            ClassicAssert.AreEqual("2001-08-01", fundDataJs["data_inception_date"].ToString());
            ClassicAssert.AreEqual("aevest", fundDataJs["source"].ToString());
            ClassicAssert.AreEqual("evestment", fundDataJs["data_source"].ToString());
        }

        [Test, Category("API Smoke Tests")]
        public void ST010_GetFundStatusAlternativesEvestmentById()
        {
            // Variables declare
            fundId = 632785;
            dataSource = "aevest";

            // Get Fund Status by Id (source=Evestment)
            var fundStatus = WorkbenchApi.GetFundStatusById(fundId, dataSource, msalIdtoken);
            Assert.That(fundStatus.StatusCode, Is.EqualTo(HttpStatusCode.OK));

            // Parse IRestResponse to JObject
            JObject fundStatusJs = JObject.Parse(fundStatus.Content);
            ClassicAssert.AreEqual(1, fundStatusJs.Count); // Update base on KS-517
            ClassicAssert.AreEqual(fundId.ToString(), fundStatusJs[dataSource]["productid"].ToString());
        }

        [Test, Category("API Smoke Tests")]
        public void ST011_GetFundDataTraditionalEvestmentById()
        {
            // Variables declare
            fundId = 660843; // John...2035
            dataSource = "evest";

            // Get Fund Data by Id (source=Evestment)
            var fundData = WorkbenchApi.GetFundDataById(fundId, dataSource, msalIdtoken);
            Assert.That(fundData.StatusCode, Is.EqualTo(HttpStatusCode.OK));

            // Parse IRestResponse to JObject
            JObject fundDataJs = JObject.Parse(fundData.Content);
            //ClassicAssert.AreEqual(23, fundDataJs.Count);
            ClassicAssert.AreEqual(fundId.ToString(), fundDataJs["fund_id"].ToString());
            ClassicAssert.AreEqual("John Hancock Multimanager Lifetime Portfolios 2035", fundDataJs["fund_name"].ToString());
            ClassicAssert.AreEqual("John Hancock Investments", fundDataJs["manager_name"].ToString());
            ClassicAssert.AreEqual("2006-10-30", fundDataJs["data_inception_date"].ToString());
            ClassicAssert.AreEqual("evest", fundDataJs["source"].ToString());
            ClassicAssert.AreEqual("evestment", fundDataJs["data_source"].ToString());
        }

        [Test, Category("API Smoke Tests")]
        public void ST012_GetFundStatusTraditionalEvestmentById()
        {
            // Variables declare
            fundId = 660843; // John...2035
            dataSource = "evest";

            // Get Fund Status by Id (source=Evestment)
            var fundStatus = WorkbenchApi.GetFundStatusById(fundId, dataSource, msalIdtoken);
            Assert.That(fundStatus.StatusCode, Is.EqualTo(HttpStatusCode.OK));

            // Parse IRestResponse to JObject
            JObject fundStatusJs = JObject.Parse(fundStatus.Content);
            ClassicAssert.AreEqual(1, fundStatusJs.Count); // Update base on KS-517
            ClassicAssert.AreEqual(fundId.ToString(), fundStatusJs[dataSource]["productid"].ToString());
        }
        #endregion
    }
}
