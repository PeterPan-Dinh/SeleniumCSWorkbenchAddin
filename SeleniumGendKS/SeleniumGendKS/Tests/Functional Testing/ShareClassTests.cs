using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using SeleniumGendKS.Core.DBConnection;
using SeleniumGendKS.FADAddInApi;
using SeleniumGendKS.FADAddInApi.PredefinedScenarios;
using System.Net;
using System.Reflection;

// Notes: Please turn on your VPN and make sure update latest msalIdtoken (at BaseFunctionTest) before running API Testing

namespace SeleniumGendKS.Tests.Functional_Testing
{
    [TestFixture]
    internal class ShareClassTests : BaseFunctionTest
    {
        #region Initiate variables
        internal static int fundId = 14590;
        #endregion

        #region TestMethod
        [Test, Category("API Smoke Tests")]
        public void ST001_GetShareClassByFundId()
        {
            // Variables declare
            string filePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"FADAddInApi\Documents\");
            string fileName = "GetShareClassOutput.json";

            // Get Share Class By Fund Id
            var shareClass = WorkbenchApi.GetShareClassByFundId(fundId, msalIdtoken);
            Assert.That(shareClass.StatusCode, Is.EqualTo(HttpStatusCode.OK));

            // Parse IRestResponse to List JObject
            List<JObject> shareClassJs = JsonConvert.DeserializeObject<List<JObject>>(shareClass.Content);
            List<JObject> shareClassJsBL = JsonConvert.DeserializeObject<List<JObject>>(File.ReadAllText(filePath + fileName));
            DatabaseConnection.RemoveFieldNameInJObject(shareClassJs, "_time_");
            DatabaseConnection.RemoveFieldNameInJObject(shareClassJsBL, "_time_");
            ClassicAssert.AreEqual(shareClassJsBL.Count, shareClassJs.Count);
            ClassicAssert.AreEqual(shareClassJsBL, shareClassJs);
        }
        #endregion
    }
}
