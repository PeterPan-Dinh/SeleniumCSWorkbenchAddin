using System.Net;
using System.Reflection;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using SeleniumGendKS.FADAddInApi;
using SeleniumGendKS.FADAddInApi.PredefinedScenarios;

// Notes: Please turn on your VPN and make sure update latest msalIdtoken (at BaseFunctionTest) before running API Testing

namespace SeleniumGendKS.Tests.Functional_Testing
{
    [TestFixture]
    internal class FileMappingTests : BaseFunctionTest
    {
        #region Initiate variables
        internal static string fileMappingName = "QA file";
        internal static string dataType = "fund_aum";
        internal static string fileMappingNameUpdate = "QA file1";
        internal static string dataTypeUpdate = "fund_returns";
        internal static int fundId = 14590;
        internal static string fileMappingById = CreateFileMapping().GetValue("_id").ToString();
        internal static JObject CreateFileMapping()
        {
            // Send request
            var fileMapping = WorkbenchApi.CreateFileMapping(fileMappingName, dataType, fundId, msalIdtoken);
            ClassicAssert.IsNotNull(fileMapping.Content);

            // parse IRestResponse to JObject
            return JObject.Parse(fileMapping.Content);
        }
        internal static JObject GetFileMapping(string? fileMappingById=null)
        {
            // Send request
            var fileMapping = WorkbenchApi.GetFileMappingById(fileMappingById, msalIdtoken);
            ClassicAssert.IsNotNull(fileMapping);

            // parse IRestResponse to JObject
            return JObject.Parse(fileMapping.Content);
        }
        internal static JObject UpdateFileMapping(string? fileMappingById = null)
        {
            // Send request
            var fileMapping = WorkbenchApi.UpdateFileMappingById(fileMappingById);
            ClassicAssert.IsNotNull(fileMapping);

            // parse IRestResponse to JObject
            return JObject.Parse(fileMapping.Content);
        }
        internal static JObject DeleteFileMapping(string? fileMappingById = null)
        {
            // Send request
            var fileMapping = WorkbenchApi.DeleteFileMappingById(fileMappingById, msalIdtoken);
            ClassicAssert.IsNotNull(fileMapping);

            // parse IRestResponse to JObject
            return JObject.Parse(fileMapping.Content);
        }
        #endregion

        #region TestMethod
        [Test, Category("API Smoke Tests")]
        public void ST001_CreateFileMappingById()
        {
            // Variables declare
            string filePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"FADAddInApi\Documents\");
            string fileName = "CreateFileMappingOutput.json";
            
            // Create File Mapping By Id
            var fileMapping = WorkbenchApi.CreateFileMapping(fileMappingName, dataType, fundId, msalIdtoken);
            Assert.That(fileMapping.StatusCode, Is.EqualTo(HttpStatusCode.Created));

            // Parse IRestResponse to JObject
            JObject fileMappingJs = JObject.Parse(fileMapping.Content);
            JObject fileMappingJsBL = JObject.Parse(File.ReadAllText(filePath + fileName));
            fileMappingJsBL["_id"] = fileMappingJs.GetValue("_id");
            fileMappingJsBL["field_mappings"][0]["_id"] = fileMappingJs["field_mappings"][0]["_id"];
            fileMappingJsBL["as_of_date"] = fileMappingJs["as_of_date"];
            fileMappingJs.Property("created_at").Remove();
            fileMappingJs.Property("updatedAt").Remove();
            fileMappingJsBL.Property("created_at").Remove();
            fileMappingJsBL.Property("updatedAt").Remove();
            ClassicAssert.IsTrue(JToken.DeepEquals(fileMappingJs, fileMappingJsBL));

            // Get File Mapping By Id
            GetFileMapping(fileMappingJs.GetValue("_id").ToString());

            // Delete File Mapping By Id
            DeleteFileMapping(fileMappingJs.GetValue("_id").ToString());
        }

        [Test, Category("API Smoke Tests")]
        public void ST002_GetFileMappingById()
        {
            // Variables declare
            string filePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"FADAddInApi\Documents\");
            string fileName = "GetFileMappingOutput.json";

            // Get File Mapping By Id
            var fileMapping = WorkbenchApi.GetFileMappingById(fileMappingById, msalIdtoken);
            Assert.That(fileMapping.StatusCode, Is.EqualTo(HttpStatusCode.OK));

            // Parse IRestResponse to JObject
            JObject fileMappingJs = JObject.Parse(fileMapping.Content);
            JObject fileMappingJsBL = JObject.Parse(File.ReadAllText(filePath + fileName));
            fileMappingJsBL["_id"] = fileMappingJs.GetValue("_id");
            fileMappingJsBL["field_mappings"][0]["_id"] = fileMappingJs["field_mappings"][0]["_id"];
            fileMappingJsBL["as_of_date"] = fileMappingJs["as_of_date"];
            fileMappingJs.Property("created_at").Remove();
            fileMappingJs.Property("updatedAt").Remove();
            fileMappingJsBL.Property("created_at").Remove();
            fileMappingJsBL.Property("updatedAt").Remove();
            ClassicAssert.IsTrue(JToken.DeepEquals(fileMappingJs, fileMappingJsBL));

            // Delete File Mapping By Id
            DeleteFileMapping(fileMappingJs.GetValue("_id").ToString());
        }

        [Test, Category("API Smoke Tests")]
        public void ST003_UpdateFileMappingById()
        {
            // Variables declare
            string filePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"FADAddInApi\Documents\");
            string fileName = "UpdateFileMappingOutput.json";

            // Create File Mapping By Id
            var fileMapping = WorkbenchApi.CreateFileMapping(fileMappingName, dataType, fundId, msalIdtoken);
            Assert.That(fileMapping.StatusCode, Is.EqualTo(HttpStatusCode.Created));

            // Update File Mapping
            var fileMappingJs = JObject.Parse(fileMapping.Content); // --> parse IRestResponse to JObject
            var fileMappingUpdate = WorkbenchApi.UpdateFileMappingById(fileMappingJs.GetValue("_id").ToString(), fileMappingNameUpdate, dataTypeUpdate, fundId, msalIdtoken);
            Assert.That(fileMappingUpdate.StatusCode, Is.EqualTo(HttpStatusCode.OK));

            // Parse IRestResponse to JObject
            JObject fileMappingJsUpdate = JObject.Parse(fileMappingUpdate.Content);
            JObject fileMappingJsBL = JObject.Parse(File.ReadAllText(filePath + fileName));
            fileMappingJsBL["_id"] = fileMappingJsUpdate.GetValue("_id");
            fileMappingJsBL["field_mappings"][0]["_id"] = fileMappingJsUpdate["field_mappings"][0]["_id"];
            fileMappingJsBL["as_of_date"] = fileMappingJsUpdate["as_of_date"];
            fileMappingJsUpdate.Property("created_at").Remove();
            fileMappingJsUpdate.Property("updatedAt").Remove();
            fileMappingJsBL.Property("created_at").Remove();
            fileMappingJsBL.Property("updatedAt").Remove();
            ClassicAssert.IsTrue(JToken.DeepEquals(fileMappingJsUpdate, fileMappingJsBL));

            // Get (the updated) File Mapping By Id
            GetFileMapping(fileMappingJsUpdate.GetValue("_id").ToString());

            // Delete (the updated) File Mapping By Id
            DeleteFileMapping(fileMappingJsUpdate.GetValue("_id").ToString());
        }

        [Test, Category("API Smoke Tests")]
        public void ST004_DeleteFileMappingById()
        {
            // Create File Mapping By Id
            var fileMapping = WorkbenchApi.CreateFileMapping(fileMappingName, dataType, fundId, msalIdtoken);
            Assert.That(fileMapping.StatusCode, Is.EqualTo(HttpStatusCode.Created));

            // Delete File Mapping By Id (Send request)
            var fileMappingJsDelete = JObject.Parse(fileMapping.Content); // --> parse IRestResponse to JObject
            var fileMappingDelete = WorkbenchApi.DeleteFileMappingById(fileMappingJsDelete.GetValue("_id").ToString(), msalIdtoken);
            Assert.That(fileMappingDelete.StatusCode, Is.EqualTo(HttpStatusCode.OK));

            // parse IRestResponse to JObject (to get the deleted file mapping)
            var deletetedfileMapping = JObject.Parse(fileMappingDelete.Content);
            ClassicAssert.AreEqual(2, deletetedfileMapping["rowsCount"].Count());
            ClassicAssert.AreEqual("True", deletetedfileMapping["rowsCount"]["acknowledged"].ToString());
            ClassicAssert.AreEqual("1", deletetedfileMapping["rowsCount"]["deletedCount"].ToString());
        }
        #endregion
    }
}
