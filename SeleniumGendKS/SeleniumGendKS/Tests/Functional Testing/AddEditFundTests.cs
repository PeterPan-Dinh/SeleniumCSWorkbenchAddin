using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using SeleniumGendKS.FADAddInApi;
using SeleniumGendKS.FADAddInApi.PredefinedScenarios;
using System.Net;
using System.Reflection;
using System.Xml.XPath;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;

// Notes: Please turn on your VPN and make sure update latest msalIdtoken (at BaseFunctionTest) before running API Testing

namespace SeleniumGendKS.Tests.Functional_Testing
{
    [TestFixture]
    internal class AddEditFundTests : BaseFunctionTest
    {
        #region Initiate variables
        private JObject? addPublicFundManagerJs;
        internal static string workbenchApi = xdoc.XPathSelectElement("config/webApis").Attribute("WorkbenchApi").Value;
        #endregion

        #region TestMethod
        [Test, Category("API Smoke Tests")]
        public void ST001_AddPublicFund()
        {
            // Variables declare
            string manager_name = "QA_ApiAuto_Manager" + @"_" + DateTime.Now.ToString().Replace("/", "_").Replace(":", "_").Replace(" ", "_");
            var body = "{" + "\n" + "\"name\"" + " : " + "\"" + manager_name + "\"" + "\n" +
                       "}";

            // Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
            if (workbenchApi.Contains("sandbox"))
            {
                #region Add Public Fund - Manager
                // Add Public Fund Manager
                var addPublicFundManager = WorkbenchApi.AddManagerApi(body, msalIdtoken);
                Assert.That(addPublicFundManager.StatusCode, Is.EqualTo(HttpStatusCode.Created));

                // Parse IRestResponse to JObject
                addPublicFundManagerJs = JObject.Parse(addPublicFundManager.Content);
                ClassicAssert.AreEqual(manager_name, addPublicFundManagerJs["name"].ToString());
                ClassicAssert.AreEqual("manual", addPublicFundManagerJs["source"].ToString());
                #endregion

                #region Add Public Fund - Fund
                // Add Public Fund (Child)
                string fund_name = "QA_ApiAuto_Fund" + @"_" + DateTime.Now.ToString().Replace("/", "_").Replace(":", "_").Replace(" ", "_");
                string sub_asset_class = "US Growth Equity";
                string asset_class = "Private Equity";
                string date = DateTime.Now.Date.ToString("yyyy-MM-dd"), inception_date = date.Replace("/", "-");
                string latest_actual_value = "QA api auto Latest Actual Value";
                string business_street = "QA api auto Business Street";
                string business_city = "QA api auto Business City";
                string business_state = "QA api auto Business State";
                string business_zip = "QA api auto Business ZIP";
                string business_country = "JP";
                string data_redemption_frequency = "Monthly";
                string data_redemption_notice_days = "59 Days";
                string data_hard_lockup = "Yes";
                string data_hard_lockup_months = "35.01";
                string data_soft_lockup_months = "14.01";
                string data_redemption_fee = "16.01";
                string data_redemption_gate = "10.01";
                string data_redemption_gate_percent = "8.51";
                string data_percent_of_nav_available = "26.51";
                string data_side_pocket = "Yes";
                string data_redemption_note = "QA Api auto test notes Liq 111";
                string data_management_fee = "-10.51";
                string data_management_fee_frequency = "Monthly";
                string data_performance_fee = "-6.51";
                string data_hwm_status = "Yes";
                string data_catch_up = "Yes";
                string data_catch_up_rate = "-17.51";
                string data_crystalization_frequency = "2";
                string data_hurdle_status = "Yes";
                string data_hurdle_type = "Fixed";
                string data_benchmark = "MS133333.USD";
                string data_benchmark_name = "MSCI CHINA A ONSHORE Net Return Index USD";
                string data_hurdle_rate = "-5.51";
                string data_hurdle_hard_soft = "Hard";
                string data_hurdle_ramp_type = "Performance Dependent";
                int manager_id = int.Parse(addPublicFundManagerJs["id"].ToString());
                string manager_source = "manual";
                body = "{"
                        + "\n" + "\"manager_name\"" + " : " + "\"" + manager_name + "\","
                        + "\n" + "\"fund_name\"" + " : " + "\"" + fund_name + "\","
                        + "\n" + "\"sub_asset_class\"" + " : " + "\"" + sub_asset_class + "\","
                        + "\n" + "\"asset_class\"" + " : " + "\"" + asset_class + "\","
                        + "\n" + "\"inception_date\"" + " : " + "\"" + inception_date + "\","
                        + "\n" + "\"latest_actual_value\"" + " : " + "\"" + latest_actual_value + "\","
                        + "\n" + "\"business_street\"" + " : " + "\"" + business_street + "\","
                        + "\n" + "\"business_city\"" + " : " + "\"" + business_city + "\","
                        + "\n" + "\"business_state\"" + " : " + "\"" + business_state + "\","
                        + "\n" + "\"business_zip\"" + " : " + "\"" + business_zip + "\","
                        + "\n" + "\"business_country\"" + " : " + "\"" + business_country + "\","
                        + "\n" + "\"data_redemption_frequency\"" + " : " + "\"" + data_redemption_frequency + "\","
                        + "\n" + "\"data_redemption_notice_days\"" + " : " + "\"" + data_redemption_notice_days + "\","
                        + "\n" + "\"data_hard_lockup\"" + " : " + "\"" + data_hard_lockup + "\","
                        + "\n" + "\"data_hard_lockup_months\"" + " : " + "\"" + data_hard_lockup_months + "\","
                        + "\n" + "\"data_soft_lockup_months\"" + " : " + "\"" + data_soft_lockup_months + "\","
                        + "\n" + "\"data_redemption_fee\"" + " : " + "\"" + data_redemption_fee + "\","
                        + "\n" + "\"data_redemption_gate\"" + " : " + "\"" + data_redemption_gate + "\","
                        + "\n" + "\"data_redemption_gate_percent\"" + " : " + "\"" + data_redemption_gate_percent + "\","
                        + "\n" + "\"data_percent_of_nav_available\"" + " : " + "\"" + data_percent_of_nav_available + "\","
                        + "\n" + "\"data_side_pocket\"" + " : " + "\"" + data_side_pocket + "\","
                        + "\n" + "\"data_redemption_note\"" + " : " + "\"" + data_redemption_note + "\","
                        + "\n" + "\"data_management_fee\"" + " : " + "\"" + data_management_fee + "\","
                        + "\n" + "\"data_management_fee_frequency\"" + " : " + "\"" + data_management_fee_frequency + "\","
                        + "\n" + "\"data_performance_fee\"" + " : " + "\"" + data_performance_fee + "\","
                        + "\n" + "\"data_hwm_status\"" + " : " + "\"" + data_hwm_status + "\","
                        + "\n" + "\"data_catch_up\"" + " : " + "\"" + data_catch_up + "\","
                        + "\n" + "\"data_catch_up_rate\"" + " : " + "\"" + data_catch_up_rate + "\","
                        + "\n" + "\"data_crystalization_frequency\"" + " : " + "\"" + data_crystalization_frequency + "\","
                        + "\n" + "\"data_hurdle_status\"" + " : " + "\"" + data_hurdle_status + "\","
                        + "\n" + "\"data_hurdle_type\"" + " : " + "\"" + data_hurdle_type + "\","
                        + "\n" + "\"data_benchmark\"" + " : " + "\"" + data_benchmark + "\","
                        + "\n" + "\"data_benchmark_name\"" + " : " + "\"" + data_benchmark_name + "\","
                        + "\n" + "\"data_hurdle_rate\"" + " : " + "\"" + data_hurdle_rate + "\","
                        + "\n" + "\"data_hurdle_hard_soft\"" + " : " + "\"" + data_hurdle_hard_soft + "\","
                        + "\n" + "\"data_hurdle_ramp_type\"" + " : " + "\"" + data_hurdle_ramp_type + "\","
                        + "\n" + "\"manager_id\"" + " : " + manager_id + ","
                        + "\n" + "\"manager_source\"" + " : " + "\"" + manager_source + "\"" + "\n" +
                       "}";
                var addPublicFund = WorkbenchApi.AddFundApi(body, msalIdtoken);
                Assert.That(addPublicFund.StatusCode, Is.EqualTo(HttpStatusCode.Created));

                // Parse IRestResponse to JObject
                JObject addPublicFundJs = JObject.Parse(addPublicFund.Content);
                ClassicAssert.AreEqual(fund_name, addPublicFundJs["fund_name"].ToString());
                ClassicAssert.AreEqual(manager_id.ToString(), addPublicFundJs["manager_id"].ToString());
                ClassicAssert.AreEqual(manager_name, addPublicFundJs["manager_name"].ToString());
                ClassicAssert.AreEqual("manual", addPublicFundJs["source"].ToString());
                ClassicAssert.AreEqual(sub_asset_class, addPublicFundJs["sub_asset_class"].ToString());
                ClassicAssert.AreEqual(asset_class, addPublicFundJs["asset_class"].ToString());
                //Assert.AreEqual(inception_date, addPublicFundJs["inception_date"].ToString());
                ClassicAssert.AreEqual(latest_actual_value, addPublicFundJs["latest_actual_value"].ToString());
                ClassicAssert.AreEqual(business_street, addPublicFundJs["business_street"].ToString());
                ClassicAssert.AreEqual(business_city, addPublicFundJs["business_city"].ToString());
                ClassicAssert.AreEqual(business_state, addPublicFundJs["business_state"].ToString());
                ClassicAssert.AreEqual(business_zip, addPublicFundJs["business_zip"].ToString());
                ClassicAssert.AreEqual(business_country, addPublicFundJs["business_country"].ToString());
                #endregion
            }
            else Console.WriteLine("Add Public Fund Api auto test is only add new Fund on Sandbox Site!!!");
        }

        [Test, Category("API Smoke Tests")]
        public void ST002_EditPublicFund()
        {
            #region Variables declare
            int? fundId = null;
            int? manager_id = null;
            string? manager_name = null;
            string? fund_name = null, fund_name_update = null;
            string? sub_asset_class = null, sub_asset_class_update = null;
            string? asset_class = null, asset_class_update = null;
            //string? date = null, inception_date = null;
            string? latest_actual_value = null, latest_actual_value_update = null;
            string? business_street = null, business_street_update = null;
            string? business_city = null, business_city_update = null;
            string? business_state = null, business_state_update = null;
            string? business_zip = null, business_zip_update = null;
            string? business_country = null, business_country_update = null;
            //string? data_redemption_frequency = null;
            //string? data_redemption_notice_days = null;
            //string? data_hard_lockup = null;
            //string? data_hard_lockup_months = null;
            //string? data_soft_lockup_months = null;
            //string? data_redemption_fee = null;
            //string? data_redemption_gate = null;
            //string? data_redemption_gate_percent = null;
            //string? data_percent_of_nav_available = null;
            //string? data_side_pocket = null;
            string? data_redemption_note = null, data_redemption_note_update = null;
            //string? data_management_fee = null;
            //string? data_management_fee_frequency = null;
            //string? data_performance_fee = null;
            //string? data_hwm_status = null;
            //string? data_catch_up = null;
            //string? data_catch_up_rate = null;
            //string? data_crystalization_frequency = null;
            //string? data_hurdle_status = null;
            //string? data_hurdle_type = null;
            //string? data_benchmark = null;
            //string? data_benchmark_name = null;
            //string? data_hurdle_rate = null;
            //string? data_hurdle_hard_soft = null;
            //string? data_hurdle_ramp_type = null;
            string? manager_source = null;
            #endregion

            #region Check if api of Sandbox or Staging then Edit Fund (on that site)
            if (workbenchApi.Contains("sandbox"))
            {
                fundId = 27;
                manager_name = "QA Test 05";
                fund_name = "Main QA Test 05";
                fund_name_update = fund_name + "_Update";
                sub_asset_class = "US Buyout";
                sub_asset_class_update = "US Growth Equity"; 
                asset_class = "Private Equity";
                asset_class_update = "Private Equity";
                //date = DateTime.Now.Date.ToString("yyyy-MM-dd"); inception_date = date.Replace("/", "-");
                latest_actual_value = "QA auto update Latest Actual Value";
                latest_actual_value_update = latest_actual_value + "_Update";
                business_street = "QA auto update Business Street";
                business_street_update = business_street + "_Update";
                business_city = "QA auto update Business City";
                business_city_update = business_city + "_Update";
                business_state = "QA auto update Business State";
                business_state_update = business_state + "_Update";
                business_zip = "QA auto update Business ZIP";
                business_zip_update = business_zip + "_Update";
                business_country = "UK";
                business_country_update = "JP";
                //data_redemption_frequency = "Monthly";
                //data_redemption_notice_days = "59 Days";
                //data_hard_lockup = "Yes";
                //data_hard_lockup_months = "35.01";
                //data_soft_lockup_months = "14.01";
                //data_redemption_fee = "16.01";
                //data_redemption_gate = "10.01";
                //data_redemption_gate_percent = "8.51";
                //data_percent_of_nav_available = "26.51";
                //data_side_pocket = "Yes";
                data_redemption_note = "QA auto update Additional Notes On Liquidity";
                data_redemption_note_update = data_redemption_note + "_Update";
                //data_management_fee = "-10.51";
                //data_management_fee_frequency = "Monthly";
                //data_performance_fee = "-6.51";
                //data_hwm_status = "Yes";
                //data_catch_up = "Yes";
                //data_catch_up_rate = "-17.51";
                //data_crystalization_frequency = "2";
                //data_hurdle_status = "Yes";
                //data_hurdle_type = "Fixed";
                //data_benchmark = "MS133333.USD";
                //data_benchmark_name = "MSCI CHINA A ONSHORE Net Return Index USD";
                //data_hurdle_rate = "-5.51";
                //data_hurdle_hard_soft = "Hard";
                //data_hurdle_ramp_type = "Performance Dependent";
                manager_id = 25;
                manager_source = "manual";
            }

            if (workbenchApi.Contains("conceptia"))
            {
                fundId = 2;
                manager_name = "QA Test 01";
                fund_name = "Main of QA Test 01";
                fund_name_update = fund_name + "_Update";
                sub_asset_class = "US Buyout";
                sub_asset_class_update = "US Growth Equity";
                asset_class = "Private Equity";
                asset_class_update = "Private Equity";
                //date = DateTime.Now.Date.ToString("yyyy-MM-dd"); inception_date = date.Replace("/", "-");
                latest_actual_value = "QA api auto Latest Actual Value";
                latest_actual_value_update = latest_actual_value + "_Update";
                business_street = "QA api auto Business Street";
                business_street_update = business_street + "_Update";
                business_city = "QA api auto Business City";
                business_city_update = business_city + "_Update";
                business_state = "QA api auto Business State";
                business_state_update = business_state + "_Update";
                business_zip = "QA api auto Business ZIP";
                business_zip_update = business_zip + "_Update";
                business_country = "JP";
                business_country_update = "AF";
                //data_redemption_frequency = "Monthly";
                //data_redemption_notice_days = "59 Days";
                //data_hard_lockup = "Yes";
                //data_hard_lockup_months = "35.01";
                //data_soft_lockup_months = "14.01";
                //data_redemption_fee = "16.01";
                //data_redemption_gate = "10.01";
                //data_redemption_gate_percent = "8.51";
                //data_percent_of_nav_available = "26.51";
                //data_side_pocket = "Yes";
                data_redemption_note = "QA api auto update Additional Notes On Liquidity";
                data_redemption_note_update = data_redemption_note + "_Update";
                //data_management_fee = "-10.51";
                //data_management_fee_frequency = "Monthly";
                //data_performance_fee = "-6.51";
                //data_hwm_status = "Yes";
                //data_catch_up = "Yes";
                //data_catch_up_rate = "-17.51";
                //data_crystalization_frequency = "2";
                //data_hurdle_status = "Yes";
                //data_hurdle_type = "Fixed";
                //data_benchmark = "MS133333.USD";
                //data_benchmark_name = "MSCI CHINA A ONSHORE Net Return Index USD";
                //data_hurdle_rate = "-5.51";
                //data_hurdle_hard_soft = "Hard";
                //data_hurdle_ramp_type = "Performance Dependent";
                manager_id = 2;
                manager_source = "manual";
            }
            #endregion

            var body = "{"
                            + "\n" + "\"manager_name\"" + " : " + "\"" + manager_name + "\","
                            + "\n" + "\"fund_name\"" + " : " + "\"" + fund_name_update + "\","
                            + "\n" + "\"sub_asset_class\"" + " : " + "\"" + sub_asset_class_update + "\","
                            + "\n" + "\"asset_class\"" + " : " + "\"" + asset_class_update + "\","
                            //+ "\n" + "\"inception_date\"" + " : " + "\"" + inception_date + "\","
                            + "\n" + "\"latest_actual_value\"" + " : " + "\"" + latest_actual_value_update + "\","
                            + "\n" + "\"business_street\"" + " : " + "\"" + business_street_update + "\","
                            + "\n" + "\"business_city\"" + " : " + "\"" + business_city_update + "\","
                            + "\n" + "\"business_state\"" + " : " + "\"" + business_state_update + "\","
                            + "\n" + "\"business_zip\"" + " : " + "\"" + business_zip_update + "\","
                            + "\n" + "\"business_country\"" + " : " + "\"" + business_country_update + "\","
                            //+ "\n" + "\"data_redemption_frequency\"" + " : " + "\"" + data_redemption_frequency + "\","
                            //+ "\n" + "\"data_redemption_notice_days\"" + " : " + "\"" + data_redemption_notice_days + "\","
                            //+ "\n" + "\"data_hard_lockup\"" + " : " + "\"" + data_hard_lockup + "\","
                            //+ "\n" + "\"data_hard_lockup_months\"" + " : " + "\"" + data_hard_lockup_months + "\","
                            //+ "\n" + "\"data_soft_lockup_months\"" + " : " + "\"" + data_soft_lockup_months + "\","
                            //+ "\n" + "\"data_redemption_fee\"" + " : " + "\"" + data_redemption_fee + "\","
                            //+ "\n" + "\"data_redemption_gate\"" + " : " + "\"" + data_redemption_gate + "\","
                            //+ "\n" + "\"data_redemption_gate_percent\"" + " : " + "\"" + data_redemption_gate_percent + "\","
                            //+ "\n" + "\"data_percent_of_nav_available\"" + " : " + "\"" + data_percent_of_nav_available + "\","
                            //+ "\n" + "\"data_side_pocket\"" + " : " + "\"" + data_side_pocket + "\","
                            + "\n" + "\"data_redemption_note\"" + " : " + "\"" + data_redemption_note_update + "\","
                            //+ "\n" + "\"data_management_fee\"" + " : " + "\"" + data_management_fee + "\","
                            //+ "\n" + "\"data_management_fee_frequency\"" + " : " + "\"" + data_management_fee_frequency + "\","
                            //+ "\n" + "\"data_performance_fee\"" + " : " + "\"" + data_performance_fee + "\","
                            //+ "\n" + "\"data_hwm_status\"" + " : " + "\"" + data_hwm_status + "\","
                            //+ "\n" + "\"data_catch_up\"" + " : " + "\"" + data_catch_up + "\","
                            //+ "\n" + "\"data_catch_up_rate\"" + " : " + "\"" + data_catch_up_rate + "\","
                            //+ "\n" + "\"data_crystalization_frequency\"" + " : " + "\"" + data_crystalization_frequency + "\","
                            //+ "\n" + "\"data_hurdle_status\"" + " : " + "\"" + data_hurdle_status + "\","
                            //+ "\n" + "\"data_hurdle_type\"" + " : " + "\"" + data_hurdle_type + "\","
                            //+ "\n" + "\"data_benchmark\"" + " : " + "\"" + data_benchmark + "\","
                            //+ "\n" + "\"data_benchmark_name\"" + " : " + "\"" + data_benchmark_name + "\","
                            //+ "\n" + "\"data_hurdle_rate\"" + " : " + "\"" + data_hurdle_rate + "\","
                            //+ "\n" + "\"data_hurdle_hard_soft\"" + " : " + "\"" + data_hurdle_hard_soft + "\","
                            //+ "\n" + "\"data_hurdle_ramp_type\"" + " : " + "\"" + data_hurdle_ramp_type + "\","
                            + "\n" + "\"manager_id\"" + " : " + manager_id + ","
                            + "\n" + "\"manager_source\"" + " : " + "\"" + manager_source + "\"" + "\n" +
                       "}";

            // Edit Public Fund
            var editPublicFund = WorkbenchApi.EditFundApi(fundId.ToString(), body, msalIdtoken);
            Assert.That(editPublicFund.StatusCode, Is.EqualTo(HttpStatusCode.OK));

            // Parse IRestResponse to JObject
            var editPublicFundJs = JObject.Parse(editPublicFund.Content);
            ClassicAssert.AreEqual(fund_name_update, editPublicFundJs["fund_name"].ToString());
            ClassicAssert.AreEqual(manager_id.ToString(), editPublicFundJs["manager_id"].ToString());
            ClassicAssert.AreEqual(manager_name, editPublicFundJs["manager_name"].ToString());
            ClassicAssert.AreEqual(sub_asset_class_update, editPublicFundJs["sub_asset_class"].ToString());
            ClassicAssert.AreEqual(asset_class_update, editPublicFundJs["asset_class"].ToString());
            ClassicAssert.AreEqual(latest_actual_value_update, editPublicFundJs["latest_actual_value"].ToString());
            ClassicAssert.AreEqual(business_street_update, editPublicFundJs["business_street"].ToString());
            ClassicAssert.AreEqual(business_city_update, editPublicFundJs["business_city"].ToString());
            ClassicAssert.AreEqual(business_state_update, editPublicFundJs["business_state"].ToString());
            ClassicAssert.AreEqual(business_zip_update, editPublicFundJs["business_zip"].ToString());
            ClassicAssert.AreEqual(business_country_update, editPublicFundJs["business_country"].ToString());
            ClassicAssert.AreEqual(data_redemption_note_update, editPublicFundJs["data_redemption_note"].ToString());

            // Re-Update the Fund Name (back to the original Fund Name)
            body = "{"
                            + "\n" + "\"manager_name\"" + " : " + "\"" + manager_name + "\","
                            + "\n" + "\"fund_name\"" + " : " + "\"" + fund_name + "\","
                            + "\n" + "\"sub_asset_class\"" + " : " + "\"" + sub_asset_class + "\","
                            + "\n" + "\"asset_class\"" + " : " + "\"" + asset_class + "\","
                            //+ "\n" + "\"inception_date\"" + " : " + "\"" + inception_date + "\","
                            + "\n" + "\"latest_actual_value\"" + " : " + "\"" + latest_actual_value + "\","
                            + "\n" + "\"business_street\"" + " : " + "\"" + business_street + "\","
                            + "\n" + "\"business_city\"" + " : " + "\"" + business_city + "\","
                            + "\n" + "\"business_state\"" + " : " + "\"" + business_state + "\","
                            + "\n" + "\"business_zip\"" + " : " + "\"" + business_zip + "\","
                            + "\n" + "\"business_country\"" + " : " + "\"" + business_country + "\","
                            //+ "\n" + "\"data_redemption_frequency\"" + " : " + "\"" + data_redemption_frequency + "\","
                            //+ "\n" + "\"data_redemption_notice_days\"" + " : " + "\"" + data_redemption_notice_days + "\","
                            //+ "\n" + "\"data_hard_lockup\"" + " : " + "\"" + data_hard_lockup + "\","
                            //+ "\n" + "\"data_hard_lockup_months\"" + " : " + "\"" + data_hard_lockup_months + "\","
                            //+ "\n" + "\"data_soft_lockup_months\"" + " : " + "\"" + data_soft_lockup_months + "\","
                            //+ "\n" + "\"data_redemption_fee\"" + " : " + "\"" + data_redemption_fee + "\","
                            //+ "\n" + "\"data_redemption_gate\"" + " : " + "\"" + data_redemption_gate + "\","
                            //+ "\n" + "\"data_redemption_gate_percent\"" + " : " + "\"" + data_redemption_gate_percent + "\","
                            //+ "\n" + "\"data_percent_of_nav_available\"" + " : " + "\"" + data_percent_of_nav_available + "\","
                            //+ "\n" + "\"data_side_pocket\"" + " : " + "\"" + data_side_pocket + "\","
                            + "\n" + "\"data_redemption_note\"" + " : " + "\"" + data_redemption_note + "\","
                            //+ "\n" + "\"data_management_fee\"" + " : " + "\"" + data_management_fee + "\","
                            //+ "\n" + "\"data_management_fee_frequency\"" + " : " + "\"" + data_management_fee_frequency + "\","
                            //+ "\n" + "\"data_performance_fee\"" + " : " + "\"" + data_performance_fee + "\","
                            //+ "\n" + "\"data_hwm_status\"" + " : " + "\"" + data_hwm_status + "\","
                            //+ "\n" + "\"data_catch_up\"" + " : " + "\"" + data_catch_up + "\","
                            //+ "\n" + "\"data_catch_up_rate\"" + " : " + "\"" + data_catch_up_rate + "\","
                            //+ "\n" + "\"data_crystalization_frequency\"" + " : " + "\"" + data_crystalization_frequency + "\","
                            //+ "\n" + "\"data_hurdle_status\"" + " : " + "\"" + data_hurdle_status + "\","
                            //+ "\n" + "\"data_hurdle_type\"" + " : " + "\"" + data_hurdle_type + "\","
                            //+ "\n" + "\"data_benchmark\"" + " : " + "\"" + data_benchmark + "\","
                            //+ "\n" + "\"data_benchmark_name\"" + " : " + "\"" + data_benchmark_name + "\","
                            //+ "\n" + "\"data_hurdle_rate\"" + " : " + "\"" + data_hurdle_rate + "\","
                            //+ "\n" + "\"data_hurdle_hard_soft\"" + " : " + "\"" + data_hurdle_hard_soft + "\","
                            //+ "\n" + "\"data_hurdle_ramp_type\"" + " : " + "\"" + data_hurdle_ramp_type + "\","
                            + "\n" + "\"manager_id\"" + " : " + manager_id + ","
                            + "\n" + "\"manager_source\"" + " : " + "\"" + manager_source + "\"" + "\n" +
                       "}";

            // Edit Public Fund
            var reEditPublicFund = WorkbenchApi.EditFundApi(fundId.ToString(), body, msalIdtoken);
            ClassicAssert.That(reEditPublicFund.StatusCode, Is.EqualTo(HttpStatusCode.OK));

            // Parse IRestResponse to JObject
            var reEditPublicFundJs = JObject.Parse(reEditPublicFund.Content);
            ClassicAssert.AreEqual(fund_name, reEditPublicFundJs["fund_name"].ToString());
        }

        [Test, Category("API Smoke Tests")]
        public void ST003_AddPrivateFund()
        {
            // Variables declare
            string firm = "QA_ApiAuto_FirmPriv" + @"_" + DateTime.Now.ToString().Replace("/", "_").Replace(":", "_").Replace(" ", "_");
            string manager_name = "QA_ApiAuto_ManagerPriv" + @"_" + DateTime.Now.ToString().Replace("/", "_").Replace(":", "_").Replace(" ", "_");
            string sub_asset_class = "FAD Real Estate";
            string date = DateTime.Now.Date.ToString("MM-dd-yyyy"), inception_date = date.Replace("/", "-");
            string latest_actual_value = "QA_ApiAuto_LAV";
            string business_street = "QA_ApiAuto_BusinessStreet";
            string business_city = "QA_ApiAuto_BusinessCity";
            string business_state = "QA_ApiAuto_BusinessState";
            string business_zip = "QA_ApiAuto_BusinessZIP";
            string business_country = "QA_ApiAuto_BusinessCountry";
            string business_contact = "QA_ApiAuto_BusinessContact";
            string business_email = "QA_ApiAuto_BusinessEmail";
            string business_phone = "123456789";
            var body = "{"
                        + "\n" + "\"firm\"" + " : " + "\"" + firm + "\","
                        + "\n" + "\"manager_name\"" + " : " + "\"" + manager_name + "\","
                        + "\n" + "\"sub_asset_class\"" + " : " + "\"" + sub_asset_class + "\","
                        + "\n" + "\"inception_date\"" + " : " + "\"" + inception_date + "\","
                        + "\n" + "\"latest_actual_value\"" + " : " + "\"" + latest_actual_value + "\","
                        + "\n" + "\"business_street\"" + " : " + "\"" + business_street + "\","
                        + "\n" + "\"business_city\"" + " : " + "\"" + business_city + "\","
                        + "\n" + "\"business_state\"" + " : " + "\"" + business_state + "\","
                        + "\n" + "\"business_zip\"" + " : " + "\"" + business_zip + "\","
                        + "\n" + "\"business_country\"" + " : " + "\"" + business_country + "\","
                        + "\n" + "\"business_contact\"" + " : " + "\"" + business_contact + "\","
                        + "\n" + "\"business_email\"" + " : " + "\"" + business_email + "\","
                        + "\n" + "\"business_phone\"" + " : " + "\"" + business_phone + "\"" + "\n" +
                       "}";

            // Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
            if (workbenchApi.Contains("sandbox"))
            {
                #region Add Private Fund - Manager
                // Add Public Fund Manager
                var addPrivateFundManager = WorkbenchApi.AddManagerPrivateApi(body, msalIdtoken);
                Assert.That(addPrivateFundManager.StatusCode, Is.EqualTo(HttpStatusCode.Created));

                // Parse IRestResponse to JObject
                var addPrivateFundManagerJs = JObject.Parse(addPrivateFundManager.Content);
                ClassicAssert.AreEqual(manager_name, addPrivateFundManagerJs["manager_name"].ToString());
                ClassicAssert.AreEqual(firm, addPrivateFundManagerJs["firm"].ToString());
                ClassicAssert.AreEqual("private_manual", addPrivateFundManagerJs["source"].ToString());
                ClassicAssert.AreEqual(sub_asset_class, addPrivateFundManagerJs["sub_asset_class"].ToString());
                ClassicAssert.AreEqual(latest_actual_value, addPrivateFundManagerJs["latest_actual_value"].ToString());
                ClassicAssert.AreEqual(business_street, addPrivateFundManagerJs["business_street"].ToString());
                ClassicAssert.AreEqual(business_city, addPrivateFundManagerJs["business_city"].ToString());
                ClassicAssert.AreEqual(business_state, addPrivateFundManagerJs["business_state"].ToString());
                ClassicAssert.AreEqual(business_zip, addPrivateFundManagerJs["business_zip"].ToString());
                ClassicAssert.AreEqual(business_country, addPrivateFundManagerJs["business_country"].ToString());
                ClassicAssert.AreEqual(business_contact, addPrivateFundManagerJs["business_contact"].ToString());
                ClassicAssert.AreEqual(business_email, addPrivateFundManagerJs["business_email"].ToString());
                ClassicAssert.AreEqual(business_phone, addPrivateFundManagerJs["business_phone"].ToString());
                #endregion

                #region Add Private Fund - Fund
                // Add Private Fund (Fund)
                string fund_name = "Fund 1 of QA_ApiAuto_FirmPriv";
                string strategy = "QA_ApiAuto_Strategy";
                string year_firm_founded = "QA_ApiAuto_YFM";
                string strategy_headquarters = "QA_ApiAuto_SH";
                string asset_class = "QA_ApiAuto_AC";
                string investment_stage = "QA_ApiAuto_IS";
                string industry_focus = "QA_ApiAuto_IF";
                string geographic_focus = "QA_ApiAuto_GF";
                string fund_size_expected_size_m = "188";
                string fund_size_expected_size_m_currency = "USD";
                int manager_id = int.Parse(addPrivateFundManagerJs["manager_id"].ToString());
                body = "{"
                        + "\n" + "\"fund_name\"" + " : " + "\"" + fund_name + "\","
                        + "\n" + "\"strategy\"" + " : " + "\"" + strategy + "\","
                        + "\n" + "\"year_firm_founded\"" + " : " + "\"" + year_firm_founded + "\","
                        + "\n" + "\"strategy_headquarters\"" + " : " + "\"" + strategy_headquarters + "\","
                        + "\n" + "\"asset_class\"" + " : " + "\"" + asset_class + "\","
                        + "\n" + "\"investment_stage\"" + " : " + "\"" + investment_stage + "\","
                        + "\n" + "\"industry_focus\"" + " : " + "\"" + industry_focus + "\","
                        + "\n" + "\"geographic_focus\"" + " : " + "\"" + geographic_focus + "\","
                        + "\n" + "\"fund_size_expected_size_m\"" + " : " + "\"" + fund_size_expected_size_m + "\","
                        + "\n" + "\"fund_size_expected_size_m_currency\"" + " : " + "\"" + fund_size_expected_size_m_currency + "\","
                        + "\n" + "\"manager_id\"" + " : " + manager_id + "\n" +
                       "}";

                var addPrivateFund = WorkbenchApi.AddFundPrivateApi(body, msalIdtoken);
                Assert.That(addPrivateFund.StatusCode, Is.EqualTo(HttpStatusCode.Created));

                // Parse IRestResponse to JObject
                var addPrivateFundJs = JObject.Parse(addPrivateFund.Content);
                ClassicAssert.AreEqual("", addPrivateFundJs["fund"].ToString());
                ClassicAssert.AreEqual(manager_id.ToString(), addPrivateFundJs["manager_id"].ToString());
                ClassicAssert.AreEqual("private_manual", addPrivateFundJs["source"].ToString());
                ClassicAssert.AreEqual(fund_name, addPrivateFundJs["fund_name"].ToString());
                ClassicAssert.AreEqual(strategy, addPrivateFundJs["strategy"].ToString());
                ClassicAssert.AreEqual(year_firm_founded, addPrivateFundJs["year_firm_founded"].ToString());
                ClassicAssert.AreEqual(strategy_headquarters, addPrivateFundJs["strategy_headquarters"].ToString());
                ClassicAssert.AreEqual(asset_class, addPrivateFundJs["asset_class"].ToString());
                ClassicAssert.AreEqual(investment_stage, addPrivateFundJs["investment_stage"].ToString());
                ClassicAssert.AreEqual(industry_focus, addPrivateFundJs["industry_focus"].ToString());
                ClassicAssert.AreEqual(geographic_focus, addPrivateFundJs["geographic_focus"].ToString());
                ClassicAssert.AreEqual(fund_size_expected_size_m, addPrivateFundJs["fund_size_expected_size_m"].ToString());
                ClassicAssert.AreEqual(fund_size_expected_size_m_currency, addPrivateFundJs["fund_size_expected_size_m_currency"].ToString());
                #endregion
            }
            else Console.WriteLine("Add Private Fund Api auto test is only add new Fund on Sandbox Site!!!");
        }

        [Test, Category("API Smoke Tests")]
        public void ST004_EditPrivateFund()
        {
            #region Variables declare
            // Manager info
            int? manager_id = null;
            string? firm = null;
            string? manager_name = null, manager_name_update = null;
            string? sub_asset_class = null, sub_asset_class_update = null;
            string? inception_date=null;
            string? latest_actual_value = null, latest_actual_value_update = null;
            string? business_street = null, business_street_update = null;
            string? business_city = null, business_city_update = null;
            string? business_state = null, business_state_update = null;
            string? business_zip = null, business_zip_update = null;
            string? business_country = null, business_country_update = null;
            string? business_contact = null, business_contact_update = null;
            string? business_email = null, business_email_update = null;
            string? business_phone = null, business_phone_update = null;

            // Fund Info
            int? fund_id = null;
            string? fund_name = null, fund_name_update = null;
            string? strategy = null, strategy_update = null;
            string? year_firm_founded = null, year_firm_founded_update = null;
            string? strategy_headquarters = null, strategy_headquarters_update = null;
            string? asset_class = null, asset_class_update = null;
            string? investment_stage = null, investment_stage_update = null;
            string? industry_focus = null, industry_focus_update = null;
            string? geographic_focus = null, geographic_focus_update = null;
            string? fund_size_expected_size_m = null, fund_size_expected_size_m_update = null;
            string? fund_size_expected_size_m_currency = null, fund_size_expected_size_m_currency_update = null;
            #endregion

            #region Check if api of Sandbox or Staging then Edit Fund (on that site)
            if (workbenchApi.Contains("sandbox"))
            {
                // Manager info
                manager_id = 286; // 31 is removed
                firm = "Firmofpriman qatest 1";
                manager_name = "Priman qatest 1"; manager_name_update = manager_name + "_Update";
                sub_asset_class = "US Buyout"; sub_asset_class_update = "FAD Real Estate";
                inception_date = "09/23/2022";
                latest_actual_value = "LAVA 01"; latest_actual_value_update = latest_actual_value + "_Update";
                business_street = "BSTR 01"; business_street_update = business_street + "_Update";
                business_city = "BCIT 01"; business_city_update = business_city + "_Update";
                business_state = "BSTA 01"; business_state_update = business_state + "_Update";
                business_zip = "BZIP 01"; business_zip_update = business_zip + "_Update";
                business_country = "United Kingdom"; business_country_update = "Japan";
                business_contact = "BCON 01"; business_contact_update = business_contact + "_Update";
                business_email = "BEMA01@test.com"; business_email_update = "_Update" + business_email;
                business_phone = "123456789"; business_phone_update = "987654321";

                // Fund Info
                //fund_id = 619; // 607 (old id)
                fund_name = "F1 of Firmofpriman qatest 1"; fund_name_update = fund_name + "_Update";
                strategy = "STRA 101"; strategy_update = strategy + "_Update";
                year_firm_founded = "2020"; year_firm_founded_update = "2025";
                strategy_headquarters = "SHQA 101"; strategy_headquarters_update = strategy_headquarters + "_Update";
                asset_class = "AC 101"; asset_class_update = asset_class + "_Update";
                investment_stage = "IS 101"; investment_stage_update = investment_stage + "_Update";
                industry_focus = "IF 101"; industry_focus_update = industry_focus + "_Update";
                geographic_focus = "GF 101"; geographic_focus_update = geographic_focus + "_Update";
                fund_size_expected_size_m = "101"; fund_size_expected_size_m_update = "1001";
                fund_size_expected_size_m_currency = "USD"; fund_size_expected_size_m_currency_update = "VND";
            }

            if (workbenchApi.Contains("conceptia"))
            {
                // Manager info
                manager_id = 6;
                firm = "Priman Sta01";
                manager_name = "Priman Sta02"; manager_name_update = manager_name + "_Update";
                sub_asset_class = "FAD Real Estate"; sub_asset_class_update = "US Buyout";
                inception_date = "09/13/2022";
                latest_actual_value = "LAVA 01"; latest_actual_value_update = latest_actual_value + "_Update";
                business_street = "BSTR 01"; business_street_update = business_street + "_Update";
                business_city = "BCIT 01"; business_city_update = business_city + "_Update";
                business_state = "BSTA 01"; business_state_update = business_state + "_Update";
                business_zip = "BZIP 01"; business_zip_update = business_zip + "_Update";
                business_country = "United Kingdom"; business_country_update = "Japan";
                business_contact = "BCON 01"; business_contact_update = business_contact + "_Update";
                business_email = "BEMA01@test.com"; business_email_update = "_Update" + business_email;
                business_phone = "123456789"; business_phone_update = "987654321";

                // Fund Info
                //fund_id = 102 // 102 (old id)
                fund_name = "Fund 1 of Firm Sta01"; fund_name_update = fund_name + "_Update";
                strategy = "STRA 101"; strategy_update = strategy + "_Update";
                year_firm_founded = "2020"; year_firm_founded_update = "2025";
                strategy_headquarters = "SHQA 101"; strategy_headquarters_update = strategy_headquarters + "_Update";
                asset_class = "AC 101"; asset_class_update = asset_class + "_Update";
                investment_stage = "IS 101"; investment_stage_update = investment_stage + "_Update";
                industry_focus = "IF 101"; industry_focus_update = industry_focus + "_Update";
                geographic_focus = "GF 101"; geographic_focus_update = geographic_focus + "_Update";
                fund_size_expected_size_m = "101"; fund_size_expected_size_m_update = "1001";
                fund_size_expected_size_m_currency = "USD"; fund_size_expected_size_m_currency_update = "VND";
            }
            #endregion

            #region Edit/Update Private Fund - Manager
            var body = "{"
                        + "\n" + "\"firm\"" + " : " + "\"" + firm + "\","
                        + "\n" + "\"manager_name\"" + " : " + "\"" + manager_name_update + "\","
                        + "\n" + "\"sub_asset_class\"" + " : " + "\"" + sub_asset_class_update + "\","
                        + "\n" + "\"inception_date\"" + " : " + "\"" + inception_date + "\","
                        + "\n" + "\"latest_actual_value\"" + " : " + "\"" + latest_actual_value_update + "\","
                        + "\n" + "\"business_street\"" + " : " + "\"" + business_street_update + "\","
                        + "\n" + "\"business_city\"" + " : " + "\"" + business_city_update + "\","
                        + "\n" + "\"business_state\"" + " : " + "\"" + business_state_update + "\","
                        + "\n" + "\"business_zip\"" + " : " + "\"" + business_zip_update + "\","
                        + "\n" + "\"business_country\"" + " : " + "\"" + business_country_update + "\","
                        + "\n" + "\"business_contact\"" + " : " + "\"" + business_contact_update + "\","
                        + "\n" + "\"business_email\"" + " : " + "\"" + business_email_update + "\","
                        + "\n" + "\"business_phone\"" + " : " + "\"" + business_phone_update + "\"" + "\n" +
                       "}";

            // Edit/Update Private Fund - Manager
            var editManagerPrivateFund = WorkbenchApi.EditManagerPrivateApi(manager_id.ToString(), body, msalIdtoken);
            Assert.That(editManagerPrivateFund.StatusCode, Is.EqualTo(HttpStatusCode.OK));

            // Parse IRestResponse to JObject
            var editManagerPrivateFundJs = JObject.Parse(editManagerPrivateFund.Content);
            ClassicAssert.AreEqual(manager_name_update, editManagerPrivateFundJs["manager_name"].ToString());
            ClassicAssert.AreEqual(firm, editManagerPrivateFundJs["firm"].ToString());
            ClassicAssert.AreEqual("private_manual", editManagerPrivateFundJs["source"].ToString());
            ClassicAssert.AreEqual(sub_asset_class_update, editManagerPrivateFundJs["sub_asset_class"].ToString());
            ClassicAssert.AreEqual(latest_actual_value_update, editManagerPrivateFundJs["latest_actual_value"].ToString());
            ClassicAssert.AreEqual(business_street_update, editManagerPrivateFundJs["business_street"].ToString());
            ClassicAssert.AreEqual(business_city_update, editManagerPrivateFundJs["business_city"].ToString());
            ClassicAssert.AreEqual(business_state_update, editManagerPrivateFundJs["business_state"].ToString());
            ClassicAssert.AreEqual(business_zip_update, editManagerPrivateFundJs["business_zip"].ToString());
            ClassicAssert.AreEqual(business_country_update, editManagerPrivateFundJs["business_country"].ToString());
            ClassicAssert.AreEqual(business_contact_update, editManagerPrivateFundJs["business_contact"].ToString());
            ClassicAssert.AreEqual(business_email_update, editManagerPrivateFundJs["business_email"].ToString());
            ClassicAssert.AreEqual(business_phone_update, editManagerPrivateFundJs["business_phone"].ToString());
            ClassicAssert.AreEqual(manager_id.ToString(), editManagerPrivateFundJs["manager_id"].ToString());

            // Re-Update the manager Name (back to the original Manager Name)
            body = "{"
                        + "\n" + "\"firm\"" + " : " + "\"" + firm + "\","
                        + "\n" + "\"manager_name\"" + " : " + "\"" + manager_name + "\","
                        + "\n" + "\"sub_asset_class\"" + " : " + "\"" + sub_asset_class + "\","
                        + "\n" + "\"inception_date\"" + " : " + "\"" + inception_date + "\","
                        + "\n" + "\"latest_actual_value\"" + " : " + "\"" + latest_actual_value + "\","
                        + "\n" + "\"business_street\"" + " : " + "\"" + business_street + "\","
                        + "\n" + "\"business_city\"" + " : " + "\"" + business_city + "\","
                        + "\n" + "\"business_state\"" + " : " + "\"" + business_state + "\","
                        + "\n" + "\"business_zip\"" + " : " + "\"" + business_zip + "\","
                        + "\n" + "\"business_country\"" + " : " + "\"" + business_country + "\","
                        + "\n" + "\"business_contact\"" + " : " + "\"" + business_contact + "\","
                        + "\n" + "\"business_email\"" + " : " + "\"" + business_email + "\","
                        + "\n" + "\"business_phone\"" + " : " + "\"" + business_phone + "\"" + "\n" +
                   "}";

            // Edit/Update Private Fund - Manager
            var ReEditManagerPrivateFund = WorkbenchApi.EditManagerPrivateApi(manager_id.ToString(), body, msalIdtoken);
            Assert.That(ReEditManagerPrivateFund.StatusCode, Is.EqualTo(HttpStatusCode.OK));
            var ReEditManagerPrivateFundJs = JObject.Parse(ReEditManagerPrivateFund.Content); // Parse IRestResponse to JObject
            ClassicAssert.AreEqual(manager_name, ReEditManagerPrivateFundJs["manager_name"].ToString());
            #endregion

            #region Edit/Update Private Fund - Fund
            // Query data in Datalake to get fund_id
            string query = "SELECT * FROM private_fund_manual WHERE manager_id=" + manager_id;
            var queryPrivateFundManual = DatalakeApi.QueryDataInDataLake("/api/datalake", "gend_ks_db", query, bearerToken);
            List<JObject>? getPrivateFundManual = JsonConvert.DeserializeObject<List<JObject>>(queryPrivateFundManual.Content);
            foreach ( var obj in getPrivateFundManual )
            {
                if (obj.ContainsKey("fund_name"))
                {
                    var getValue = obj.GetValue("fund_name");
                    if (getValue.ToString() == fund_name) fund_id = (int?)obj.GetValue("fund_id");
                    break;
                }
            }

            // Edit/Update Private Fund - Fund
            body = "{"
                        + "\n" + "\"fund_id\"" + " : " + fund_id + ","
                        + "\n" + "\"fund_name\"" + " : " + "\"" + fund_name_update + "\","
                        + "\n" + "\"strategy\"" + " : " + "\"" + strategy_update + "\","
                        + "\n" + "\"year_firm_founded\"" + " : " + "\"" + year_firm_founded_update + "\","
                        + "\n" + "\"strategy_headquarters\"" + " : " + "\"" + strategy_headquarters_update + "\","
                        + "\n" + "\"asset_class\"" + " : " + "\"" + asset_class_update + "\","
                        + "\n" + "\"investment_stage\"" + " : " + "\"" + investment_stage_update + "\","
                        + "\n" + "\"industry_focus\"" + " : " + "\"" + industry_focus_update + "\","
                        + "\n" + "\"geographic_focus\"" + " : " + "\"" + geographic_focus_update + "\","
                        + "\n" + "\"fund_size_expected_size_m\"" + " : " + fund_size_expected_size_m_update + ","
                        + "\n" + "\"fund_size_expected_size_m_currency\"" + " : " + "\"" + fund_size_expected_size_m_currency_update + "\""
                        + "\n" +
                   "}";
            var editPrivateFund = WorkbenchApi.EditFundPrivateApi(fund_id.ToString(), body, msalIdtoken);
            Assert.That(editPrivateFund.StatusCode, Is.EqualTo(HttpStatusCode.OK));

            // Parse IRestResponse to JObject
            var editPrivateFundJs = JObject.Parse(editPrivateFund.Content);
            ClassicAssert.AreEqual("", editPrivateFundJs["fund"].ToString());
            ClassicAssert.AreEqual(manager_id.ToString(), editPrivateFundJs["manager_id"].ToString());
            ClassicAssert.AreEqual("private_manual", editPrivateFundJs["source"].ToString());
            ClassicAssert.AreEqual(fund_name, editPrivateFundJs["fund_name"].ToString());
            //Assert.AreEqual(fund_name_update, editPrivateFundJs["fund_name"].ToString());
            //Assert.AreEqual(strategy_update, editPrivateFundJs["strategy"].ToString());
            //Assert.AreEqual(year_firm_founded_update, editPrivateFundJs["year_firm_founded"].ToString());
            //Assert.AreEqual(strategy_headquarters_update, editPrivateFundJs["strategy_headquarters"].ToString());
            //Assert.AreEqual(asset_class_update, editPrivateFundJs["asset_class"].ToString());
            //Assert.AreEqual(investment_stage_update, editPrivateFundJs["investment_stage"].ToString());
            //Assert.AreEqual(industry_focus_update, editPrivateFundJs["industry_focus"].ToString());
            //Assert.AreEqual(geographic_focus_update, editPrivateFundJs["geographic_focus"].ToString());
            //Assert.AreEqual(fund_size_expected_size_m_update, editPrivateFundJs["fund_size_expected_size_m"].ToString());
            //Assert.AreEqual(fund_size_expected_size_m_currency_update, editPrivateFundJs["fund_size_expected_size_m_currency"].ToString());
            ClassicAssert.AreEqual(fund_id.ToString(), editPrivateFundJs["fund_id"].ToString());

            // Re-Update the Fund Name (back to the original Fund Name)
            body = "{"
                        + "\n" + "\"fund_id\"" + " : " + fund_id + ","
                        + "\n" + "\"fund_name\"" + " : " + "\"" + fund_name + "\","
                        + "\n" + "\"strategy\"" + " : " + "\"" + strategy + "\","
                        + "\n" + "\"year_firm_founded\"" + " : " + "\"" + year_firm_founded + "\","
                        + "\n" + "\"strategy_headquarters\"" + " : " + "\"" + strategy_headquarters + "\","
                        + "\n" + "\"asset_class\"" + " : " + "\"" + asset_class + "\","
                        + "\n" + "\"investment_stage\"" + " : " + "\"" + investment_stage + "\","
                        + "\n" + "\"industry_focus\"" + " : " + "\"" + industry_focus + "\","
                        + "\n" + "\"geographic_focus\"" + " : " + "\"" + geographic_focus + "\","
                        + "\n" + "\"fund_size_expected_size_m\"" + " : " + fund_size_expected_size_m + ","
                        + "\n" + "\"fund_size_expected_size_m_currency\"" + " : " + "\"" + fund_size_expected_size_m_currency + "\""
                        + "\n" +
                   "}";

            // ReEdit/Update Private Fund - Fund
            var reEditPrivateFund = WorkbenchApi.EditFundPrivateApi(fund_id.ToString(), body, msalIdtoken);
            Assert.That(reEditPrivateFund.StatusCode, Is.EqualTo(HttpStatusCode.OK));
            var reEditPrivateFundJs = JObject.Parse(reEditPrivateFund.Content); // Parse IRestResponse to JObject
            //Assert.AreEqual(fund_name, reEditPrivateFundJs["fund_name"].ToString());
            ClassicAssert.AreEqual(fund_name_update, reEditPrivateFundJs["fund_name"].ToString());
            ClassicAssert.AreEqual(strategy_update, reEditPrivateFundJs["strategy"].ToString());
            ClassicAssert.AreEqual(year_firm_founded_update, reEditPrivateFundJs["year_firm_founded"].ToString());
            ClassicAssert.AreEqual(strategy_headquarters_update, reEditPrivateFundJs["strategy_headquarters"].ToString());
            ClassicAssert.AreEqual(asset_class_update, reEditPrivateFundJs["asset_class"].ToString());
            ClassicAssert.AreEqual(investment_stage_update, reEditPrivateFundJs["investment_stage"].ToString());
            ClassicAssert.AreEqual(industry_focus_update, reEditPrivateFundJs["industry_focus"].ToString());
            ClassicAssert.AreEqual(geographic_focus_update, reEditPrivateFundJs["geographic_focus"].ToString());
            ClassicAssert.AreEqual(fund_size_expected_size_m_update, reEditPrivateFundJs["fund_size_expected_size_m"].ToString());
            ClassicAssert.AreEqual(fund_size_expected_size_m_currency_update, reEditPrivateFundJs["fund_size_expected_size_m_currency"].ToString());
            #endregion
        }
        #endregion
    }
}
