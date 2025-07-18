using AventStack.ExtentReports;
using NUnit.Framework;
using SeleniumGendKS.Core.BaseTestCase;
using SeleniumGendKS.Pages;
using System;
using System.Windows.Forms;

namespace SeleniumGendKS.Tests.Features_Testing.Regression_Testing
{
    [TestFixture, Order(2)]
    internal class SearchFundTest : BaseTestCase
    {
        // Variables declare
        [Obsolete]
        readonly ExtentReports rep = ExtReportgetInstance();
        private string? summaryTC;

        [Test, Category("Regression Testing")]
        public void TC001_SearchFund_Albourne_FundInfo()
        {
            #region Variables declare
            string managerNameAlbourne = FundDashboardPage.managerNameAlbourne;
            string fundNameAlbourne = FundDashboardPage.fundNameAlbourne;
            const string sourceIcon = "A";
            string fundType = "Hedge Fund";
            string inceptionDate = @"11/01/1990";
            //string fundStatus = "Closed";
            //string latestActualValue;
            //string redemptionNotedays = "60";
            //string hardLockmonths = "";
            //string softLockmonths = "24";
            //string earlyRedemptionFee = "";
            //string gate = "No";
            //string gatePerc = "";
            //string percentofNAVAvailable = "";
            string sidepocket = "No";
            string additionalNotesOnLiquidity = "Quarterly classes may be redeemed upon 60 days' notice and are generally subject to redemption charges ranging from 5%7% if aggregate net redemptions (excluding those made by the Manager's affiliates) exceed 3% and investors exceed the quarterly redemption allowance. Committed classes may be redeemed upon 90 days' notice and are not subject to redemption charges and offer rolling 24 months liquidity. Annual appreciation can be redeemed at December 31 each year.";
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Search Fund (Albourne) Test - TC001");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Log into the application
                LoginAction.Instance.LoginSiteNoGodaddy(); // LoginSite();

                #region Search the 1st Fund (citadel)
                // Search a Fund - Source = Albourne
                FundDashboardAction.Instance.InputNameToSearchFund(10, "cita")
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResultsWithItemSource(managerNameAlbourne, sourceIcon))
                                            .ClickFundNameReturnOfResults(10, managerNameAlbourne, sourceIcon)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResultsWithItemSource(fundNameAlbourne, sourceIcon))
                                            .ClickFundNameReturnOfResults(10, fundNameAlbourne, sourceIcon)
                                            .WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);

                // Verify all informations from Fund Info
                verifyPoint = fundNameAlbourne == FundDashboardAction.Instance.FundNameTitleGetText(10);
                verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Name (Albourne) is shown correctly after searching: '" + fundNameAlbourne + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = fundType == FundDashboardAction.Instance.FundInfoGetText(10, "Fund Type");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Type (Albourne) is shown correctly after searching: '" + fundType + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = managerNameAlbourne == FundDashboardAction.Instance.FundInfoGetText(10, "Manager Name");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Manager Name (Albourne) is shown correctly after searching: '" + managerNameAlbourne + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = inceptionDate == FundDashboardAction.Instance.FundInfoGetText(10, "Inception Date");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Inception Date (Albourne) is shown correctly after searching: '" + inceptionDate + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = fundStatus == FundDashboardAction.Instance.FundInfoGetText(10, "Fund Status");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Status (Albourne) is shown correctly after searching: '" + fundStatus + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                string label = "Latest Actual Value";
                verifyPoint = label == FundDashboardAction.Instance.FundInfoGetLabel(10, "Latest Actual Value");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Latest Actual Value label (Albourne) is shown correctly after searching: '" + label + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                label = "Redemption Note (days)";
                verifyPoint = label == FundDashboardAction.Instance.FundInfoGetLabel(10, "Redemption Note (days)");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Redemption Note (days) label (Albourne) is shown correctly after searching: '" + label + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                label = "Hard Lock (months)";
                verifyPoint = label == FundDashboardAction.Instance.FundInfoGetLabel(10, "Hard Lock (months)");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Hard Lock (months) label (Albourne) is shown correctly after searching: '" + label + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                label = "Soft Lock (months)";
                verifyPoint = label == FundDashboardAction.Instance.FundInfoGetLabel(10, "Soft Lock (months)");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Soft Lock (months) label (Albourne) is shown correctly after searching: '" + label + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                label = "Early Redemption Fee";
                verifyPoint = label == FundDashboardAction.Instance.FundInfoGetLabel(10, "Early Redemption Fee");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Early Redemption Fee label (Albourne) is shown correctly after searching: '" + label + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                label = "Gate";
                verifyPoint = label == FundDashboardAction.Instance.FundInfoGetLabel(10, "Gate");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Gate label (Albourne) is shown correctly after searching: '" + label + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                label = "Gate %";
                verifyPoint = label == FundDashboardAction.Instance.FundInfoGetLabel(10, "Gate %");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Gate % label (Albourne) is shown correctly after searching: '" + label + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                label = "Percent of NAV Available";
                verifyPoint = label == FundDashboardAction.Instance.FundInfoGetLabel(10, "Percent of NAV Available");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Percent of NAV Available label (Albourne) is shown correctly after searching: '" + label + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                //latestActualValue = "47349000000";
                //verifyPoint = latestActualValue == FundDashboardAction.Instance.FundInfoGetText(10, "Latest Actual Value");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Latest Actual Value (Albourne) is shown correctly after searching: '" + latestActualValue + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = redemptionNotedays == FundDashboardAction.Instance.FundInfoGetText(10, "Redemption Note (days)");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Redemption Note (days) (Albourne) is shown correctly after searching: '" + redemptionNotedays + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = hardLockmonths == FundDashboardAction.Instance.FundInfoGetText(10, "Hard Lock (months)");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Hard Lock (months) (Albourne) is shown correctly after searching: '" + hardLockmonths + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = softLockmonths == FundDashboardAction.Instance.FundInfoGetText(10, "Soft Lock (months)");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Soft Lock (months) (Albourne) is shown correctly after searching: '" + softLockmonths + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = earlyRedemptionFee == FundDashboardAction.Instance.FundInfoGetText(10, "Early Redemption Fee");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Early Redemption Fee (Albourne) is shown correctly after searching: '" + earlyRedemptionFee + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = gate == FundDashboardAction.Instance.FundInfoGetText(10, "Gate");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Gate (Albourne) is shown correctly after searching: '" + gate + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = gatePerc == FundDashboardAction.Instance.FundInfoGetText(10, "Gate %");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Gate % (Albourne) is shown correctly after searching: '" + gatePerc + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = percentofNAVAvailable == FundDashboardAction.Instance.FundInfoGetText(10, "Percent of NAV Available");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Percent of NAV Available (Albourne) is shown correctly after searching: '" + percentofNAVAvailable + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = sidepocket == FundDashboardAction.Instance.FundInfoGetText(10, "Sidepocket?");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Sidepocket? (Albourne) is shown correctly after searching: '" + sidepocket + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = additionalNotesOnLiquidity == FundDashboardAction.Instance.FundInfoGetText(10, "Additional Notes on Liquidity");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Additional Notes on Liquidity (Albourne) is shown correctly after searching:\n<br/> '" + additionalNotesOnLiquidity + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                #endregion

                #region Search the 2nd Fund (melvin)
                // Search another Fund (2nd Fund) - Source = Albourne (melvin)
                managerNameAlbourne = "Melvin Capital Management LP";
                fundNameAlbourne = "Melvin Capital Fund";
                fundType = "Hedge Fund";
                inceptionDate = "10/20/2014";
                //fundStatus = "In Liquidation"; // Closed, Hard
                //latestActualValue = "3385700000";
                //redemptionNotedays = "90";
                //hardLockmonths = "";
                //softLockmonths = "";
                //earlyRedemptionFee = "";
                //gate = "";
                //gatePerc = "";
                //percentofNAVAvailable = "";
                sidepocket = "No";
                additionalNotesOnLiquidity = "Compulsory Withdrawal - The General Partner may require a compulsory withdrawal, in that any Limited Partner withdraw all or any portion of its Interest in the Fund for any or no reason upon 5 days prior written notice to such Limited Partner.";
                FundDashboardAction.Instance.InputNameToSearchFund(10, "melvin")
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(managerNameAlbourne))
                                            .ClickFundNameReturnOfResults(10, managerNameAlbourne)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(fundNameAlbourne))
                                            .ClickFundNameReturnOfResults(10, fundNameAlbourne)
                                            .WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);

                // Verify all informations from Fund Info
                verifyPoint = fundNameAlbourne == FundDashboardAction.Instance.FundNameTitleGetText(10);
                verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Name (Albourne) is shown correctly after searching Fund 2: '" + fundNameAlbourne + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = fundType == FundDashboardAction.Instance.FundInfoGetText(10, "Fund Type");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Type (Albourne) is shown correctly after searching Fund 2: '" + fundType + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = managerNameAlbourne == FundDashboardAction.Instance.FundInfoGetText(10, "Manager Name");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Manager Name (Albourne) is shown correctly after searching Fund 2: '" + managerNameAlbourne + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = inceptionDate == FundDashboardAction.Instance.FundInfoGetText(10, "Inception Date");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Inception Date (Albourne) is shown correctly after searching Fund 2: '" + inceptionDate + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = fundStatus == FundDashboardAction.Instance.FundInfoGetText(10, "Fund Status");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Status (Albourne) is shown correctly after searching Fund 2: '" + fundStatus + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = latestActualValue == FundDashboardAction.Instance.FundInfoGetText(10, "Latest Actual Value");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Latest Actual Value (Albourne) is shown correctly after searching Fund 2: '" + latestActualValue + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = redemptionNotedays == FundDashboardAction.Instance.FundInfoGetText(10, "Redemption Note (days)");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Redemption Note (days) (Albourne) is shown correctly after searching Fund 2: '" + redemptionNotedays + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = hardLockmonths == FundDashboardAction.Instance.FundInfoGetText(10, "Hard Lock (months)");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Hard Lock (months) (Albourne) is shown correctly after searching Fund 2: '" + hardLockmonths + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = softLockmonths == FundDashboardAction.Instance.FundInfoGetText(10, "Soft Lock (months)");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Soft Lock (months) (Albourne) is shown correctly after searching Fund 2: '" + softLockmonths + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = earlyRedemptionFee == FundDashboardAction.Instance.FundInfoGetText(10, "Early Redemption Fee");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Early Redemption Fee (Albourne) is shown correctly after searching Fund 2: '" + earlyRedemptionFee + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = gate == FundDashboardAction.Instance.FundInfoGetText(10, "Gate");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Gate (Albourne) is shown correctly after searching Fund 2: '" + gate + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = gatePerc == FundDashboardAction.Instance.FundInfoGetText(10, "Gate %");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Gate % (Albourne) is shown correctly after searching Fund 2: '" + gatePerc + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = percentofNAVAvailable == FundDashboardAction.Instance.FundInfoGetText(10, "Percent of NAV Available");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Percent of NAV Available (Albourne) is shown correctly after searching Fund 2: '" + percentofNAVAvailable + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = sidepocket == FundDashboardAction.Instance.FundInfoGetText(10, "Sidepocket?");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Sidepocket? (Albourne) is shown correctly after searching Fund 2: '" + sidepocket + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = additionalNotesOnLiquidity == FundDashboardAction.Instance.FundInfoGetText(10, "Additional Notes on Liquidity");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Additional Notes on Liquidity (Albourne) is shown correctly after searching Fund 2:\n<br/> '" + additionalNotesOnLiquidity + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                #endregion

                #region Search the 3rd Fund (oasis Japan)
                /*
                // Search another Fund (3rd Fund) - Source = Albourne (Oasis Japan)
                managerNameAlbourne = "Oasis Management (Hong Kong)";
                fundNameAlbourne = "Oasis Japan Strategic Fund";
                fundType = "Hedge Fund";
                inceptionDate = "09/14/2018";
                //fundStatus = "Open";
                //latestActualValue = "875050358";
                //redemptionNotedays = "90";
                //hardLockmonths = "36";
                //softLockmonths = "";
                //earlyRedemptionFee = "";
                //gate = "";
                //gatePerc = "";
                //percentofNAVAvailable = "";
                sidepocket = "No";
                additionalNotesOnLiquidity = "";
                FundDashboardAction.Instance.InputNameToSearchFund(10, "oasis management (Hong Kong)")
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(managerNameAlbourne))
                                            .ClickFundNameReturnOfResults(10, managerNameAlbourne)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(fundNameAlbourne))
                                            .ClickFundNameReturnOfResults(10, fundNameAlbourne)
                                            .WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);

                // Verify all informations from Fund Info
                verifyPoint = fundNameAlbourne == FundDashboardAction.Instance.FundNameTitleGetText(10);
                verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Name (Albourne) is shown correctly after searching Fund 3: '" + fundNameAlbourne + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = fundType == FundDashboardAction.Instance.FundInfoGetText(10, "Fund Type");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Type (Albourne) is shown correctly after searching Fund 3: '" + fundType + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = managerNameAlbourne == FundDashboardAction.Instance.FundInfoGetText(10, "Manager Name");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Manager Name (Albourne) is shown correctly after searching Fund 3: '" + managerNameAlbourne + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = inceptionDate == FundDashboardAction.Instance.FundInfoGetText(10, "Inception Date");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Inception Date (Albourne) is shown correctly after searching Fund 3: '" + inceptionDate + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = fundStatus == FundDashboardAction.Instance.FundInfoGetText(10, "Fund Status");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Status (Albourne) is shown correctly after searching Fund 3: '" + fundStatus + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = latestActualValue == FundDashboardAction.Instance.FundInfoGetText(10, "Latest Actual Value");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Latest Actual Value (Albourne) is shown correctly after searching Fund 3: '" + latestActualValue + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = redemptionNotedays == FundDashboardAction.Instance.FundInfoGetText(10, "Redemption Note (days)");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Redemption Note (days) (Albourne) is shown correctly after searching Fund 3: '" + redemptionNotedays + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = hardLockmonths == FundDashboardAction.Instance.FundInfoGetText(10, "Hard Lock (months)");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Hard Lock (months) (Albourne) is shown correctly after searching Fund 3: '" + hardLockmonths + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = softLockmonths == FundDashboardAction.Instance.FundInfoGetText(10, "Soft Lock (months)");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Soft Lock (months) (Albourne) is shown correctly after searching Fund 3: '" + softLockmonths + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = earlyRedemptionFee == FundDashboardAction.Instance.FundInfoGetText(10, "Early Redemption Fee");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Early Redemption Fee (Albourne) is shown correctly after searching Fund 3: '" + earlyRedemptionFee + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = gate == FundDashboardAction.Instance.FundInfoGetText(10, "Gate");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Gate (Albourne) is shown correctly after searching Fund 3: '" + gate + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = gatePerc == FundDashboardAction.Instance.FundInfoGetText(10, "Gate %");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Gate % (Albourne) is shown correctly after searching Fund 3: '" + gatePerc + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = percentofNAVAvailable == FundDashboardAction.Instance.FundInfoGetText(10, "Percent of NAV Available");
                //verifyPoints.Add(summaryTC = "Verify Fund Info - Percent of NAV Available (Albourne) is shown correctly after searching Fund 3: '" + percentofNAVAvailable + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = sidepocket == FundDashboardAction.Instance.FundInfoGetText(10, "Sidepocket?");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Sidepocket? (Albourne) is shown correctly after searching Fund 3: '" + sidepocket + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = additionalNotesOnLiquidity == FundDashboardAction.Instance.FundInfoGetText(10, "Additional Notes on Liquidity");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Additional Notes on Liquidity (Albourne) is shown correctly after searching Fund 3:\n<br/> '" + additionalNotesOnLiquidity + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC); */
                #endregion
            }
            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Warning
                ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong! Please check console log.");
            }
            #endregion
        }

        [Test, Category("Regression Testing")]
        public void TC002_SearchFund_Manual_DataStatus()
        {
            #region Variables declare
            string urlInstance = LoginPage.url;
            const string exposure = "Exposure";
            const string fundReturns = "Fund Returns";
            const string fundAUM = "Fund AUM";
            const string rowCounts = "Row Counts";
            const string startDate = "Start Date";
            const string endDate = "End Date";
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Search Fund (Manual) Test - TC002");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Log into the application
                LoginAction.Instance.LoginSiteNoGodaddy(urlInstance); // LoginSite(urlInstance);

                #region Search a Fund Manual
                // Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
                string? inputSearch = null;
                string? managerNameManual = null;
                string? fundNameManual = null;
                string? exposureNRowCountVal = null;
                string? exposureStartDateVal = null;
                string? exposureEndDateVal = null;
                string? fundReturnsRowCountsVal = null;
                string? fundReturnsStartDateVal = null;
                string? fundReturnsEndDateVal = null;
                string? fundAUMRowCountsVal = null;
                string? fundAUMStartDateVal = null;
                string? fundAUMEndDateVal = null;
                if (urlInstance.Contains("sandbox"))
                {
                    inputSearch = "qa test 10";
                    managerNameManual = "QA Test 10";
                    fundNameManual = "Main of QA Test 10 edited";
                    exposureNRowCountVal = "2";
                    exposureStartDateVal = "04/30/2020";
                    exposureEndDateVal = "05/31/2021";
                    fundReturnsRowCountsVal = "314";
                    fundReturnsStartDateVal = "07/31/1995";
                    fundReturnsEndDateVal = "08/31/2021";
                    fundAUMRowCountsVal = "2";
                    fundAUMStartDateVal = "03/31/2021";
                    fundAUMEndDateVal = "04/30/2021";
                }
                if (urlInstance.Contains("conceptia"))
                {
                    inputSearch = "aqr";
                    managerNameManual = "AQR Capital QA";
                    fundNameManual = "AQR Delta Fund QA";
                    exposureNRowCountVal = "2";
                    exposureStartDateVal = "04/30/2011";
                    exposureEndDateVal = "05/31/2011";
                    fundReturnsRowCountsVal = "132";
                    fundReturnsStartDateVal = "05/31/2010";
                    fundReturnsEndDateVal = "04/30/2021";
                    fundAUMRowCountsVal = "12";
                    fundAUMStartDateVal = "01/31/2011";
                    fundAUMEndDateVal = "12/31/2011";
                }

                // Search a Fund - Source = Manual (QA Test 10)
                FundDashboardAction.Instance.InputNameToSearchFund(10, inputSearch)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(managerNameManual))
                                            .ClickFundNameReturnOfResults(10, managerNameManual)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(fundNameManual))
                                            .ClickFundNameReturnOfResults(10, fundNameManual)
                                            .WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);

                // Verify the Fund Name is shown correctly after searching
                verifyPoint = fundNameManual == FundDashboardAction.Instance.FundNameTitleGetText(10);
                verifyPoints.Add(summaryTC = "Verify the Fund Name (Manual) is shown correctly after searching: (" + fundNameManual + ")", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // Click on Data Status tab
                FundDashboardAction.Instance.ClickDataStatusTab(10);

                // verify all informations from Data Status - Exposure
                verifyPoint = FundDashboardAction.Instance.SectionTitleContentGetText(20, exposure, rowCounts) == exposureNRowCountVal;
                verifyPoints.Add(summaryTC = "Verify Data Status - " + exposure + " (Manual) " + rowCounts + " is shown correctly: " + exposureNRowCountVal + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = FundDashboardAction.Instance.SectionTitleContentGetText(10, exposure, startDate) == exposureStartDateVal;
                verifyPoints.Add(summaryTC = "Verify Data Status - " + exposure + " (Manual) " + startDate + " is shown correctly: " + exposureStartDateVal + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = FundDashboardAction.Instance.SectionTitleContentGetText(10, exposure, endDate) == exposureEndDateVal;
                verifyPoints.Add(summaryTC = "Verify Data Status - " + exposure + " (Manual) " + endDate + " is shown correctly: " + exposureEndDateVal + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // verify all informations from Data Status - Fund Returns
                verifyPoint = FundDashboardAction.Instance.SectionTitleContentGetText(10, fundReturns, rowCounts) == fundReturnsRowCountsVal;
                verifyPoints.Add(summaryTC = "Verify Data Status - " + fundReturns + " (Manual) " + rowCounts + " is shown correctly: " + fundReturnsRowCountsVal + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = FundDashboardAction.Instance.SectionTitleContentGetText(10, fundReturns, startDate) == fundReturnsStartDateVal;
                verifyPoints.Add(summaryTC = "Verify Data Status - " + fundReturns + " (Manual) " + startDate + " is shown correctly: " + fundReturnsStartDateVal + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = FundDashboardAction.Instance.SectionTitleContentGetText(10, fundReturns, endDate) == fundReturnsEndDateVal;
                verifyPoints.Add(summaryTC = "Verify Data Status - " + fundReturns + " (Manual) " + endDate + " is shown correctly: " + fundReturnsEndDateVal + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                // verify all informations from Data Status - Fund AUM
                verifyPoint = FundDashboardAction.Instance.SectionTitleContentGetText(10, fundAUM, rowCounts) == fundAUMRowCountsVal;
                verifyPoints.Add(summaryTC = "Verify Data Status - " + fundAUM + " (Manual) " + rowCounts + " is shown correctly: "+ fundAUMRowCountsVal + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = FundDashboardAction.Instance.SectionTitleContentGetText(10, fundAUM, startDate) == fundAUMStartDateVal;
                verifyPoints.Add(summaryTC = "Verify Data Status - " + fundAUM + " (Manual) " + startDate + " is shown correctly: " + fundAUMStartDateVal + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = FundDashboardAction.Instance.SectionTitleContentGetText(10, fundAUM, endDate) == fundAUMEndDateVal;
                verifyPoints.Add(summaryTC = "Verify Data Status - " + fundAUM + " (Manual) " + endDate + " is shown correctly: " + fundAUMEndDateVal + "", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                #endregion
            }
            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Warning
                ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong! Please check console log.");
            }
            #endregion
        }

        [Test, Category("Regression Testing")]
        public void TC003_SearchFund_Cambridge()
        {
            #region Variables declare
            const string cambridgeFund = "VGO Capital Partners"; // (firm)
            const string sourceIcon = "C";
            const string yearFirmFounded = "2013";
            const string strategyHeadquarter = "Heathcoat House, 2nd Floor 20 Saville Row London W1S 3PR United Kingdom";
            const string businessContact = "Ivan Capriello General Partner & Co-Founder";
            const string contactPhone = "44.20.7087.9070";
            const string contactEmail = "";
            const string fundName1 = "VGO Special Situations Fund I";
            const string fundName2 = "VGO Special Situations Fund II";
            const string strategy = "European Mezzanine";
            const string vintageYear = "2013";
            const string assetClass = "Subordinated Capital";
            const string investmentStage = "Mezzanine Debt";
            const string industryFocus = "Multi Industry";
            const string geographicFocus = "Europe DEV Cross-Region General";
            const string fundSizeM1 = "";
            const string fundSizeM2 = "288 USD";
            //const string currency = "USD";
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Search Fund (Cambridge) Test - TC003");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Log into the application
                LoginAction.Instance.LoginSiteNoGodaddy();

                // Search a Cambridge Fund (VGO) (KS-455 the Firm should be searched in the search bar)
                FundDashboardAction.Instance.InputNameToSearchFund(10, "vgo ")
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResultsWithItemSource(cambridgeFund, sourceIcon))
                                            .ClickFundNameReturnOfResults(10, cambridgeFund, sourceIcon); Thread.Sleep(1000);

                // Check if loading icon shows then wait for it to disappear
                if (FundDashboardAction.Instance.IsElementPresent(FundDashboardPage.loadingIcon))
                {
                    FundDashboardAction.Instance.WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);
                }

                #region Verify all informations from Manager Info (tab)
                // Click on Manager Info tab
                Thread.Sleep(1000);
                FundDashboardAction.Instance.ClickManagerInfoTab(10);

                // Verify all informations from Manager Info (tab)
                verifyPoint = cambridgeFund == FundDashboardAction.Instance.FundNameTitleGetText(10);
                verifyPoints.Add(summaryTC = "Verify Manager Info - Firm (Cambridge Fund) is shown correctly after searching: '" + cambridgeFund + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = yearFirmFounded == FundDashboardAction.Instance.FundInfoGetText(10, FundDashboardPage.yearFirmFounded);
                verifyPoints.Add(summaryTC = "Verify Fund Info - Year Firm Founded (Cambridge Fund) is shown correctly after searching: '" + yearFirmFounded + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = strategyHeadquarter == FundDashboardAction.Instance.FundInfoGetText(10, FundDashboardPage.strategyHeadquarter);
                verifyPoints.Add(summaryTC = "Verify Fund Info - Strategy Headquarter (Cambridge Fund) is shown correctly after searching: '" + strategyHeadquarter + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = businessContact == FundDashboardAction.Instance.FundInfoGetText(10, FundDashboardPage.businessContact);
                verifyPoints.Add(summaryTC = "Verify Fund Info - Business Contact (Cambridge Fund) is shown correctly after searching: '" + businessContact + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = contactPhone == FundDashboardAction.Instance.FundInfoGetText(10, FundDashboardPage.contactPhone);
                verifyPoints.Add(summaryTC = "Verify Fund Info - Contact Phone (Cambridge Fund) is shown correctly after searching: '" + contactPhone + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = contactEmail == FundDashboardAction.Instance.FundInfoGetText(10, FundDashboardPage.contactEmail);
                verifyPoints.Add(summaryTC = "Verify Fund Info - Contact Email (Cambridge Fund) is shown correctly after searching: '" + contactEmail + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                #endregion

                #region Verify all informations from Fund List (tab)
                // Click on Fund List tab
                FundDashboardAction.Instance.ClickFundListTab(10)
                                            .WaitForElementVisible(10, FundDashboardPage.expandAllBtn)
                                            .ClickExpandAllButton(10); // --> CLick Expand All button

                // Verify all informations from Fund List (tab)
                verifyPoint = strategy == FundDashboardAction.Instance.ListStrategiesGetText(10, 1);
                verifyPoints.Add(summaryTC = "Verify Fund List - Strategy (Cambridge Fund) is shown correctly after searching: '" + strategy + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = fundName1 == FundDashboardAction.Instance.ListFundsNameGetText(10, 1, 1);
                verifyPoints.Add(summaryTC = "Verify Fund List - Fund name 1 (Cambridge Fund) is shown correctly after searching: '" + fundName1 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = fundName2 == FundDashboardAction.Instance.ListFundsNameGetText(10, 1, 2);
                verifyPoints.Add(summaryTC = "Verify Fund List - Fund name 2 (Cambridge Fund) is shown correctly after searching: '" + fundName2 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = vintageYear == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.vintageYear)
                           && vintageYear == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 2, FundDashboardPage.vintageYear);
                verifyPoints.Add(summaryTC = "Verify Fund List - Vintage Year (Cambridge Fund) is shown correctly after searching: '" + vintageYear + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = assetClass == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.assetClass)
                           && assetClass == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 2, FundDashboardPage.assetClass);
                verifyPoints.Add(summaryTC = "Verify Fund List - Asset Class (Cambridge Fund) is shown correctly after searching: '" + assetClass + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = investmentStage == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.investmentStage)
                           && investmentStage == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 2, FundDashboardPage.investmentStage);
                verifyPoints.Add(summaryTC = "Verify Fund List - Investment Stage (Cambridge Fund) is shown correctly after searching: '" + investmentStage + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = industryFocus == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.industryFocus)
                           && industryFocus == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 2, FundDashboardPage.industryFocus);
                verifyPoints.Add(summaryTC = "Verify Fund List - Industry Focus (Cambridge Fund) is shown correctly after searching: '" + industryFocus + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = geographicFocus == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.geographicFocus)
                           && geographicFocus == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 2, FundDashboardPage.geographicFocus);
                verifyPoints.Add(summaryTC = "Verify Fund List - Geographic Focus (Cambridge Fund) is shown correctly after searching: '" + geographicFocus + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = fundSizeM1 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.fundSizeM)
                           && fundSizeM2 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 2, FundDashboardPage.fundSizeM);
                verifyPoints.Add(summaryTC = "Verify Fund List - fund Size (M) (Cambridge Fund) is shown correctly after searching: fsm1='" + fundSizeM1 + "', fsm2='" + fundSizeM2 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                //verifyPoint = currency == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.currency)
                //           && currency == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 2, FundDashboardPage.currency);
                //verifyPoints.Add(summaryTC = "Verify Fund List - Currency (Cambridge Fund) is shown correctly after searching: '" + currency + "'", verifyPoint);
                //ExtReportResult(verifyPoint, summaryTC);
                #endregion
            }
            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Warning
                ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong! Please check console log.");
            }
            #endregion
        }

        [Test, Category("Regression Testing")]
        public void TC004_SearchFund_Cambridge_manual()
        {
            #region Variables declare
            string urlInstance = LoginPage.url;
            string? inputSearch = null;
            const string sourceIcon = "M";
            // Manager Info
            string? cambridgeFund = null; // --> firm
            string? yearFirmFounded = null;
            string? strategyHeadquarter = null;
            string? businessContact = null;
            string? contactPhone = null;
            string? contactEmail = null;

            // Fund Info
            string? fundName1 = null, fundName2 = null, fundName3 = null;
            string? strategy1 = null, strategy2 = null, strategy3 = null;
            string? vintageYear1 = null, vintageYear2 = null, vintageYear3 = null;
            string? assetClass1 = null, assetClass2 = null, assetClass3 = null;
            string? investmentStage1 = null, investmentStage2 = null, investmentStage3 = null;
            string? industryFocus1 = null, industryFocus2 = null, industryFocus3 = null;
            string? geographicFocus1 = null, geographicFocus2 = null, geographicFocus3 = null;
            string? fundSizeM1 = null, fundSizeM2 = null, fundSizeM3 = null;
            string? currency1 = null, currency2 = null, currency3 = null;
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Search Fund (Cambridge Manual) Test - TC004");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Log into the application
                LoginAction.Instance.LoginSiteNoGodaddy();

                // Check if the data of Sandbox or Staging (Conceptia) site then verify data (on that site)
                if (urlInstance.Contains("sandbox")) 
                {
                    inputSearch = "priman F";
                    cambridgeFund = "PriMan F03"; // --> firm
                    yearFirmFounded = "3102";
                    strategyHeadquarter = "3103";
                    businessContact = "306";
                    contactPhone = "308";
                    contactEmail = "307";
                    fundName1 = "Fund 1 of PriMan F03"; fundName2 = "Fund 2 of PriMan F03"; fundName3 = "Fund 3 of PriMan F03";
                    strategy1 = "3101"; strategy2 = "3201"; strategy3 = "3301";
                    vintageYear1 = "3102"; vintageYear2 = "3202"; vintageYear3 = "3302";
                    assetClass1 = "77"; assetClass2 = "11"; assetClass3 = "1";
                    investmentStage1 = "88"; investmentStage2 = "22"; investmentStage3 = "2";
                    industryFocus1 = "99"; industryFocus2 = "33"; industryFocus3 = "3";
                    geographicFocus1 = "111"; geographicFocus2 = "44"; geographicFocus3 = "4";
                    fundSizeM1 = "222"; fundSizeM2 = "55"; fundSizeM3 = "5";
                    currency1 = "333"; currency2 = "66"; currency3 = "6";
                }
                if (urlInstance.Contains("conceptia")) 
                {
                    // Manager Info
                    inputSearch = "prim";
                    cambridgeFund = "Priman Sta01"; // --> firm
                    yearFirmFounded = "2020";
                    strategyHeadquarter = "SHQA 101";
                    businessContact = "BCON 01";
                    contactPhone = "123456789";
                    contactEmail = "BEMA01@test.com";

                    // Fund Info
                    fundName1 = "Fund 1 of Firm Sta01"; fundName2 = "Fund 2 of Firm Sta01"; fundName3 = "Fund 3 of Firm Sta01";
                    strategy1 = "STRA 101"; strategy2 = "STRA 201"; strategy3 = "STRA 301";
                    vintageYear1 = "2020"; vintageYear2 = "2021"; vintageYear3 = "2022";
                    assetClass1 = "AC 101"; assetClass2 = "AC 201"; assetClass3 = "AC 301";
                    investmentStage1 = "IS 101"; investmentStage2 = "IS 201"; investmentStage3 = "IS 301";
                    industryFocus1 = "IF 101"; industryFocus2 = "IF 201"; industryFocus3 = "IF 301";
                    geographicFocus1 = "GF 101"; geographicFocus2 = "GF 201"; geographicFocus3 = "GF 301";
                    fundSizeM1 = "101"; fundSizeM2 = "201"; fundSizeM3 = "301";
                    currency1 = "USD"; currency2 = "USD"; currency3 = "USD";
                }

                // Search a Cambridge Fund (Manual --> added by user) (KS-455 the Firm should be searched in the search bar)
                FundDashboardAction.Instance.InputNameToSearchFund(10, inputSearch)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResultsWithItemSource(cambridgeFund, sourceIcon))
                                            .ClickFundNameReturnOfResults(10, cambridgeFund, sourceIcon); Thread.Sleep(1000);

                // Check if loading icon shows then wait for it to disappear
                if (FundDashboardAction.Instance.IsElementPresent(FundDashboardPage.loadingIcon))
                {
                    FundDashboardAction.Instance.WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);
                }

                #region Verify all informations from Manager Info (tab)
                // Click on Manager Info tab
                Thread.Sleep(1000);
                FundDashboardAction.Instance.ClickManagerInfoTab(10);

                // Verify all informations from Manager Info (tab)
                verifyPoint = cambridgeFund == FundDashboardAction.Instance.FundNameTitleGetText(10);
                verifyPoints.Add(summaryTC = "Verify Manager Info - Firm (Cambridge Fund - Manual) is shown correctly after searching: '" + cambridgeFund + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = yearFirmFounded == FundDashboardAction.Instance.FundInfoGetText(10, FundDashboardPage.yearFirmFounded);
                verifyPoints.Add(summaryTC = "Verify Fund Info - firm (Cambridge Fund - Manual) is shown correctly after searching: '" + yearFirmFounded + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = strategyHeadquarter == FundDashboardAction.Instance.FundInfoGetText(10, FundDashboardPage.strategyHeadquarter);
                verifyPoints.Add(summaryTC = "Verify Fund Info - firm (Cambridge Fund - Manual) is shown correctly after searching: '" + strategyHeadquarter + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = businessContact == FundDashboardAction.Instance.FundInfoGetText(10, FundDashboardPage.businessContact);
                verifyPoints.Add(summaryTC = "Verify Fund Info - firm (Cambridge Fund - Manual) is shown correctly after searching: '" + businessContact + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = contactPhone == FundDashboardAction.Instance.FundInfoGetText(10, FundDashboardPage.contactPhone);
                verifyPoints.Add(summaryTC = "Verify Fund Info - firm (Cambridge Fund - Manual) is shown correctly after searching: '" + contactPhone + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = contactEmail == FundDashboardAction.Instance.FundInfoGetText(10, FundDashboardPage.contactEmail);
                verifyPoints.Add(summaryTC = "Verify Fund Info - firm (Cambridge Fund - Manual) is shown correctly after searching: '" + contactEmail + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                #endregion

                #region Verify all informations from Fund List (tab)
                // Click on Fund List tab
                FundDashboardAction.Instance.ClickFundListTab(10)
                                            .WaitForElementVisible(10, FundDashboardPage.expandAllBtn)
                                            .ClickExpandAllButton(10); // --> CLick Expand All button

                // Verify all informations from Fund List (tab)
                verifyPoint = strategy1 == FundDashboardAction.Instance.ListStrategiesGetText(10, 1);
                verifyPoints.Add(summaryTC = "Verify Fund List - Strategy 1 (Cambridge Fund - Manual) is shown correctly after searching: '" + strategy1 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = strategy2 == FundDashboardAction.Instance.ListStrategiesGetText(10, 2);
                verifyPoints.Add(summaryTC = "Verify Fund List - Strategy 2 (Cambridge Fund - Manual) is shown correctly after searching: '" + strategy2 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = strategy3 == FundDashboardAction.Instance.ListStrategiesGetText(10, 3);
                verifyPoints.Add(summaryTC = "Verify Fund List - Strategy 3 (Cambridge Fund - Manual) is shown correctly after searching: '" + strategy3 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = fundName1 == FundDashboardAction.Instance.ListFundsNameGetText(10, 1, 1);
                verifyPoints.Add(summaryTC = "Verify Fund List - Fund name 1 (Cambridge Fund - Manual) is shown correctly after searching: '" + fundName1 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = fundName2 == FundDashboardAction.Instance.ListFundsNameGetText(10, 2, 1);
                verifyPoints.Add(summaryTC = "Verify Fund List - Fund name 2 (Cambridge Fund - Manual) is shown correctly after searching: '" + fundName2 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = fundName3 == FundDashboardAction.Instance.ListFundsNameGetText(10, 3, 1);
                verifyPoints.Add(summaryTC = "Verify Fund List - Fund name 3 (Cambridge Fund - Manual) is shown correctly after searching: '" + fundName3 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = vintageYear1 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.vintageYear)
                           && vintageYear2 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 2, 1, FundDashboardPage.vintageYear)
                           && vintageYear3 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 3, 1, FundDashboardPage.vintageYear);
                verifyPoints.Add(summaryTC = "Verify Fund List - Vintage Year (Cambridge Fund) is shown correctly after searching: VY1='" + vintageYear1 + "', VY2='" + vintageYear2 + "', VY3='" + vintageYear3 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = assetClass1 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.assetClass)
                           && assetClass2 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 2, 1, FundDashboardPage.assetClass)
                           && assetClass3 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 3, 1, FundDashboardPage.assetClass);
                verifyPoints.Add(summaryTC = "Verify Fund List - Asset Class (Cambridge Fund) is shown correctly after searching: AC1='" + assetClass1 + "', AC2='" + assetClass2 + "', AC3='" + assetClass3 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = investmentStage1 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.investmentStage)
                           && investmentStage2 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 2, 1, FundDashboardPage.investmentStage)
                           && investmentStage3 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 3, 1, FundDashboardPage.investmentStage);
                verifyPoints.Add(summaryTC = "Verify Fund List - Investment Stage (Cambridge Fund) is shown correctly after searching: IS1='" + investmentStage1 + "', IS2='" + investmentStage2 + "', IS3='" + investmentStage3 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = industryFocus1 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.industryFocus)
                           && industryFocus2 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 2, 1, FundDashboardPage.industryFocus)
                           && industryFocus3 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 3, 1, FundDashboardPage.industryFocus);
                verifyPoints.Add(summaryTC = "Verify Fund List - Industry Focus (Cambridge Fund) is shown correctly after searching: IF1='" + industryFocus1 + "', IF2='" + industryFocus2 + "', IF3='" + industryFocus3 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = geographicFocus1 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.geographicFocus)
                           && geographicFocus2 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 2, 1, FundDashboardPage.geographicFocus)
                           && geographicFocus3 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 3, 1, FundDashboardPage.geographicFocus);
                verifyPoints.Add(summaryTC = "Verify Fund List - Geographic Focus (Cambridge Fund) is shown correctly after searching: GF1='" + geographicFocus1 + "', GF2='" + geographicFocus2 + "', GF3='" + geographicFocus3 + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = fundSizeM1 + " " + currency1 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 1, 1, FundDashboardPage.fundSizeM)
                           && fundSizeM2 + " " + currency2 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 2, 1, FundDashboardPage.fundSizeM)
                           && fundSizeM3 + " " + currency3 == FundDashboardAction.Instance.ListFundsInfoGetText(10, 3, 1, FundDashboardPage.fundSizeM);
                verifyPoints.Add(summaryTC = "Verify Fund List - Fund Size(M) (Cambridge Fund) is shown correctly after searching: FS1='" + fundSizeM1 + " "+currency1+"', FS2='" + fundSizeM2 + " "+currency2+"', FS3='" + fundSizeM3 + " "+currency3+"'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                #endregion
            }
            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Warning
                ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong! Please check console log.");
            }
            #endregion
        }

        [Test, Category("Regression Testing")]
        public void TC005_SearchFund_Solovis()
        {
            #region Variables declare
            const string managerName = "Laurion Capital Management LP";
            const string subManagerName = "";
            const string fundName = "Laurion Capital Ltd.";
            const string sourceIcon = "S";
            const string fundType = "Accounting Fund";
            const string inceptionDate = @"09/01/2005";  // --> KS-512
            const string ksInceptionDate = @"12/29/2020";  // --> KS-512
            const string fundStatus = "";
            //string latestActualValue = "194771097.21";
            //string redemptionNotedays = "";
            //string hardLockmonths = "";
            //string softLockmonths = "";
            //string earlyRedemptionFee = "";
            //string gate = "";
            //string gatePerc = "";
            //string percentofNAVAvailable = "";
            const string sidepocket = "No";
            const string additionalNotesOnLiquidity = "";
            const string trackingError = "0.12";
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Search Fund (Solovis) Test - TC005");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Log into the application
                LoginAction.Instance.LoginSiteNoGodaddy();

                // Search a Fund - Source = Solovis
                FundDashboardAction.Instance.InputNameToSearchFund(10, "laurion")
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResultsWithItemSource(managerName, sourceIcon))
                                            .ClickFundNameReturnOfResults(10, managerName, sourceIcon)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResultsWithItemSource(fundName, sourceIcon))
                                            .ClickFundNameReturnOfResults(10, fundName, sourceIcon)
                                            .WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);

                // Verify all informations from Fund Info
                verifyPoint = fundName == FundDashboardAction.Instance.FundNameTitleGetText(10);
                verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Name (Solovis) is shown correctly after searching: '" + fundName + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = fundType == FundDashboardAction.Instance.FundInfoGetText(10, "Fund Type");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Type (Solovis) is shown correctly after searching: '" + fundType + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = managerName == FundDashboardAction.Instance.FundInfoGetText(10, "Manager Name");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Manager Name (Solovis) is shown correctly after searching: '" + managerName + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = subManagerName == FundDashboardAction.Instance.FundInfoGetText(10, "Sub Manager Name");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Sub Manager Name (Solovis) is shown correctly after searching: '" + subManagerName + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = inceptionDate == FundDashboardAction.Instance.FundInfoGetText(10, "Inception Date");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Inception Date (Solovis) is shown correctly after searching: '" + inceptionDate + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = ksInceptionDate == FundDashboardAction.Instance.FundInfoGetText(10, "KS Inception Date"); // --> KS-512
                verifyPoints.Add(summaryTC = "Verify Fund Info - KS Inception Date (Solovis) is shown correctly after searching: '" + inceptionDate + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = fundStatus == FundDashboardAction.Instance.FundInfoGetText(10, "Fund Status");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Status (Solovis) is shown correctly after searching: '" + fundStatus + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = sidepocket == FundDashboardAction.Instance.FundInfoGetText(10, "Sidepocket?");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Sidepocket? (Solovis) is shown correctly after searching: '" + sidepocket + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = additionalNotesOnLiquidity == FundDashboardAction.Instance.FundInfoGetText(10, "Additional Notes on Liquidity");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Additional Notes on Liquidity (Solovis) is shown correctly after searching:\n<br/> '" + additionalNotesOnLiquidity + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = trackingError == FundDashboardAction.Instance.FundInfoGetText(10, "Tracking Error");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Sidepocket? (Solovis) is shown correctly after searching: '" + trackingError + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
            }
            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Warning
                ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong! Please check console log.");
            }
            #endregion
        }

        [Test, Category("Regression Testing")]
        public void TC006_SearchFund_Evestment()
        {
            #region Variables declare
            string managerName = "Advent Capital Management, LLC";
            const string sourceIcon = "E";
            string fundName = "Advent Global Partners";
            string fundType = "Hedge Funds";
            string inceptionDate = "08/01/2001";
            string sidepocket = "No";
            #endregion

            #region Workflow scenario
            //Verify steps, if getting issue show warning message
            test = rep.CreateTest("GenD KS - Search Fund (Evestment) Test - TC006");
            try
            {
                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Log into the application
                LoginAction.Instance.LoginSiteNoGodaddy();

                #region Search Alternatives Evestment Fund
                FundDashboardAction.Instance.InputNameToSearchFund(10, "advent Ca")
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResultsWithItemSource(managerName, sourceIcon))
                                            .ClickFundNameReturnOfResults(10, managerName, sourceIcon)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResultsWithItemSource(fundName, sourceIcon))
                                            .ClickFundNameReturnOfResults(10, fundName, sourceIcon)
                                            .WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);

                // Verify all informations from Fund Info
                verifyPoint = fundName == FundDashboardAction.Instance.FundNameTitleGetText(10);
                verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Name (Alt Evestment) is shown correctly after searching: '" + fundName + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = fundType == FundDashboardAction.Instance.FundInfoGetText(10, "Fund Type");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Type (Alt Evestment) is shown correctly after searching: '" + fundType + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = managerName == FundDashboardAction.Instance.FundInfoGetText(10, "Manager Name");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Manager Name (Alt Evestment) is shown correctly after searching: '" + managerName + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = inceptionDate == FundDashboardAction.Instance.FundInfoGetText(10, "Inception Date");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Inception Date (Alt Evestment) is shown correctly after searching: '" + inceptionDate + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = sidepocket == FundDashboardAction.Instance.FundInfoGetText(10, "Sidepocket?");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Sidepocket? (Alt Evestment) is shown correctly after searching: '" + sidepocket + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                #endregion

                #region Search Traditional Evestment Fund
                managerName = "John Hancock Investments";
                fundName = "John Hancock Multimanager Lifetime Portfolios 2035";
                FundDashboardAction.Instance.InputNameToSearchFund(10, fundName)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(managerName))
                                            .ClickFundNameReturnOfResults(10, managerName, sourceIcon)
                                            .WaitForElementVisible(10, FundDashboardPage.fundNameReturnOfResults(fundName))
                                            .ClickFundNameReturnOfResults(10, fundName)
                                            .WaitForLoadingIconToDisappear(10, FundDashboardPage.loadingIcon);

                // Verify all informations from Fund Info
                verifyPoint = fundName == FundDashboardAction.Instance.FundNameTitleGetText(10);
                verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Name (Traditional Evestment) is shown correctly after searching: '" + fundName + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                fundType = "Balanced/Multi-Asset";
                verifyPoint = fundType == FundDashboardAction.Instance.FundInfoGetText(10, "Fund Type");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Fund Type (Traditional Evestment) is shown correctly after searching: '" + fundType + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = managerName == FundDashboardAction.Instance.FundInfoGetText(10, "Manager Name");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Manager Name (Traditional Evestment) is shown correctly after searching: '" + managerName + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                inceptionDate = "10/30/2006";
                verifyPoint = inceptionDate == FundDashboardAction.Instance.FundInfoGetText(10, "Inception Date");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Inception Date (Traditional Evestment) is shown correctly after searching: '" + inceptionDate + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);

                verifyPoint = sidepocket == FundDashboardAction.Instance.FundInfoGetText(10, "Sidepocket?");
                verifyPoints.Add(summaryTC = "Verify Fund Info - Sidepocket? (Traditional Evestment) is shown correctly after searching: '" + sidepocket + "'", verifyPoint);
                ExtReportResult(verifyPoint, summaryTC);
                #endregion
            }
            catch (Exception exception)
            {
                // Print exception
                Console.WriteLine(exception);

                // Delete Manifest file
                LoginAction.Instance.DeleteAllFilesManifest();

                // Warning
                ExtReportResult(false, "Something wrong! Please check console log." + "<br/>" + exception);
                Assert.Inconclusive("Something wrong! Please check console log.");
            }
            #endregion
        }
    }
}