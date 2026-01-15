using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Playwright;


namespace QA.AutomationTests.SmokeTests
{


    [TestClass]
    [DoNotParallelize]
    public sealed class SmokeTests : TestBase
    {
        
        BenefitToBeCreated benefit = new BenefitToBeCreated { BenefitName = "GitHubSmokeTest", Period = 2026 };

        [TestMethod]
        public async Task SmokeTestStaging()
        {            
            await Manager.CreateBenefit(page, benefit);
            await Manager.RenewBenefit(page, benefit);
        }
        [TestMethod]
        public async Task SmokeTestStagingIntendedFail()
        {
            Assert.IsTrue(false);
        }
    }
}
