using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Playwright;


namespace QA.AutomationTests.SmokeTests
{


    [TestClass]
    [DoNotParallelize]
    public sealed class SmokeTests : TestBase
    {
        
        BenefitToBeCreated benefit = new BenefitToBeCreated { BenefitName = "stagingGitHubSmokeTest", Period = 2026, BenefitTerm = BenefitTerm.Indefinite };

        [TestMethod]
        public async Task SmokeTest()
        {            
            await Manager.CreateBenefit(page, benefit);
            await Manager.RenewBenefit(page, benefit);
        }

    }
}
