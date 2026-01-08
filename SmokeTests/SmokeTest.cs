using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Playwright;


namespace QA.AutomationTests.SmokeTests
{


    [TestClass]
    [DoNotParallelize]
    public sealed class SmokeTests : TestBase
    {
        
        BenefitToBeCreated benefit = new BenefitToBeCreated { BenefitName = "FromGitHubXX", Period = 2026 };

        [TestMethod]
        public async Task SmokeTestStaging()
        {

            await page.GotoAsync("https://staging.originbenefits.ai/login");
            await TestHelper.FinishLogin(page, benefit.Period.ToString());
            await Manager.CreateBenefit(page, benefit);
            await Manager.RenewBenefit(page, benefit);
           
        }
        [TestMethod]
        public async Task SmokeTestDemo()
        {

            await page.GotoAsync("https://demo.originbenefits.ai/login");
            await TestHelper.FinishLogin(page, benefit.Period.ToString());
            await Manager.CreateBenefit(page, benefit);
            await Manager.RenewBenefit(page, benefit);

        }
        [TestMethod]
        public async Task SmokeTestLive()
        {

            await page.GotoAsync("https://web-origin-live.azurewebsites.net/login");
            await TestHelper.FinishLogin(page, benefit.Period.ToString());
            await Manager.CreateBenefit(page, benefit);
            await Manager.RenewBenefit(page, benefit);

        }
    }
}
