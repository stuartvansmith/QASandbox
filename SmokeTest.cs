using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Playwright;


namespace QA.AutomationTests
{


    [TestClass]
    public sealed class SmokeTests : TestBase
    {
        
        BenefitToBeCreated benefit = new BenefitToBeCreated { BenefitName = "FromGitHub" };

        [TestMethod]
        public async Task SmokeTestStaging()
        {

            await page.GotoAsync("https://staging.originbenefits.ai/login");
            await TestHelper.FinishLogin(page);
            await Manager.CreateBenefit(page, benefit);
           // await Manager.RenewBenefit(page, benefit);
           
        }
        [TestMethod]
        public async Task SmokeTestDemo()
        {

            await page.GotoAsync("https://demo.originbenefits.ai/login");
            await TestHelper.FinishLogin(page);
            await Manager.CreateBenefit(page, benefit);
            await Manager.RenewBenefit(page, benefit);

        }
        [TestMethod]
        public async Task SmokeTestLive()
        {

            await page.GotoAsync("https://web-origin-live.azurewebsites.net/login");
            await TestHelper.FinishLogin(page);
            await Manager.CreateBenefit(page, benefit);
            await Manager.RenewBenefit(page, benefit);

        }
    }
}
