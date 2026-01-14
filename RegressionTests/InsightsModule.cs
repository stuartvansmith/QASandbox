using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Playwright;


namespace QA.AutomationTests.RegressionTests
{
    [TestClass]
    [DoNotParallelize]
    public class InsightsModule : TestBase
    {
        [TestMethod]
        public async Task BenefitSummaryReport()
        {
             
            await page.GetByText("Run Report").Nth(0).ClickAsync();

            var filePath = await TestHelper.DownloadAndVerifyAsync(
                page,
                async () => await page.GetByText("Export Report").First.ClickAsync(),
                downloadsDir: "test-downloads",
                timeoutMs: 30000
            );

            Console.WriteLine($"Downloaded file: {filePath}");
        }
        [TestMethod]
        public async Task BenefitDetailReport()
        {

            await page.GetByText("Run Report").Nth(1).ClickAsync();

            await page.GetByText("Benefit family").ClickAsync();
            await page.GetByText("Risk").ClickAsync();

            await page.GetByText("Benefit type", new() { Exact = true }).ClickAsync();
            await page.GetByText("Critical illness").ClickAsync();

            

            var filePath = await TestHelper.DownloadAndVerifyAsync(
                page,
                async () => await page.GetByText("Export Report").First.ClickAsync(),
                downloadsDir: "test-downloads",
                timeoutMs: 30000
            );

            Console.WriteLine($"Downloaded file: {filePath}");
        }
        

        public override async Task Navigate() 
        {
            await page.GotoAsync("https://staging.originbenefits.ai/login");
            await TestHelper.FinishLogin(page, "2026");
            await page.GetByRole(AriaRole.Link, new() { Name = "lightbulb_circle Insights" }).ClickAsync();

        }
    }
}
