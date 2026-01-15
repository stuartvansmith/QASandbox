using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Playwright;
using System.Text.RegularExpressions;


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
            
            await Task.Delay(1000);
            await page.GetByText("Benefit family", new() { Exact = true }).ClickAsync();
            await Task.Delay(1000);
            await page.GetByText("Risk").ClickAsync();


            await page.GetByText("Benefit type", new() { Exact = true }).ClickAsync();
            await Task.Delay(1000);
            await page.GetByText("Critical illness").ClickAsync();
            await Task.Delay(1000);

            var text = page.GetByText("Benefit detail report (Critical illness)", new() { Exact = true });

            await text.WaitForAsync(new()
            {
                State = WaitForSelectorState.Visible,
                Timeout = 10000
            });
            

            var filePath = await TestHelper.DownloadAndVerifyAsync(
                page,
                async () => await page.GetByText("Export Report").First.ClickAsync(),
                downloadsDir: "test-downloads",
                timeoutMs: 30000
            );

            Console.WriteLine($"Downloaded file: {filePath}");
        }

        //[TestCleanup]
        //public override async Task TestCleanUp()
        //{
        //    await page.GetByRole(AriaRole.Link, new() { Name = "lightbulb_circle Insights" }).ClickAsync();
        //}

        [TestInitialize]
        public override async Task TestSetup() 
        {

            await page.GetByRole(AriaRole.Link, new() { Name = "lightbulb_circle Insights" }).ClickAsync();
        }
    }
}
