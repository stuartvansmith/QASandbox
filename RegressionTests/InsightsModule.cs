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
            Console.WriteLine("Description: This test opens the Benefit summary report and exports data with no filters.");
            Console.WriteLine("Asserts: The excel sheet exists and is not 0 bytes");
            Console.WriteLine("Assumes: The SmokeTest succesfully ran.");
            TestBase.StartTimer("Benefit Summary Report");
            await page.GetByText("Run Report").Nth(0).ClickAsync();

            var filePath = await TestHelper.DownloadAndVerifyAsync(
                page,
                async () => await page.GetByText("Export Report").First.ClickAsync(),
                downloadsDir: "test-downloads",
                timeoutMs: 30000
            );
            TestBase.StopTimer("Benefit Details Report");
        }

        [TestMethod]
        public async Task BenefitDetailReport()
        {
            Console.WriteLine("Description: This test opens the Benefit detail report family=Risk Type=Critical Illness, exports data.");
            Console.WriteLine("Asserts: The excel sheet exists and is not 0 bytes");
            Console.WriteLine("Assumes: The SmokeTest succesfully ran.");

            TestBase.StartTimer("Benefit Details Report");
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
            TestBase.StopTimer("Benefit Details Report");
        }


        [TestInitialize]
        public override async Task TestSetup() 
        {
            await page.GetByRole(AriaRole.Link, new() { Name = "lightbulb_circle Insights" }).ClickAsync();
        }
    }
}
