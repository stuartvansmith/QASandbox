using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Playwright;


namespace QA.AutomationTests.RegressionTests
{
    [TestClass]
    [DoNotParallelize]
    public class InsightsModule : TestBase
    {
        [TestMethod]
        public async Task InsightsStaging()
        {
            await page.GotoAsync("https://staging.originbenefits.ai/login");
            await TestHelper.FinishLogin(page, "2026");
            await page.GetByRole(AriaRole.Link, new() { Name = "lightbulb_circle Insights" }).ClickAsync();
            
            await page.GetByText("Run Report").First.ClickAsync();

            // Ensure context was created with AcceptDownloads = true
            var filePath = await TestHelper.DownloadAndVerifyAsync(
                page,
                async () => await page.GetByText("Export Report").First.ClickAsync(),
                downloadsDir: "test-downloads",
                timeoutMs: 30000
            );

            // Now you can parse it or upload it as an artifact
            Console.WriteLine($"Downloaded file: {filePath}");


        }
        [TestMethod]
        public async Task deliberateFailToTestReport() 
        {
            Assert.IsTrue(false, "Deliberatly fail...");
        }
    }
}
