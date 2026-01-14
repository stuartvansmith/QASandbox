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

            //var download = await page.RunAndWaitForDownloadAsync(async () =>
            //{
            //    var exportBtn = page.GetByText("Export Report").First;
            //    await exportBtn.WaitForAsync();
            //    await exportBtn.ClickAsync();
            //});
   
            //await page.PauseAsync();
            return;

            await page.Locator("#RYc3xSz0vU > .rz-dropdown-label").ClickAsync();
            await page.GetByRole(AriaRole.Listitem, new() { Name = "Afghanistan (AFG)" }).Locator("div").Nth(1).ClickAsync();
            await page.GetByRole(AriaRole.Listbox).GetByText("Afghanistan (AFG)").ClickAsync();
            await page.GetByRole(AriaRole.Listbox).GetByText("Afghanistan (AFG)").ClickAsync();
            await page.Locator("#i0DYrveBvU > .rz-dropdown-label").ClickAsync();
            await page.GetByRole(AriaRole.Listitem, new() { Name = "Afghanistan (AFG)" }).ClickAsync();
            await page.Locator("#QCEQ2j7V4k > .rz-dropdown-label").ClickAsync();
            await page.GetByRole(AriaRole.Listitem, new() { Name = "Active" }).Locator("div").Nth(1).ClickAsync();
            await page.Locator("#oCo54gB2QU > .rz-dropdown-label").ClickAsync();
            await page.GetByRole(AriaRole.Listitem, new() { Name = "Provider One" }).Locator("div").Nth(1).ClickAsync();
            await page.Locator("#gKdtCk0iIE > .rz-dropdown-trigger > .notranslate").ClickAsync();
            await page.GetByRole(AriaRole.Listitem, new() { Name = "Risk" }).Locator("div").Nth(1).ClickAsync();
            await page.Locator("#b5bYY8IAd0 > .rz-dropdown-trigger > .notranslate").ClickAsync();
            await page.GetByRole(AriaRole.Listitem, new() { Name = "Critical illness" }).Locator("div").Nth(1).ClickAsync();
            var download1 = await page.RunAndWaitForDownloadAsync(async () =>
            {
                await page.GetByRole(AriaRole.Button, new() { Name = "export_notes Export Report" }).ClickAsync();
            });
            await page.PauseAsync();   
        }
    }
}
