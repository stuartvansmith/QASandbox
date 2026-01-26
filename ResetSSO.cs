using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Playwright;

namespace QA.AutomationTests;

[TestClass]
public class ResetSSO 
{
    [TestMethod]
    public async Task StoreSSO()
    {
        var playwright = await Playwright.CreateAsync();
        var headless = Environment.GetEnvironmentVariable("CI") == "true";
        var browser = await playwright.Chromium.LaunchAsync(new()
        {
            Headless = headless,
        });
        
        var root = Directory.GetParent(Environment.CurrentDirectory)!.Parent!.Parent!.FullName;

        // Build path to SSO/authState.json
        var authPath = Path.Combine(root, "SSO", "authState.json");

        // Make sure folder exists
        Directory.CreateDirectory(Path.GetDirectoryName(authPath)!);

        //// Optional but VERY helpful in CI:
        //if (!File.Exists(authPath))
        //{
        //    throw new FileNotFoundException($"authState.json not found at {authPath}");
        //}
        var browserContext = await browser.NewContextAsync(new BrowserNewContextOptions
        {
            Locale = "en-GB",
            TimezoneId = "Europe/London",
            Permissions = new[] { "geolocation" },
            AcceptDownloads = true,
            // RecordVideoDir = "test-videos",
            // RecordVideoSize = new() { Width = 1280, Height = 720 } 
        });

        var page = await browserContext.NewPageAsync();
        await page.GotoAsync($"https://staging.originbenefits.ai/login");
        await page.PauseAsync();    
        await browserContext.StorageStateAsync(new()
        {
            Path = authPath  
        });

        await browser.CloseAsync();

    }
}
