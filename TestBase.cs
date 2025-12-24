using Microsoft.Playwright;

namespace QA.AutomationTests
{
    public class TestBase
    {
        protected IBrowser browser;
        protected IPage page;
        protected IPlaywright playwright;
        protected IBrowserContext context;
        protected string? authPath;

        [TestInitialize]
        public async Task Setup()
        {

            playwright = await Playwright.CreateAsync();
            var headless = Environment.GetEnvironmentVariable("CI") == "true";
            browser = await playwright.Chromium.LaunchAsync(new()
            {
                Headless = headless,
            });
            
            authPath = Path.Combine(Directory.GetParent(Environment.CurrentDirectory)!.Parent!.Parent!.FullName,
                                        "SSO",
                                        "authState.json");

            context = await browser.NewContextAsync(new BrowserNewContextOptions
            {
                Locale = "en-GB",
                TimezoneId = "Europe/London",
                Permissions = new[] { "geolocation" },
                StorageStatePath = authPath
                //RecordVideoDir = "test-videos",
                //RecordVideoSize = new() { Width = 1280, Height = 720 } 
            });
            page = await context.NewPageAsync();

            await Navigate();
        }

        public virtual async Task Navigate()
        {
            
        }

        [TestCleanup]
        public async Task Teardown()
        {
            if (context != null)
                await context.CloseAsync();  

            if (browser != null)
                await browser.CloseAsync();

            playwright?.Dispose();
        }

    }

    
}
