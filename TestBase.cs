using Microsoft.Playwright;
using System.Runtime.InteropServices.Marshalling;

namespace QA.AutomationTests
{
    [TestClass]
    public class TestBase
    {
        protected static IBrowser browser;
        protected static IPage page;
        protected static IPlaywright playwright;
        protected static IBrowserContext browserContext;

        protected static string? authPath;
        protected static string BaseUrl => Environment.GetEnvironmentVariable("TEST_BASE_URL")
                                       ?? "https://staging.originbenefits.ai"; // Default fallback

        [TestInitialize]
        public virtual async Task TestSetup()
        {

        }

        [ClassInitialize(InheritanceBehavior.BeforeEachDerivedClass)]
        public static async Task Setup(TestContext context)
        {

            playwright = await Playwright.CreateAsync();
            var headless = Environment.GetEnvironmentVariable("CI") == "true";
            browser = await playwright.Chromium.LaunchAsync(new()
            {
                Headless = headless,
            });

            
            var root = Environment.GetEnvironmentVariable("GITHUB_WORKSPACE")
                       ?? Directory.GetParent(Environment.CurrentDirectory)!.Parent!.Parent!.FullName;

            // Build path to SSO/authState.json
            var authPath = Path.Combine(root, "SSO", "authState.json");

            // Make sure folder exists
            Directory.CreateDirectory(Path.GetDirectoryName(authPath)!);

            // Optional but VERY helpful in CI:
            if (!File.Exists(authPath))
            {
                throw new FileNotFoundException($"authState.json not found at {authPath}");
            }

            browserContext = await browser.NewContextAsync(new BrowserNewContextOptions
            {
                Locale = "en-GB",
                TimezoneId = "Europe/London",
                Permissions = new[] { "geolocation" },
                StorageStatePath = authPath,
                AcceptDownloads = true,
               // RecordVideoDir = "test-videos",
               // RecordVideoSize = new() { Width = 1280, Height = 720 } 
            });

            page = await browserContext.NewPageAsync();
            await page.GotoAsync($"{BaseUrl}/login");
            await TestHelper.FinishLogin(page, "2026");
        }

        [TestCleanup]
        public virtual async Task TestCleanUp()
        { 
        }

        [ClassCleanup(InheritanceBehavior.BeforeEachDerivedClass)]
        public static async Task Teardown()
        {
            if (browserContext != null)
                await browserContext.CloseAsync();  

            if (browser != null)
                await browser.CloseAsync();

            playwright?.Dispose();
        }

    }

    
}
