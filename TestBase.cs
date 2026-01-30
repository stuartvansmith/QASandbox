using Microsoft.Playwright;
using System.Collections.Concurrent;
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
                                       ?? "https://staging.originbenefits.ai/"; // Default fallback
        protected static List<string> slowRequests;
        protected static ConcurrentDictionary<string, DateTime> requestStartTimes; // Add this
        protected static Dictionary<string, DateTime> namedTimers = new Dictionary<string, DateTime>();
        protected static Dictionary<string, double> timedOperations = new Dictionary<string, double>();
        protected static List<string> allNetworkRequests = new List<string>();
        protected static string RootDirectory => Environment.GetEnvironmentVariable("GITHUB_WORKSPACE")
                       ?? Directory.GetParent(Environment.CurrentDirectory)!.Parent!.Parent!.FullName;

        [TestInitialize]
        public virtual async Task TestSetup()
        {
            slowRequests?.Clear();
            namedTimers?.Clear();
            timedOperations?.Clear();
            allNetworkRequests?.Clear();
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

            
           

            // Build path to SSO/authState.json
            var authPath = Path.Combine(RootDirectory, "SSO", "authState.json");

            // Make sure folder exists
            Directory.CreateDirectory(Path.GetDirectoryName(authPath)!);

            // Optional but VERY helpful in CI:
            if (!File.Exists(authPath))
            {
                throw new FileNotFoundException($"authState.json not found at {authPath}");
            }
            // Video directory at root
            var videoDirectory = Path.Combine(RootDirectory, "test-videos");
            Directory.CreateDirectory(videoDirectory);

            browserContext = await browser.NewContextAsync(new BrowserNewContextOptions
            {
                Locale = "en-GB",
                TimezoneId = "Europe/London",
                Permissions = new[] { "geolocation" },
                StorageStatePath = authPath,
                AcceptDownloads = true,
                //RecordVideoDir = videoDirectory,
                //RecordVideoSize = new() { Width = 1280, Height = 720 } 
            });

            page = await browserContext.NewPageAsync();

            slowRequests = new List<string>();
            requestStartTimes = new ConcurrentDictionary<string, DateTime>();

            // Track when requests START
            page.Request += (_, request) =>
            {
                requestStartTimes[request.Url] = DateTime.Now;
            };

            // Track when responses arrive
            page.Response += async (_, response) =>
            {
                var url = response.Url;

                if (requestStartTimes.TryGetValue(url, out var startTime))
                {
                    var duration = (DateTime.Now - startTime).TotalMilliseconds;
                    allNetworkRequests.Add($"📡 {duration:F0}ms - {url}");

                    if (duration > 1500) 
                    {
                        var message = $"{duration:F0}ms - {url}";
                        slowRequests.Add(message);
                    }
                    DateTime removed;
                    requestStartTimes.TryRemove(url, out removed); // Clean up
                }
            };

            await page.GotoAsync($"{BaseUrl}/login");
            await TestHelper.FinishLogin(page, "2026");
        }

        [TestCleanup]
        public virtual async Task TestCleanUp()
        {
            if (timedOperations != null && timedOperations.Any())
            {
                Console.WriteLine($"\n⏱ Timed Operations:");
                foreach (var op in timedOperations.OrderByDescending(x => x.Value))
                {
                    Console.WriteLine($"  {op.Key}: {op.Value:F0}ms");
                }
            }

            if (slowRequests != null && slowRequests.Any())
            {
                Console.WriteLine($"\n⚠️ Test completed with {slowRequests.Count} slow request(s):");
                foreach (var req in slowRequests)
                {
                    Console.WriteLine($"  {req}");
                }
            }
            else
            {
                Console.WriteLine("✅ No slow requests detected in this test");
            }

            if (allNetworkRequests != null && allNetworkRequests.Any())
            {
                try
                {
                    await LogNetworkRequestsToFile();
                    Console.WriteLine("✅ Network logs written successfully");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"⚠️ Failed to write network logs: {ex.Message}");
                }
            }
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
        public static void StartTimer(string operationName)
        {
            namedTimers[operationName] = DateTime.Now;
           // Console.WriteLine($"⏱️ Started: {operationName}");
        }

        public static void StopTimer(string operationName)
        {
            if (namedTimers.TryGetValue(operationName, out var startTime))
            {
                var duration = (DateTime.Now - startTime).TotalMilliseconds;
                timedOperations[operationName] = duration;
                //Console.WriteLine($"✅ Completed: {operationName} took {duration:F0}ms");

                //if (duration > 64738) // Flag operations over 5 seconds
                //{
                //    Console.WriteLine($"⚠️ WARNING: {operationName} took longer than expected!");
                //}

                namedTimers.Remove(operationName);
            }
            else
            {
                Console.WriteLine($"❌ Error: Timer '{operationName}' was not started");
            }
        }

        protected static double GetTimerDuration(string operationName)
        {
            return timedOperations.ContainsKey(operationName)
                ? timedOperations[operationName]
                : -1;
        }

        private async Task LogNetworkRequestsToFile()
        {
            // In CI: write to repo root (for artifacts)
            // Locally: write to temp folder (to avoid VS freeze)
            var logDirectory = Environment.GetEnvironmentVariable("GITHUB_WORKSPACE") != null
                ? Path.Combine(Environment.GetEnvironmentVariable("GITHUB_WORKSPACE")!, "network-logs")
                : Path.Combine(Path.GetTempPath(), "OriginNetworkLogs");

            Directory.CreateDirectory(logDirectory);

            var timestamp = DateTime.UtcNow;
            var fileName = $"network-requests-{timestamp:yyyy-MM-dd HHmmss}.csv";
            var filePath = Path.Combine(logDirectory, fileName);

            bool fileExists = File.Exists(filePath);

            var testEnv = GetEnvironmentFromBaseUrl(BaseUrl);

            await using (var writer = new StreamWriter(filePath, append: true))
            {
                if (!fileExists)
                {
                    await writer.WriteLineAsync("Date,Timestamp,Env,Category,DurationMs,Url");
                }

                foreach (var req in allNetworkRequests)
                {
                    var parts = req.Replace("📡 ", "").Split(new[] { "ms - " }, StringSplitOptions.None);
                    if (parts.Length == 2)
                    {
                        var duration = parts[0].Trim();
                        var url = parts[1].Trim().Replace(",", ";");
                        var category = GetCategory(url);
                        var date = timestamp.ToString("yyyy-MM-dd");
                        var time = timestamp.ToString("yyyy-MM-dd HH:mm:ss");

                        await writer.WriteLineAsync($"{date},{time},{testEnv},{category},{duration},{url}");
                    }
                }

                await writer.FlushAsync();
            }

            Console.WriteLine($"✅ Logs written to: {logDirectory}");
        }

        private string GetCategory(string url)
        {
            
            if (url.Contains("api.origin", StringComparison.OrdinalIgnoreCase))
                return "Origin Api";


            if (url.Contains("/login", StringComparison.OrdinalIgnoreCase))
                return "Login";

            if (url.Contains(".wasm", StringComparison.OrdinalIgnoreCase) ||
                url.Contains("blazor", StringComparison.OrdinalIgnoreCase))
                return "Framework";

            if (url.Contains(".png", StringComparison.OrdinalIgnoreCase) ||
                url.Contains(".jpg", StringComparison.OrdinalIgnoreCase) ||
                url.Contains(".svg", StringComparison.OrdinalIgnoreCase))
                return "Images";

            if (url.Contains(".js", StringComparison.OrdinalIgnoreCase) ||
                url.Contains(".css", StringComparison.OrdinalIgnoreCase))
                return "Scripts";

            if (url.Contains("/api/", StringComparison.OrdinalIgnoreCase) ||
                url.Contains("/Hub/", StringComparison.OrdinalIgnoreCase))
                return "API";

            return "Other";
        }

        static string GetEnvironmentFromBaseUrl(string url)
        {
            if (url.Contains("staging.originbenefits.ai", StringComparison.OrdinalIgnoreCase))
                return "Staging";
            if (url.Contains("demo.originbenefits.ai", StringComparison.OrdinalIgnoreCase))
                return "Demo";
            if (url.Contains("web-origin-live.azurewebsites.net", StringComparison.OrdinalIgnoreCase))
                return "Production";

            return "Unknown";
        }
    }

    
}
