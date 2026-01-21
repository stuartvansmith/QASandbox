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
                                       ?? "https://staging.originbenefits.ai"; // Default fallback
        protected static List<string> slowRequests;
        protected static ConcurrentDictionary<string, DateTime> requestStartTimes; // Add this
        protected static Dictionary<string, DateTime> namedTimers = new Dictionary<string, DateTime>();
        protected static Dictionary<string, double> timedOperations = new Dictionary<string, double>();



        [TestInitialize]
        public virtual async Task TestSetup()
        {
            slowRequests?.Clear();
            namedTimers?.Clear();
            timedOperations?.Clear();
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



            Console.WriteLine("🔧 Setting up network monitoring...");

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
                    Console.WriteLine($"📡 {duration:F0}ms - {url}");

                    if (duration > 1500) 
                    {
                        var message = $"{duration:F0}ms - {url}";
                        slowRequests.Add(message);
                        Console.WriteLine($"⚠️ Slow request: {message}");
                    }
                    DateTime removed;
                    requestStartTimes.TryRemove(url, out removed); // Clean up
                }
            };

            Console.WriteLine("🔧 Network monitoring ready");


            await page.GotoAsync($"{BaseUrl}/login");
            await TestHelper.FinishLogin(page, "2026");
        }

        [TestCleanup]
        public virtual async Task TestCleanUp()
        {
            // Report network stats after each test
            if (slowRequests != null && slowRequests.Any())
            {
                Console.WriteLine($"\n⚠️ Test completed with {slowRequests.Count} slow request(s):");
                foreach (var req in slowRequests.Take(10))
                {
                    Console.WriteLine($"  {req}");
                }
            }
            else
            {
                Console.WriteLine("✅ No slow requests detected in this test");
            }

            // Report timed operations
            if (timedOperations != null && timedOperations.Any())
            {
                Console.WriteLine($"\n⏱️ Timed Operations:");
                foreach (var op in timedOperations.OrderByDescending(x => x.Value))
                {
                    Console.WriteLine($"  {op.Key}: {op.Value:F0}ms");
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
            Console.WriteLine($"⏱️ Started: {operationName}");
        }

        public static void StopTimer(string operationName)
        {
            if (namedTimers.TryGetValue(operationName, out var startTime))
            {
                var duration = (DateTime.Now - startTime).TotalMilliseconds;
                timedOperations[operationName] = duration;
                Console.WriteLine($"✅ Completed: {operationName} took {duration:F0}ms");

                if (duration > 5000) // Flag operations over 5 seconds
                {
                    Console.WriteLine($"⚠️ WARNING: {operationName} took longer than expected!");
                }

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

    }

    
}
