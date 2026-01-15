using Microsoft.Playwright;

namespace QA.AutomationTests;

[TestClass]
public class ResetSSO : TestBase
{
    [TestMethod]
    public async Task StoreSSO()
    {

        await browserContext.StorageStateAsync(new()
        {
            Path = authPath  
        });

        await browser.CloseAsync();

    }
}
