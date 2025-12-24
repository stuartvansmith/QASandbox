using Microsoft.Playwright;

namespace QA.AutomationTests;

[TestClass]
public class ResetSSO : TestBase
{
    [TestMethod]
    public async Task StoreSSO()
    {

        await context.StorageStateAsync(new()
        {
            Path = authPath  
        });

        await browser.CloseAsync();

    }
}
