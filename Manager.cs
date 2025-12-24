using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Playwright;
using System.Text.RegularExpressions;
using System.Threading;

namespace QA.AutomationTests
{
    public static class Manager
    {
        public static async Task CreatePolicy(IPage page, string policyName)
        {

            await page.GetByRole(AriaRole.Link, new() { Name = "assignment Policies" }).ClickAsync();
            await page.GetByText("check_circleActive policies").ClickAsync();

            try
            {
                await page.GetByRole(AriaRole.Button, new() { Name = "more_vert" }).ClickAsync(new() { Timeout = 2000 });
                await page.Locator("span").Filter(new() { HasTextRegex = new Regex("^Archive$") }).ClickAsync(new() { Timeout = 2000 });
            }
            catch (Exception ex) { }


            await page.GetByText("architectureDraft policies").ClickAsync();
            try
            {
                await page.GetByRole(AriaRole.Button, new() { Name = "more_vert" }).ClickAsync(new() { Timeout = 2000 });
                await page.Locator("span").Filter(new() { HasTextRegex = new Regex("^Archive$") }).ClickAsync(new() { Timeout = 2000 });
            }
            catch (Exception) { }


            await page.GetByText("archiveArchive").ClickAsync();
            try
            {
                await page.GetByRole(AriaRole.Button, new() { Name = "more_vert" }).ClickAsync(new() { Timeout = 2000 });
                await page.GetByText("Delete", new() { Exact = true }).ClickAsync(new() { Timeout = 2000 });
            }
            catch (Exception) { }




            await Task.Delay(3000);

            await page.GetByRole(AriaRole.Link, new() { Name = "assignment Policies" }).ClickAsync();
            await page.GetByRole(AriaRole.Button, new() { Name = "add_box Add a policy" }).ClickAsync();
            await page.Locator("input[name='Country']").EvaluateAsync("el => el.parentElement.click()");
            await page.GetByRole(AriaRole.Option, new() { Name = "Afghanistan (AFG)" }).ClickAsync();

            await page.Locator("[id=\"Policy Name\"]").ClickAsync();
            await page.Locator("[id=\"Policy Name\"]").FillAsync(policyName);

            await page.GetByRole(AriaRole.Radio, new() { Name = "Health and wellbeing", Exact = true }).ClickAsync();
            await page.GetByRole(AriaRole.Button, new() { Name = "> Next" }).ClickAsync();

            var filePath = Path.Combine(Directory.GetParent(Environment.CurrentDirectory)!.Parent!.Parent!.FullName,
                                        "Documents",
                                        "Diversity Policy Portugal.pdf");

            var fileInput = page.Locator("input[type='file']");


            await fileInput.WaitForAsync();


            await fileInput.SetInputFilesAsync(filePath);


            await page.GetByText("Validate document import", new() { Exact = true })
                      .WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = 300_000 });
            await Task.Delay(3000);

            await page.GetByRole(AriaRole.Button, new() { Name = "Save", Exact = true }).ClickAsync();
            await page.GetByRole(AriaRole.Button, new() { Name = "Save without adding a note" }).ClickAsync();
            await Task.Delay(3000);

            var publishBtn = page.GetByRole(AriaRole.Button, new() { Name = "task_alt Publish" });
            await publishBtn.ScrollIntoViewIfNeededAsync();
            await publishBtn.ClickAsync();

        }
        public static async Task CreateBenefit(IPage page, BenefitToBeCreated benefit)
        {
            await DeleteBenefit(page, benefit);


            await page.GetByRole(AriaRole.Link, new() { Name = "star_border Benefits" }).ClickAsync();
            await page.GetByRole(AriaRole.Button, new() { Name = "add_box Add Benefit" }).ClickAsync();
            try
            {
                var loc = page.Locator("input[name='Country domiciled']");
                await loc.WaitForAsync(new()
                {
                    State = WaitForSelectorState.Attached, 
                    Timeout = 1000
                });

                await loc.EvaluateAsync("el => el.parentElement.click()");
            }
            catch (TimeoutException)
            {
                var fallback = page.Locator("input[name='Country']");
                await fallback.WaitForAsync(new()
                {
                    State = WaitForSelectorState.Attached, 
                    Timeout = 1000
                });

                await fallback.EvaluateAsync("el => el.parentElement.click()");
            }


            await page.GetByRole(AriaRole.Option, new() { Name = "Afghanistan (AFG)" }).ClickAsync();
            await page.Locator("[id=\"Benefit Name\"]").ClickAsync();
            await page.Locator("[id=\"Benefit Name\"]").FillAsync(benefit.BenefitName);
            await page.GetByRole(AriaRole.Radio, new() { Name = "Risk", Exact = true }).ClickAsync();

            await page.Locator("input[name='Benefit Type']").EvaluateAsync("el => el.parentElement.click()");
            await page.GetByRole(AriaRole.Option, new() { Name = "Critical illness" }).ClickAsync();

            await page.GetByText("None").Nth(2).ClickAsync();
            await page.GetByRole(AriaRole.Option, new() { Name = "Annual" }).ClickAsync();


            await page.Locator("input[name='Benefit term period']").EvaluateAsync("el => el.parentElement.click()");
            await page.GetByRole(AriaRole.Option, new() { Name = "2025" }).ClickAsync();

            await page.GetByRole(AriaRole.Button, new() { Name = "> Next" }).ClickAsync();

            
            //var filePath = Path.Combine(Directory.GetParent(Environment.CurrentDirectory)!.Parent!.Parent!.FullName,
            //                            "Documents",
            //                            "Critical illness policy.pdf");

            //var fileInput = page.Locator("input[type='file']");


            //await fileInput.WaitForAsync();


            //await fileInput.SetInputFilesAsync(filePath);


            //await page.GetByText("Validate document import", new() { Exact = true })
            //          .WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = 300_000 });

            await page.GetByRole(AriaRole.Button, new() { Name = "Save", Exact = true }).ClickAsync();
            await page.GetByRole(AriaRole.Button, new() { Name = "Save without adding a note" }).ClickAsync();

            return;

            await page.Locator("input[name=\"End Date\"]").ClickAsync();
            await page.Locator("[id=\"Start Date\"]").DblClickAsync();
            await page.Locator("[id=\"Start Date\"]").FillAsync("01/01/2025");
            await page.Locator("[id=\"Start Date\"]").PressAsync("Tab");
            await page.Locator("input[name=\"End Date\"]").FillAsync("31/12/2025");
            await page.Locator("input[name=\"End Date\"]").PressAsync("Tab");
            await page.Locator("[id=\"Renewal Date\"]").FillAsync("01/01/2026");
            await page.Locator("[id=\"Renewal Date\"]").PressAsync("Tab");

            await page.Locator("input[name='Provider']").EvaluateAsync("el => el.parentElement.click()");
            await page.GetByRole(AriaRole.Option, new() { Name = "Provider One" }).ClickAsync();

            await page.GetByRole(AriaRole.Button, new() { Name = "Save", Exact = true }).ClickAsync();
            await page.GetByRole(AriaRole.Button, new() { Name = "Save without adding a note" }).ClickAsync();


            await page.GetByText("Financials", new() { Exact = true }).ClickAsync();
            await page.GetByText("Costs", new() { Exact = true }).ClickAsync();
            await page.GetByText("No costs").ClickAsync();
            await page.GetByText("Commissions", new() { Exact = true }).ClickAsync();
            await page.GetByText("No commission").ClickAsync();
            await page.GetByText("Fees", new() { Exact = true }).ClickAsync();
            await page.GetByText("No fees").ClickAsync();

            await page.GetByRole(AriaRole.Button, new() { Name = "task_alt Publish" }).ClickAsync();
        }
        internal static async Task DeleteBenefit(IPage page, BenefitToBeCreated benefit)
        {
            await page.GetByRole(AriaRole.Link, new() { Name = "star_border Benefits" }).ClickAsync();




            
            var year = page
                    .Locator("i.rzi", new() { HasText = "calendar_month" })
                    .Locator("xpath=ancestor::div[contains(@class,'rz-form-field')]")
                    .Locator(".rz-dropdown");

            await year.ClickAsync();
            await year.PressAsync("ArrowUp");
            await year.PressAsync("Enter");
            await DeleteYear(page, benefit);

                
            await year.ClickAsync();
            await year.PressAsync("ArrowDown");
            await year.PressAsync("Enter");
            await DeleteYear(page, benefit);


            



        }

        private static async Task DeleteYear(IPage page, BenefitToBeCreated benefit)
        {
            await page.GetByText("Published benefits", new() { Exact = true }).ClickAsync();
            try
            {
                await page.GetByRole(AriaRole.Button, new() { Name = "more_vert" }).ClickAsync(new() { Timeout = 2000 });
                await page.Locator("span").Filter(new() { HasTextRegex = new Regex("^Archive$") }).ClickAsync(new() { Timeout = 2000 });
            }
            catch (Exception) { }
            await page.GetByText("Draft benefits", new() { Exact = true }).ClickAsync();
            try
            {
                await page.GetByRole(AriaRole.Button, new() { Name = "more_vert" }).ClickAsync(new() { Timeout = 2000 });
                await page.Locator("span").Filter(new() { HasTextRegex = new Regex("^Archive$") }).ClickAsync(new() { Timeout = 2000 });
            }
            catch (Exception) { }
            await page.GetByText("Archive", new() { Exact = true }).ClickAsync();
            try
            {
                await page.GetByRole(AriaRole.Button, new() { Name = "more_vert" }).ClickAsync(new() { Timeout = 2000 });
                await page.GetByText("Delete", new() { Exact = true }).ClickAsync(new() { Timeout = 2000 });
            }
            catch (Exception) { }
        }

        internal static async Task RenewBenefit(IPage page, BenefitToBeCreated benefit)
        {
            
            await page.GetByRole(AriaRole.Link, new() { Name = "star_border Benefits" }).ClickAsync();
            await page.GetByRole(AriaRole.Link, new() { Name = benefit.BenefitName }).ClickAsync();
            await page.GetByRole(AriaRole.Button, new() { Name = "autorenew Renew" }).ClickAsync();
            await page.GetByText("Renew on same terms and provider").ClickAsync();
            await page.GetByRole(AriaRole.Button, new() { Name = "chevron_right Next" }).ClickAsync();
            await page.GetByRole(AriaRole.Button, new() { Name = "chevron_right Next" }).ClickAsync();
            await page.GetByRole(AriaRole.Button, new() { Name = "chevron_right Next" }).ClickAsync();
            await page.GetByRole(AriaRole.Button, new() { Name = "task_alt Complete renenal" }).ClickAsync();

            await page.GetByRole(AriaRole.Button, new() { Name = "task_alt Publish" }).ClickAsync();
        }
    }
}