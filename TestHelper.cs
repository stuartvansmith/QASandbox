using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Playwright;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using static Microsoft.Playwright.Assertions;

namespace QA.AutomationTests
{
    public static class TestHelper
    {

        public static async Task<string> DownloadAndVerifyAsync(
            IPage page,
            Func<Task> triggerDownload,
            string downloadsDir = "downloads",
            int timeoutMs = 5000)
        {
         
            var download = await page.RunAndWaitForDownloadAsync(triggerDownload, new()
            {
                Timeout = timeoutMs
            });

            var fullDir = Path.Combine(Directory.GetCurrentDirectory(), downloadsDir);
            Directory.CreateDirectory(fullDir);

            var filePath = Path.Combine(fullDir, download.SuggestedFilename);
            await download.SaveAsAsync(filePath);

         
            Assert.IsTrue(File.Exists(filePath), $"Download file does not exist: {filePath}");

            var info = new FileInfo(filePath);
            Assert.IsTrue(info.Length > 0, $"Downloaded file is empty: {filePath}");

            return filePath;
        }
        public static async Task FinishLogin(IPage page, string period, string tenant = "Smoke Test")
        {
            
            try
            {
                await page.GetByRole(AriaRole.Button, new() { Name = "button Microsoft" }).ClickAsync(new() { Timeout = 60000 });
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.WriteLine("Couldn't find the microsft button for sso");
            }

            try
            {
                await page.Locator(".notranslate").First.ClickAsync(new() { Timeout = 60000 });
                await page.GetByRole(AriaRole.Option, new() { Name = tenant }).ClickAsync(new() { Timeout = 2000 });
            }
            catch (Exception ex) 
            {
                Console.WriteLine(ex);  
                Console.WriteLine("Couldn't find the select tenant dropdown when logging in");
            } 
            
            await TestHelper.SelectDropdownOptionByAriaAsync(
                page,
                "Origin.Common.Scheme.BenefitTermPeriod",
                period
            );
            await page.GetByRole(AriaRole.Button, new() { Name = "navigate_next Next" }).ClickAsync(new() { Timeout = 20000 });
        }

        public static async Task SelectDropdownOptionByAriaAsync(
           IPage page,
           string inputAriaLabel,
           string optionText,
           int timeoutMs = 20000)
        {

            
            var dropdown = page.Locator($".rz-dropdown:has(.rz-helper-hidden-accessible input[aria-label='{inputAriaLabel}'])");
            await dropdown.First.WaitForAsync(new() { Timeout = timeoutMs });


            await dropdown.First.ClickAsync();


            var panel = page.Locator(".rz-dropdown-panel").Filter(new() { HasText = optionText }).First;
            await panel.WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = timeoutMs });


            try
            {
                await panel.GetByRole(AriaRole.Option, new() { Name = optionText, Exact = true }).ClickAsync();
            }
            catch
            {
                await panel.Locator($".rz-dropdown-item:has-text('{optionText}')").First.ClickAsync();
            }


            await panel.WaitForAsync(new() { State = WaitForSelectorState.Hidden, Timeout = timeoutMs });

        }

        internal static async Task AskCuidoQuestion(IPage page, string question, string filterOnBen)
        {

            await page.GetByRole(AriaRole.Link, new() { Name = "image Ask Cuido" }).ClickAsync(new() { Timeout = 30000 });

            var editBtn = page.GetByRole(AriaRole.Button, new() { Name = "edit_square" });

            await editBtn.ClickAsync(new() { Timeout = 30000 });


            bool ratingQuestion = await page.GetByText("Overall, how satisfied were you with Cuido’s last conversation?").IsVisibleAsync();
            

            if (ratingQuestion)
            {
                await page.GetByLabel("Rate").Nth(2).ClickAsync();
                await page.GetByRole(AriaRole.Button, new() { Name = "save Submit rating" }).ClickAsync();
            }

            if (filterOnBen != null)
            {
                await AddBenefitFilter(page, filterOnBen);
            }

            var input = page.GetByRole(AriaRole.Textbox, new() { Name = "How can I help?" });

            await input.WaitForAsync(new()
            {
                State = WaitForSelectorState.Visible,  
                Timeout = 10000
            });
            await input.ClickAsync();   
            await input.FillAsync(question);
            await page.GetByRole(AriaRole.Button, new() { Name = "send" }).ClickAsync();

        }
        internal static async Task AddBenefitFilter(IPage page, string filterOnBen)
        {
            await page.GetByRole(AriaRole.Button, new() { Name = "filter_alt" }).ClickAsync();
            var loc = page.Locator("input[name='Benefit']");
            await loc.EvaluateAsync("el => el.parentElement.click()");
            await page.GetByRole(AriaRole.Option, new() { Name = filterOnBen }).First.ClickAsync();
            await page.GetByRole(AriaRole.Button, new() { Name = "save Set conversation filters" }).ClickAsync();
        }
        internal static async Task ClearBenefitFilter(IPage page)
        {
         
            await page.GetByRole(AriaRole.Button, new() { Name = "filter_alt" }).ClickAsync();
            await page.GetByRole(AriaRole.Dialog, new() { Name = "Conversation filters" }).Locator("i").Nth(1).ClickAsync();
            await page.GetByRole(AriaRole.Button, new() { Name = "save Set conversation filters" }).ClickAsync();
        }

        internal static async Task SwitchCountry(IPage page, string countryName, int timeoutMs = 60000)
        {
            var regionPicker = page.Locator("div.rz-dropdown.rz-clear").First;

            // Log what we're seeing
            Console.WriteLine($"=== DROPDOWN DEBUG START ===");
            Console.WriteLine($"Page URL: {page.Url}");
            Console.WriteLine($"Dropdown count: {await page.Locator("div.rz-dropdown.rz-clear").CountAsync()}");
            Console.WriteLine($"Dropdown visible: {await regionPicker.IsVisibleAsync()}");
            Console.WriteLine($"Dropdown enabled: {await regionPicker.IsEnabledAsync()}");

            // Take screenshot before
            await page.ScreenshotAsync(new() { Path = "test-videos/before-dropdown-click.png", FullPage = true });

            await regionPicker.ClickAsync(new() { Timeout = timeoutMs });
            Console.WriteLine("Click completed");

            // Take screenshot after
            await page.ScreenshotAsync(new() { Path = "test-videos/after-dropdown-click.png", FullPage = true });

            // Check what happened
            var listboxCount = await page.GetByRole(AriaRole.Listbox).CountAsync();
            Console.WriteLine($"Listbox count after click: {listboxCount}");

            // Check if dropdown has aria-expanded
            var ariaExpanded = await regionPicker.GetAttributeAsync("aria-expanded");
            Console.WriteLine($"Dropdown aria-expanded: {ariaExpanded}");

            // Log all listboxes on the page
            var allListboxes = await page.Locator("[role='listbox']").CountAsync();
            Console.WriteLine($"All elements with role=listbox: {allListboxes}");

            Console.WriteLine($"=== DROPDOWN DEBUG END ===");

            var listbox = page.GetByRole(AriaRole.Listbox).First;
            await listbox.WaitForAsync(new()
            {
                State = WaitForSelectorState.Visible,
                Timeout = timeoutMs
            });
            //var regionPicker = page.Locator("div.rz-dropdown.rz-clear").First;
            //await regionPicker.ClickAsync();

            //var listbox = page.GetByRole(AriaRole.Listbox).First;
            //await listbox.WaitForAsync(new()
            //{
            //    State = WaitForSelectorState.Visible,
            //    Timeout = timeoutMs
            //});

            //var option = listbox.GetByText(countryName, new() { Exact = true }).First;

            //await option.ClickAsync(new()
            //{
            //    Timeout = timeoutMs
            //});

            //await listbox.WaitForAsync(new()
            //{
            //    State = WaitForSelectorState.Hidden,
            //    Timeout = timeoutMs
            //});

        }

        internal static async Task<string> WaitForCuidoAnswer(IPage page)
        {
            var lastAnswer = page.Locator(".question-answer").Last;
            bool thinking = true;
            var responseText = String.Empty;
            while (thinking)
            {
                responseText = await lastAnswer.InnerTextAsync();
                thinking = responseText == @"Thinking...";

                await Task.Delay(1000);
            }
            return responseText;
        }
    }
}
