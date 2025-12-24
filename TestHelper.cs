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

namespace QA.AutomationTests
{
    public static class TestHelper
    {
        public static async Task FinishLogin(IPage page)
        {
            await page.GetByRole(AriaRole.Button, new() { Name = "button Microsoft" }).ClickAsync();
            

            await page.Locator(".notranslate").First.ClickAsync();
            await page.GetByRole(AriaRole.Option, new() { Name = "Smoke Test" }).ClickAsync();
            await TestHelper.SelectDropdownOptionByAriaAsync(
                page,
                "Origin.Common.Scheme.BenefitTermPeriod",
                "2025"
            );
            await page.GetByRole(AriaRole.Button, new() { Name = "navigate_next Next" }).ClickAsync();
        }

        public static async Task SelectDropdownOptionByAriaAsync(
           IPage page,
           string inputAriaLabel,
           string optionText,
           int timeoutMs = 10000)
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

            await page.GetByRole(AriaRole.Link, new() { Name = "image Ask Cuido" }).ClickAsync();
            
            var editBtn = page.GetByRole(AriaRole.Button, new() { Name = "edit_square" });
            //await editBtn.WaitForAsync(new() { State = WaitForSelectorState.Visible });
            await editBtn.ClickAsync();
            
            
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

        internal static async Task SwitchCountry(IPage page, string countryName, int timeoutMs = 5000)
        {
            // 1. Open the first country dropdown
            var regionPicker = page.Locator("div.rz-dropdown.rz-clear").First;
            await regionPicker.ClickAsync();

            // 2. Select country inside the listbox
            var listbox = page.GetByRole(AriaRole.Listbox).First;
            await listbox.WaitForAsync(new()
            {
                State = WaitForSelectorState.Visible,
                Timeout = timeoutMs
            });

            var option = listbox.GetByText(countryName, new() { Exact = true }).First;

            await option.ClickAsync(new()
            {
                Timeout = timeoutMs
            });

            // 3. Wait for the listbox to disappear (dropdown closed)
            await listbox.WaitForAsync(new()
            {
                State = WaitForSelectorState.Hidden,
                Timeout = timeoutMs
            });

            //// 4. Wait for the selected country label in the dropdown to show the new country
            //await regionPicker.GetByText(countryName, new() { Exact = true }).WaitForAsync(new()
            //{
            //    State = WaitForSelectorState.Visible,
            //    Timeout = timeoutMs
            //});

            // 5. OPTIONAL: wait for any loading overlay to go away
            // e.g. if you have a spinner like ".loading-overlay"
            // await page.Locator(".loading-overlay").WaitForAsync(new() { State = WaitForSelectorState.Hidden, Timeout = timeoutMs });
        }


        //internal static async Task SwitchCountry(IPage page, string countryName, int timeoutMs = 2000)
        //{

        //    var regionPicker = page.Locator("div.rz-dropdown.rz-clear").First;
        //    await regionPicker.ClickAsync();

        //    var listbox = page.GetByRole(AriaRole.Listbox).First;
        //    await listbox.WaitForAsync(new()
        //    {
        //        State = WaitForSelectorState.Visible,
        //        Timeout = timeoutMs
        //    });


        //    var option = listbox
        //        .GetByText(countryName, new() { Exact = true })
        //        .First;  

        //    await option.ClickAsync(new()
        //    {
        //        Timeout = timeoutMs
        //    });

        //}

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
