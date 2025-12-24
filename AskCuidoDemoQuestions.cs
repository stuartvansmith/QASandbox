using ClosedXML.Excel;
using Microsoft.Playwright;

namespace QA.AutomationTests;

public class QAPair
{
    public int Number { get; set; }
    public string Country { get; set; } = "";
    public string FilteredOn { get; set; } = "";
    public string Question { get; set; } = "";
    public string Answer { get; set; } = "";
}

[TestClass]
[DoNotParallelize]
public class AskCuidoDemoQuestions : TestBase
{
   
    [TestMethod]
    public async Task SalesDemoQuestions()
    {

        var results = new List<QAPair>();
        int counter = 1;
        string lastCountry = String.Empty;

        async Task AskAndRecord(string question, string filterOnBen = null, string country = null)
        {

            if (country != null)
            {
                lastCountry = country;
            }
            
            await TestHelper.AskCuidoQuestion(page, question, filterOnBen);
            var answer = await TestHelper.WaitForCuidoAnswer(page);

            results.Add(new QAPair
            {
                Number = counter++,
                Country = lastCountry,
                FilteredOn = filterOnBen,
                Question = question,
                Answer = answer
            });
        }

        await TestHelper.SwitchCountry(page, "Germany (DEU)");
        await AskAndRecord("Please outline which leaves are available, including which are statutory and which are non statutory.", country: "Germany");
        await AskAndRecord("Which plans have suicide exclusions?");

        await TestHelper.SwitchCountry(page, "Brazil (BRA)");
        await AskAndRecord("Which plans are self insured?", country: "Brazil");
        await AskAndRecord("What is our per employee per year cost for our Gym membership?", "Gym Membership");
        await TestHelper.ClearBenefitFilter(page);

        await AskAndRecord("Please tell me the fees associated with the Restaurant Ticket.", "Meal & Food Vouchers");
        await TestHelper.ClearBenefitFilter(page);
        await AskAndRecord("Please list the included coverage for our medical plan.");

        await TestHelper.SwitchCountry(page, "United States of America (USA)");
        await AskAndRecord("What are the contribution options for the 401k plan?", country: "USA");

        await TestHelper.SwitchCountry(page, "South Korea (KOR)");
        await AskAndRecord("What are the claim limits for AD&D insurance?", country: "South Korea");

        await TestHelper.SwitchCountry(page, "United Kingdom (GBR)");
        await AskAndRecord("Please create a list of all family friendly benefits that we offer", country: "United Kingdom");
        await AskAndRecord("What benefits have any clauses that have and identify any language, requirements, or assumptions that may conflict with principles of diversity, equity, and inclusion. Specifically, highlight: Terms or phrases that may be exclusionary, biased, or discriminatory. Provisions that disproportionately impact underrepresented or marginalized groups. Implicit norms or assumptions that reinforce homogeneity (e.g., around gender, culture, ability, or socioeconomic status). Lack of inclusive language or accommodations. Omissions where D&I considerations should be explicitly addressed. Provide suggestions for how the policy could be revised to better support inclusion and equity.");
        await AskAndRecord("Does our annual leave benefit have any clauses that have and identify any language, requirements, or assumptions that may conflict with principles of diversity, equity, and inclusion. Specifically, highlight: Terms or phrases that may be exclusionary, biased, or discriminatory. Provisions that disproportionately impact underrepresented or marginalized groups. Implicit norms or assumptions that reinforce homogeneity (e.g., around gender, culture, religion, ability, or socioeconomic status). Lack of inclusive language or accommodations. Omissions where D&I considerations should be explicitly addressed. Provide suggestions for how the policy could be revised to better support inclusion and equity.");

        await AskAndRecord("An employee has submitted their travel expenses 6 months after the travel took place. Are they still eligible for reimbursement?", "Commuting Allowance");
        await TestHelper.ClearBenefitFilter(page);
        await AskAndRecord("What exclusions are there in our global business travel policy? Please answer in bullet point format");
        await AskAndRecord("Are there any countries excluded from our global business travel program? If so, for what reason?");
        await AskAndRecord("Does our travel policy cover winter sports?");
        await AskAndRecord("An employee wants to add a 3 week vacation onto the end of their business trip - would they still be covered? If not, what do you suggest they do?");

        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Results");

        ws.Cell(1, 1).Value = "Number";
        ws.Cell(1, 2).Value = "Country";
        ws.Cell(1, 3).Value = "Benefit Filter";
        ws.Cell(1, 4).Value = "Question";
        ws.Cell(1, 5).Value = "Answer";

        int row = 2;
        foreach (var item in results)
        {
            ws.Cell(row, 1).Value = item.Number;
            ws.Cell(row, 2).Value = item.Country;
            ws.Cell(row, 3).Value = item.FilteredOn;
            ws.Cell(row, 4).Value = item.Question;
            ws.Cell(row, 5).Value = item.Answer.Length>10000 ? item.Answer.Substring(0,10000) + "...(truncated for report)": item.Answer;
            row++;
        }
        
        var filePath = Path.Combine(Directory.GetParent(Environment.CurrentDirectory)!.Parent!.Parent!.FullName,
                            "AskCuidoResults",
                            $"SalesDemoQuestions_{DateTime.UtcNow:yyyyMMdd_HHmmss}.xlsx");

        wb.SaveAs(filePath);

    }


    public override async Task Navigate()
    {
        await page.GotoAsync("https://demo.originbenefits.ai/login");
        await page.GetByRole(AriaRole.Button, new() { Name = "button Microsoft" }).ClickAsync();
        await page.Locator(".notranslate").First.ClickAsync();
        await page.GetByRole(AriaRole.Option, new() { Name = "Global Corp" }).ClickAsync();
        await TestHelper.SelectDropdownOptionByAriaAsync(
            page,
            "Origin.Common.Scheme.BenefitTermPeriod",
            "2025"
        );
        await page.GetByRole(AriaRole.Button, new() { Name = "navigate_next Next" }).ClickAsync();
        
    }
}

