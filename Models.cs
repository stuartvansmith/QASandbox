using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QA.AutomationTests
{
    public class BenefitToBeCreated()
    {
        public string BenefitName { get; set; } = String.Empty;
        public int Period { get; set; }
        public BenefitTerm BenefitTerm { get; set; }
    }
}

public enum BenefitTerm
{
    None = 0,
    Annual = 1,
    Indefinite = 2,
    MultiYear = 3
}