using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CRM_Raviz.Models
{
    public class AgentCasesViewModel
    {
        public string ModifiedAgent { get; set; }
        public int AccountNo { get; set; }
        public string Agent { get; set; }
        public string Segment { get; set; }
        public int Count { get; set; }
        public int CasesCountToday { get; set; }
        public int CasesCountThisMonth { get; set; }
        public int CallBackCountToday { get; set; }
        public int CallBackCountPrev { get; set; }
    }
}