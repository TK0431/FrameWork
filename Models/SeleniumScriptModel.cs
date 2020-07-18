using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FrameWork.Models
{
    public class SeleniumScriptModel
    {
        public List<SeleniumOrder> Orders { get; set; } = new List<SeleniumOrder>();

        public Dictionary<string, Dictionary<string, List<SeleniumEvent>>> Events { get; set; } = new Dictionary<string, Dictionary<string, List<SeleniumEvent>>>();
    }

    public class SeleniumOrder
    {
        public string Case { get; set; }

        public string View { get; set; }

        public string ViewName { get; set; }

        public string Event { get; set; }
    }

    public class SeleniumEvent
    {
        public string No { get; set; }
        public string Event { get; set; }
    }
}
