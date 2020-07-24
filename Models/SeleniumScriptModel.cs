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

        public List<SeleniumCheckItem> Initems = new List<SeleniumCheckItem>();

        public List<SeleniumCheckItem> Outitems = new List<SeleniumCheckItem>();
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
        public string Key { get; set; }
        public string Event { get; set; }
        public string Back { get; set; }
    }

    public class SeleniumCheckItem
    {
        public int col1 { get; set; }
        public string name1 { get; set; }

        public int col2 { get; set; }
        public string name2 { get; set; }
    }
}
