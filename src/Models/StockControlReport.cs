using System;
using System.Collections.Generic;
using System.Text;

namespace CMGenerator.Models
{
    public class StockControlReport
    {
        public string Area { get; set; }

        public int ActionOutOfTime { get; set; }

        public int ActionOnTime { get; set; }

        public int ActionClosed { get; set; }

        public int ActionCanceled { get; set; }

        public int Total { get; set; }
    }
}
