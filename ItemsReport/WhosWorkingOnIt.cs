using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ItemsReport
{
    class WhosWorkingOnIt
    {
        public string name { get; set; }
        public DateTime workingDate { get; set; }
        public byte status { get; set; }
    }

    public enum WorkingStatus
    {
        NotWorkingOnIt = 0,
        WorkingOnIt = 1,
        DoneWorkingOnIt = 2
    }
}
