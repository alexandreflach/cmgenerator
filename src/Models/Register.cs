using System;
using System.Collections.Generic;
using System.Text;

namespace cmgenerator.Models
{
    public class Register
    {
        public string Number { get; set; }

        public Area Area { get; set; }

        public string Action { get; set; }

        public DateTime PrevisionDate { get; set; }

        public DateTime ConclusionDate { get; set; }

        public DateTime ExtensionOne { get; set; }

        public DateTime ExtensionTwo { get; set; }

        public DateTime ExtensionThree { get; set; }

        public string Source { get; set; }

        public override string ToString()
        {
            return Number;
        }
    }
}
