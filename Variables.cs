using System;
using System.Collections.Generic;
using System.Data;

namespace SberHTMLParser
{
    class Variables
    {
        public List<Table> tables_list = new List<Table>();
        public List<DateTime> report_dates = new List<DateTime>();
        public string html = "";

        public Variables() { }
    }
}
