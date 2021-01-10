using System;
using System.Collections.Generic;

namespace SberHTMLParser
{
    class Variables
    {
        // list of class Table
        public List<Table> tables_list = new List<Table>();
        // list of report dates
        public List<DateTime> report_dates = new List<DateTime>();
        // original html
        public string html = "";

        /******************************************************************/
        public Variables() { }
    }
}
