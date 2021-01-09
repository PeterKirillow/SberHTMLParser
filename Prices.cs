using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

namespace SberHTMLParser
{
    class Prices
    {
        public Table I;
        public Table C;

        public String filePath;

        Variables va;
        DataTable instruments;

        DateTime Date;
        String ISIN;
        double Nominal;
        String NominalCurrency;
        double Price;
        String PriceCurrency;
        double Accrued;
        String Instrument;
        double Course2USD;
        double Course2RUB;

        /******************************************************************/
        public Prices(Variables va)
        {
            this.va = va;
            I = new Table("Prices");
            C = new Table("Rates");
            instruments = (from t in va.tables_list where t.Name == "instruments" select t.table).ToArray().FirstOrDefault();
            va.tables_list.Add(this.I);
            va.tables_list.Add(this.C);
        }

        /******************************************************************/
        private void add_row_I()
        {
            I.table.Rows.Add(new string[] {
                Date.ToString(),
                ISIN,
                Nominal.ToString(),
                NominalCurrency,
                Price.ToString(),
                PriceCurrency,
                Accrued.ToString()
            });
        }

        /******************************************************************/
        private void add_row_C()
        {
            C.table.Rows.Add(new string[] {
                Date.ToString(),
                Instrument,
                Course2USD.ToString(),
                Course2RUB.ToString(),
            });
        }

        /******************************************************************/
        public bool calculate()
        {
            IXLWorksheet ws_p_instr;
            IXLWorksheet ws_p_curr;
            bool ret = false;
            try
            {
                using (XLWorkbook workBook = new XLWorkbook(filePath))
                {
                    /*----------------------------------------------------------------------------*/
                    ws_p_instr = workBook.Worksheet("prices_instr");
                    if (ws_p_instr != null)
                    {
                        var firstRowUsed = ws_p_instr.FirstRowUsed();
                        var firstPossibleAddress = ws_p_instr.Row(firstRowUsed.RowNumber()).FirstCell().Address;
                        var lastPossibleAddress = ws_p_instr.LastCellUsed().Address;
                        var range = ws_p_instr.Range(firstPossibleAddress, lastPossibleAddress).AsRange();
                        var table = range.AsTable();
                        var dataList = new List<string[]>
                    {
                     table.DataRange.Rows().Select(tableRow => tableRow.Field("Date").GetString()).ToArray(),
                     table.DataRange.Rows().Select(tableRow => tableRow.Field("ISIN").GetString()).ToArray(),
                     table.DataRange.Rows().Select(tableRow => tableRow.Field("Nominal").GetString()).ToArray(),
                     table.DataRange.Rows().Select(tableRow => tableRow.Field("NominalCurrency").GetString()).ToArray(),
                     table.DataRange.Rows().Select(tableRow => tableRow.Field("Price").GetString()).ToArray(),
                     table.DataRange.Rows().Select(tableRow => tableRow.Field("PriceCurrency").GetString()).ToArray(),
                     table.DataRange.Rows().Select(tableRow => tableRow.Field("AccInt").GetString()).ToArray(),
                    };
                        var rows = dataList.Select(array => array.Length).Concat(new[] { 0 }).Max();
                        for (var j = 0; j < rows; j++)
                        {                            
                            Date = Convert.ToDateTime(dataList[0][j]);
                            ISIN = dataList[1][j];
                            Nominal = Convert.ToDouble(dataList[2][j].Replace("NULL", "0"));
                            NominalCurrency = dataList[3][j].Replace("RUR", "RUB"); ;
                            Price = Convert.ToDouble(dataList[4][j].Replace("NULL", "0"));
                            PriceCurrency = dataList[5][j].Replace("RUR", "RUB");
                            Accrued = Convert.ToDouble(dataList[6][j].Replace("NULL", "0"));
                            add_row_I();
                        }
                        ret = true;
                    }
                    /*----------------------------------------------------------------------------*/
                    ws_p_curr = workBook.Worksheet("prices_curr");
                    if (ws_p_curr != null)
                    {
                        var firstRowUsed = ws_p_curr.FirstRowUsed();
                        var firstPossibleAddress = ws_p_curr.Row(firstRowUsed.RowNumber()).FirstCell().Address;
                        var lastPossibleAddress = ws_p_curr.LastCellUsed().Address;
                        var range = ws_p_curr.Range(firstPossibleAddress, lastPossibleAddress).AsRange();
                        var table = range.AsTable();
                        var dataList = new List<string[]>
                    {
                     table.DataRange.Rows().Select(tableRow => tableRow.Field("Date").GetString()).ToArray(),
                     table.DataRange.Rows().Select(tableRow => tableRow.Field("Instrument").GetString()).ToArray(),
                     table.DataRange.Rows().Select(tableRow => tableRow.Field("Course2USD").GetString()).ToArray(),
                     table.DataRange.Rows().Select(tableRow => tableRow.Field("Course2RUR").GetString()).ToArray(),
                    };

                        var rows = dataList.Select(array => array.Length).Concat(new[] { 0 }).Max();
                        for (var j = 0; j < rows; j++)
                        {
                            Date = Convert.ToDateTime(dataList[0][j]);
                            Instrument = dataList[1][j].Replace("RUR", "RUB");
                            Course2USD = Convert.ToDouble(dataList[2][j].Replace("NULL", "0"));
                            Course2RUB = Convert.ToDouble(dataList[3][j].Replace("NULL", "0"));
                            add_row_C();
                        }
                        ret = true;
                    }
                    /*----------------------------------------------------------------------------*/
                }
            } catch (Exception e)
            {
                Console.WriteLine($"Ошибка при парсинге файла с ценами {filePath}");
                Console.WriteLine(e);
                ret = false;
            }
            return ret;
        }

        /******************************************************************/
        public String[] get_price(String isin, DateTime dt)
        {
            string[] s_prices = new string[5];
            if (I.table != null)
            {
                var x = from r in I.table.AsEnumerable()
                        where r.Field<String>("ISIN") == isin && r.Field<DateTime>("Date") == dt && r.Field<Double>("Price") != 0
                        select new
                        {
                            Price = r.Field<double>("Price")
                            , PriceCurrency = r.Field<String>("PriceCurrency")
                            , Nominal = r.Field<double>("Nominal")
                            , NominalCurrency = r.Field<String>("NominalCurrency")
                            , Accrued = r.Field<double>("AccInt")
                        };
                if (x.Any())
                {
                    s_prices[0] = x.First().Price.ToString();
                    s_prices[1] = x.First().PriceCurrency;
                    s_prices[2] = x.First().Nominal.ToString();
                    s_prices[3] = x.First().NominalCurrency;
                    s_prices[4] = x.First().Accrued.ToString();
                }
            }
            return s_prices;
        }

    }
}
