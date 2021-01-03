using System;
using HtmlAgilityPack;
using System.Data;
using System.Linq;
using System.Collections.Generic;

namespace SberHTMLParser
{
    class Position
    {
        public Table T;

        Variables va;
        DataTable instruments;
        DataTable money;
        DataTable portfolio;
        DataTable operations;
        DataTable deals;
        DataTable operations_other;

        /******************************************************************/
        public Position(Variables va)
        {
            this.va = va;
            T = new Table("position");
        }

        /******************************************************************/
        private void add_row(string[] row)
        {
            T.table.Rows.Add(row);
        }

        /******************************************************************/
        private void delete_row(DateTime d)
        {
            var query = from r in T.table.AsEnumerable() where r.Field<DateTime>("Date") == d && Math.Round(r.Field<Double>("Quantity"),4) == 0 select r;
            foreach (var row in query.ToList()) { row.Delete(); }
        }

        /******************************************************************/
        private void add_operations(DateTime d)
        {
            if (operations != null)
            {
                // все строки позиции на дату
                var pos = from r in T.table.AsEnumerable() where r.Field<DateTime>("Date") == d select r;

                // всe операции на дату сгрупированные по Дате, Площадке, Валюте
                var op = from r in operations.AsEnumerable()
                         where r.Field<DateTime>("Date") == d
                         group new { Amount = r.Field<double>("AmountIn") - r.Field<double>("AmountOut") }
                         by new { TradingSys = r.Field<string>("TradingSys"), Currency = r.Field<string>("Currency") }
                         into g select new { g.Key.TradingSys, g.Key.Currency, Amount = g.Sum(x => x.Amount) };

                // Запрос для апдейта существующих строк в позиции
                var j =
                    from p in pos join o in op
                     on new { TradingSys = p.Field<string>("TradingSys"), Currency = p.Field<string>("Instrument") } equals new { TradingSys = o.TradingSys, Currency = o.Currency }
                    select new { DataRow = p, Qty = p.Field<double>("Quantity") + o.Amount };
                foreach (var x in j) x.DataRow["Quantity"] = x.Qty;

                // Вставляем новые строки из операций
                var n = from o in op.AsEnumerable()
                        where !pos.AsEnumerable().Any(p => o.TradingSys == p.Field<String>("TradingSys") && o.Currency == p.Field<String>("Instrument"))
                        select o;
                foreach (var x in n) { add_row(new string[] { x.TradingSys, d.ToString(), x.Currency, "", x.Amount.ToString() }); }
            }
        }

        /******************************************************************/
        private void add_deals(DateTime d)
        {
            if (deals != null)
            {
                // все строки позиции на дату
                var pos = from r in T.table.AsEnumerable() where r.Field<DateTime>("Date") == d select r;

                // всe сделки на дату сгрупированные по Дате Поставки, Площадке, Инструменту
                var dd = from r in deals.AsEnumerable()
                         where r.Field<DateTime>("DateSettlement") == d
                         group new { Quantity = r.Field<double>("Quantity") * (r.Field<string>("Type") == "Продажа" ? -1 : 1) }
                         by new { TradingSys = r.Field<string>("TradingSys"), Instrument = r.Field<string>("Instrument") }
                         into g select new { g.Key.TradingSys, g.Key.Instrument, Quantity = g.Sum(x => x.Quantity) };

                // Запрос для апдейта существующих строк в позиции
                var j =
                    from p in pos join o in dd
                    on new { TradingSys = p.Field<string>("TradingSys"), Instrument = p.Field<string>("Instrument") } equals new { TradingSys = o.TradingSys, Instrument = o.Instrument }
                    select new { DataRow = p, Qty = p.Field<double>("Quantity") + o.Quantity };
                foreach (var x in j) x.DataRow["Quantity"] = x.Qty;

                // Вставляем новые строки из сделок по новым инструментам
                var n = from o in dd.AsEnumerable()
                        join i in instruments.AsEnumerable() on o.Instrument equals i.Field<string>("Instrument")
                        where !pos.AsEnumerable().Any(p => o.TradingSys == p.Field<String>("TradingSys") && o.Instrument == p.Field<String>("Instrument"))
                        select new
                        {
                            TradingSys = o.TradingSys,
                            Instrument = o.Instrument,
                            ISIN = i.Field<string>("ISIN"),
                            Quantity = o.Quantity
                        };
                foreach (var x in n) { add_row(new string[] { x.TradingSys, d.ToString(), x.Instrument, x.ISIN , x.Quantity.ToString() }); }
            }
        }

        /******************************************************************/
        private void add_operations_other(DateTime d)
        {
            if (operations_other != null)
            {
                // все строки позиции на дату
                var pos = from r in T.table.AsEnumerable() where r.Field<DateTime>("Date") == d select r;

                // Перевод ЦБ
                var o1 = from r in operations_other.AsEnumerable()
                         where r.Field<DateTime>("DateOperation") == d && r.Field<String>("Type").Equals("Перевод ЦБ")
                         group new { Quantity = r.Field<double>("Quantity") }
                         by new { TradingSys = r.Field<string>("TradingSys"), Instrument = r.Field<string>("Instrument") }
                         into g
                         select new { g.Key.TradingSys, g.Key.Instrument, Quantity = g.Sum(x => x.Quantity) };
                // Погашение ЦБ
                var o2 = from r in operations_other.AsEnumerable()
                         where r.Field<DateTime>("DateOperation") == d && r.Field<String>("Type").Equals("Погашение ЦБ")
                         group new { Quantity = r.Field<double>("Quantity") * (-1) }
                         by new { TradingSys = r.Field<string>("TradingSys"), Instrument = r.Field<string>("Instrument") }
                         into g
                         select new { g.Key.TradingSys, g.Key.Instrument, Quantity = g.Sum(x => x.Quantity) };
                // UNION
                var op = o1.Union(o2);

                // Запрос для апдейта существующих строк в позиции
                var j =
                    from p in pos join o in op
                    on new { TradingSys = p.Field<string>("TradingSys"), Instrument = p.Field<string>("Instrument") } equals new { TradingSys = o.TradingSys, Instrument = o.Instrument }
                    select new { DataRow = p, Qty = p.Field<double>("Quantity") + o.Quantity };
                foreach (var x in j) x.DataRow["Quantity"] = x.Qty;

                // Вставляем новые строки из операций по новым инструментам
                var n = from o in op.AsEnumerable()
                        join i in instruments.AsEnumerable() on o.Instrument equals i.Field<string>("Instrument")
                        where !pos.AsEnumerable().Any(p => o.TradingSys == p.Field<String>("TradingSys") && o.Instrument == p.Field<String>("Instrument"))
                        select new
                        {
                            TradingSys = o.TradingSys,
                            Instrument = o.Instrument,
                            ISIN = i.Field<string>("ISIN"),
                            Quantity = o.Quantity
                        };
                foreach (var x in n) { add_row(new string[] { x.TradingSys, d.ToString(), x.Instrument, x.ISIN, x.Quantity.ToString() }); }
            }
        }

        /******************************************************************/
        public void calculate()
        {
            DateTime dt_begin;
            DateTime dt_end;

            instruments = (from t in va.tables_list where t.Name == "instruments" select t.table).ToArray().FirstOrDefault();
            money = (from t in va.tables_list where t.Name == "money" select t.table).ToArray().FirstOrDefault();
            portfolio = (from t in va.tables_list where t.Name == "portfolio" select t.table).ToArray().FirstOrDefault();
            operations = (from t in va.tables_list where t.Name == "operations" select t.table).ToArray().FirstOrDefault();
            deals = (from t in va.tables_list where t.Name == "deals" select t.table).ToArray().FirstOrDefault();
            operations_other = (from t in va.tables_list where t.Name == "operations_other" select t.table).ToArray().FirstOrDefault();

            dt_begin = va.report_dates[0];
            dt_end = va.report_dates[1];

            /*----------------------------------------------------------------------------*/
            // position on date begin
            // portfolio
            var instruments_quey =
                from r in portfolio.AsEnumerable()
                join i in instruments.AsEnumerable()
                on r.Field<string>("Instrument") equals i.Field<string>("Instrument")
                where r.Field<double>("PeriodBeginQuantity") != 0
                select new
                {
                    TradingSys = r.Field<string>("TradingSys"),
                    Instrument = r.Field<string>("Instrument"),
                    ISIN = i.Field<string>("ISIN"),
                    PeriodBeginQuantity = r.Field<double>("PeriodBeginQuantity"),
                };
            foreach (var o in instruments_quey) { add_row(new string[] { o.TradingSys, dt_begin.ToString(), o.Instrument, o.ISIN, o.PeriodBeginQuantity.ToString() }); }

            // money (grouped by TradingSys)
            var money_query =
                from r in money.AsEnumerable()
                where r.Field<double>("PeriodBegin") != 0
                group new { Qty = r.Field<double>("PeriodBegin") }
                by new
                {
                    TradingSys = r.Field<string>("TradingSys").Replace("Основной счет,", "").Replace("Торговый счет,", "").Trim(),
                    Currency = r.Field<string>("Currency")
                }
                into g
                select new
                {
                    g.Key.TradingSys,
                    g.Key.Currency,
                    Qty = g.Sum(x => x.Qty),
                };
            foreach (var o in money_query) { add_row(new string[] {o.TradingSys, dt_begin.ToString() , o.Currency, "", o.Qty.ToString()}); }

            // add today operations
            add_operations(dt_begin);
            add_operations_other(dt_begin);
            // add today deals
            add_deals(dt_begin);
            /*----------------------------------------------------------------------------*/

            /*----------------------------------------------------------------------------*/
            // all dates to ds_end
            for (DateTime d = dt_begin.AddDays(1); d <= dt_end; d=d.AddDays(1))
            //for (DateTime d = dt_begin.AddDays(1); d <= dt_begin.AddDays(2); d = d.AddDays(1))
            {
                // копируем все строки позиции из предыдущего дня
                // сначала в List, т.к. модифицировать объект внутри foreach нельзы
                List<string[]> s = new List<string[]>();
                var pos = from r in T.table.AsEnumerable() where r.Field<DateTime>("Date") == d.AddDays(-1) select r;
                foreach (var x in pos)
                {
                    s.Add(new string[] { x.Field<string>("TradingSys"), d.ToString(), x.Field<string>("Instrument"), x.Field<string>("ISIN"), x.Field<double>("Quantity").ToString() });
                }                    
                // теперь переносим строки из списка в таблицу позиции
                foreach (string[] ss in s)
                {
                    add_row(ss);
                }
                // add today operations
                add_operations(d);
                add_operations_other(d);
                // add today deals
                add_deals(d);
                // delete rows wit Quantity = 0
                delete_row(d);
            }
            /*----------------------------------------------------------------------------*/

        }

    }
}
