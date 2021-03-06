﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

/*
 * 
 */

namespace SberHTMLParser
{
    class PL
    {
        public Table T;

        Variables va;
        Prices prices;

        DataTable instruments;
        DataTable money_out;
        DataTable portfolio;
        DataTable operations;
        DataTable deals;
        DataTable operations_other;
        DataTable position;

        String Instrument = "";
        String ISIN = "";
        String PortfolioCurrency = "";
        String Currency = "";

        DateTime MinDate = new DateTime();
        double QtyBegin = 0;
        double AmountBegin = 0;
        double NominalBegin = 0;
        double AvgPriceBegin = 0;

        double BuyQty = 0;
        double BuyAmount = 0;
        double SellQty = 0;
        double SellAmount = 0;
        double CommBrok = 0;
        double CommExch = 0;
        double Coupons = 0;
        double Dividends = 0;
        double Tax = 0;
        double Repaiment = 0;

        DateTime MaxDate = new DateTime();
        double QtyEnd = 0;
        double AmountEnd = 0;

        bool IsBond = false;

        /******************************************************************/
        public PL(Variables va, Prices prices)
        {
            this.va = va;
            this.prices = prices;
            T = new Table("PL");
        }

        /******************************************************************/
        private void clear_vars()
        {
            NominalBegin = 0;
            AvgPriceBegin = 0;
            AmountBegin = 0;
            AmountEnd = 0;
            QtyBegin = 0;
            QtyEnd = 0;
            BuyQty = 0;
            BuyAmount = 0;
            SellQty = 0;
            SellAmount = 0;
            CommBrok = 0;
            CommExch = 0;
            Repaiment = 0;
            Coupons = 0;
            Dividends = 0;
            Tax = 0;
        }

        /******************************************************************/
        private void add_row()
        {
            T.table.Rows.Add(new string[] {
                Instrument,
                ISIN,
                Currency,

                MinDate.ToString(),
                QtyBegin.ToString(),
                AmountBegin.ToString(),

                NominalBegin.ToString(),
                AvgPriceBegin.ToString(),

                BuyQty.ToString(),
                BuyAmount.ToString(),
                SellQty.ToString(),
                SellAmount.ToString(),
                CommBrok.ToString(),
                CommExch.ToString(),
                Coupons.ToString(),
                Dividends.ToString(),
                Tax.ToString(),
                Repaiment.ToString(),

                MaxDate.ToString(),
                QtyEnd.ToString(),
                AmountEnd.ToString()
            });
        }

        /******************************************************************/
        private void position_from_portfolio(String instr, String curr)
        {
            // если есть позиция на начало периода в отчете в таблице "Портфель Ценных Бумаг",
            // берем данные из нее
            if (portfolio != null)
            {
                var x = from r in portfolio.AsEnumerable()
                        where r.Field<String>("Instrument") == instr && r.Field<String>("Currency") == curr
                        group new
                        {
                            AmountBegin = r.Field<double>("PeriodBeginAmount") + r.Field<double>("PeriodBeginAccrued"),
                            AmountEnd = r.Field<double>("PeriodEndAmount") + r.Field<double>("PeriodEndAccrued"),
                            QtyBegin = r.Field<double>("PeriodBeginQuantity"),
                            QtyEnd = r.Field<double>("PeriodEndQuantity"),
                            PeriodBeginNominal = r.Field<double>("PeriodBeginNominal")
                        }
                        by new { Instrument = r.Field<string>("Instrument"), Currency = r.Field<string>("Currency") }
                        into g
                        select new
                        {
                            g.Key.Instrument,
                            g.Key.Currency,
                            AmountBegin = g.Sum(x => x.AmountBegin),
                            QtyBegin = g.Sum(x => x.QtyBegin),
                            AmountEnd = g.Sum(x => x.AmountEnd),
                            QtyEnd = g.Sum(x => x.QtyEnd),
                            PeriodBeginNominal = g.Min(x => x.PeriodBeginNominal)
                        };
                if (x.Any())
                {
                    AmountBegin = x.First().AmountBegin;
                    QtyBegin = x.First().QtyBegin;
                    AmountEnd = x.First().AmountEnd;
                    QtyEnd = x.First().QtyEnd;
                    NominalBegin = x.First().PeriodBeginNominal;
                    if (AmountBegin != 0)
                    {
                        AvgPriceBegin = Math.Abs(AmountBegin / QtyBegin / (IsBond ? NominalBegin : 1) * (IsBond ? 100 : 1));
                    }

                }
            }
        }

        /******************************************************************/
        private void add_inoperations(String instr, String mode)
        {
            DateTime dt_begin = (mode == "begin" ? MinDate : MinDate.AddDays(1));
            DateTime dt_end = (mode == "begin" ? MinDate : MaxDate);

            double Price;
            String PriceCurrency;
            double Nominal;
            String NominalCurrency;
            double Accrued;

            if (operations_other != null)
            {
                /*
                 * проходим по каждой строчке ввода
                 * если это первая дата, то заполняем только данные на начало
                 * иначе суммируем вводы в кол-вах и в деньгах
                 * пробуем получить на каждый день цену из таблички Prices, если они там есть. Если цена найдена и валюта цены совпадает с валютой строки позиции, то считаем Amounnt
                 */
                var x = from r in operations_other.AsEnumerable()
                        join i in instruments.AsEnumerable() on r.Field<string>("Instrument") equals i.Field<string>("Instrument")
                        where r.Field<DateTime>("DateOperation") >= dt_begin && r.Field<DateTime>("DateOperation") <= dt_end
                                && r.Field<String>("Instrument") == instr
                                && r.Field<String>("Type").Contains("Перевод ЦБ")
                        select new
                        {
                            Date = r.Field<DateTime>("DateOperation"),
                            Instrument = r.Field<string>("Instrument"),
                            ISIN = i.Field<string>("ISIN"),
                            QtyBegin = (mode == "begin" ? r.Field<double>("Quantity") : 0),
                            QtyBuy = (mode == "begin" ? 0 : r.Field<double>("Quantity"))
                        };
                foreach (var s in x)
                {
                    String[] s_prices = prices.get_price(s.ISIN, s.Date);
                    QtyBegin = QtyBegin + s.QtyBegin;
                    BuyQty = BuyQty + s.QtyBuy;
                    if ((s_prices[0] ?? "") != "")
                    {
                        Price = Convert.ToDouble(s_prices[0]);
                        PriceCurrency = s_prices[1];
                        Nominal = Convert.ToDouble(s_prices[2]);
                        NominalCurrency = s_prices[3];
                        Accrued = Convert.ToDouble(s_prices[4]);
                        if (this.Currency == (IsBond ? NominalCurrency : PriceCurrency))
                        {
                            if (mode == "begin")
                            {
                                AmountBegin = (QtyBegin * Price * (IsBond ? Nominal : 1) / (IsBond ? 100 : 1) + (IsBond ? Accrued : 0));
                                NominalBegin = Nominal;
                                AvgPriceBegin = Price;
                            }
                            else
                            {
                                BuyAmount = (BuyQty * Price * (IsBond ? Nominal : 1) / (IsBond ? 100 : 1) + (IsBond ? Accrued : 0));
                            }
                        }
                        Console.WriteLine($"Ввод инструмента {Instrument} - ({s.Date.ToShortDateString()}, price: {Price}, pricecurrency: {PriceCurrency}, nominal: {Nominal}, nominalcurrency: {NominalCurrency})");
                    }
                    else
                    {
                        Console.WriteLine($"Ввод инструмента {Instrument} на дату {s.Date.ToShortDateString()} цена не найдена");
                    }
                }
            }
        }

        /******************************************************************/
        private void add_deals(String instr, String curr, String mode)
        {
            /* если первая позиция по инструменту образовалась от сделки
             * стоимость позиции на начало складывается как Amount - Comissions - Accrued
             * дальнгейшие сделки раскидываем по колонкам
            */
            DateTime dt_begin = (mode == "begin" ? MinDate : MinDate.AddDays(1));
            // MaxDate - напомнгю себе - это последняя дата позиции по инструменту. А значит сделка полной продажи пройдет "завтра".
            // А купоны так вообще придут неизвестно когда, но тут это не важно. Но пусть будет.
            DateTime dt_end = (mode == "begin" ? MinDate : MaxDate.AddDays(30));

            if (deals != null)
            {
                var x = from r in deals.AsEnumerable()
                        join i in instruments.AsEnumerable() on r.Field<string>("Instrument") equals i.Field<string>("Instrument")
                        //join c in rates.AsEnumerable() on r.Field<string>("Currency") equals c.Field<string>("Instrument")
                        where r.Field<DateTime>("DateSettlement") >= dt_begin && r.Field<DateTime>("DateSettlement") <= dt_end
                              && r.Field<String>("Instrument") == instr && r.Field<String>("Currency") == curr
                        group new
                        {
                            Buy = r.Field<double>("Amount") * (r.Field<string>("Type") == "Покупка" ? 1 : 0),
                            QtyBuy = r.Field<double>("Quantity") * (r.Field<string>("Type") == "Покупка" ? 1 : 0),
                            Sell = (r.Field<double>("Amount")) * (r.Field<string>("Type") == "Продажа" ? 1 : 0),
                            QtySell = r.Field<double>("Quantity") * (r.Field<string>("Type") == "Продажа" ? 1 : 0),
                            AccruedBuy = r.Field<double>("Accrued") * (r.Field<string>("Type") == "Покупка" ? 1 : 0),
                            AccruedSell = r.Field<double>("Accrued") * (r.Field<string>("Type") == "Продажа" ? 1 : 0),
                            CommBrok = r.Field<double>("CommBrok"),
                            CommExch = r.Field<double>("CommExch"),
                            QtySumm = r.Field<double>("Quantity") * (r.Field<string>("Type") == "Покупка" ? 1 : -1),
                            Nominal = r.Field<double>("Amount") / r.Field<double>("Quantity") / r.Field<double>("Price") * (IsBond ? 100 : 1)
                        }
                         by new
                         {
                             Instrument = r.Field<string>("Instrument"),
                             Currency = r.Field<string>("Currency")
                         }
                        into g
                        select new
                        {
                            g.Key.Currency,
                            g.Key.Instrument,
                            Buy = g.Sum(x => x.Buy),
                            BuyQty = g.Sum(x => x.QtyBuy),
                            Sell = g.Sum(x => x.Sell),
                            SellQty = g.Sum(x => x.QtySell),
                            CommBrok = g.Sum(x => x.CommBrok),
                            CommExch = g.Sum(x => x.CommExch),
                            AccruedBuy = g.Sum(x => x.AccruedBuy),
                            AccruedSell = g.Sum(x => x.AccruedSell),
                            QtySumm = g.Sum(x => x.QtySumm),
                            Nominal = g.Min(x => x.Nominal)
                        };

                if (x.Any())
                {
                    if (mode.Equals("begin"))
                    {
                        AmountBegin = AmountBegin + x.First().Buy - x.First().Sell - x.First().CommBrok - x.First().CommExch + x.First().AccruedSell - x.First().AccruedBuy;
                        NominalBegin = x.First().Nominal;
                        QtyBegin = QtyBegin + x.First().QtySumm;
                        AvgPriceBegin = Math.Abs(AmountBegin / QtyBegin / NominalBegin * (IsBond ? 100 : 1));
                    }
                    else
                    {
                        BuyQty = BuyQty + x.First().BuyQty;
                        BuyAmount = BuyAmount + x.First().Buy;
                        SellQty = SellQty + x.First().SellQty;
                        SellAmount = SellAmount + x.First().Sell;
                        CommBrok = CommBrok + x.First().CommBrok;
                        CommExch = CommExch + x.First().CommExch;
                    }
                }
            }
        }

        /******************************************************************/
        private void add_coupons(String instr, String curr)
        {
            Match match;
            // зачисление купонов на брокерский счет
            if (operations != null)
            {
                string pattern_coupon = $"Зачисление д/с.*купон.*{instr}";
                var x = from r in operations.AsEnumerable()
                        where r.Field<String>("Currency") == curr && r.Field<String>("Description").Contains(instr)
                        select new { Description = r.Field<String>("Description"), Amount = r.Field<double>("AmountIn") };
                foreach (var xx in x)
                {
                    match = Regex.Matches(xx.Description, pattern_coupon).FirstOrDefault();
                    if (match != null) { Coupons = Coupons + xx.Amount; }
                }
            }
        }

        /******************************************************************/
        private void add_coupons_ext(String instr, String curr)
        {
            Match match;
            // выводы купонов на внешний счет
            if (money_out != null)
            {
                string pattern_coupon = $"Зачисление д/с.*купон.*{instr}";
                var x = from r in money_out.AsEnumerable()
                        where r.Field<String>("Currency") == curr && r.Field<String>("Description").Contains(instr)
                        select new { Description = r.Field<String>("Description"), Amount = r.Field<double>("Amount") };
                foreach (var xx in x)
                {
                    match = Regex.Matches(xx.Description, pattern_coupon).FirstOrDefault();
                    if (match != null) { Coupons = Coupons + xx.Amount; }
                }
            }
        }

        /******************************************************************/
        private void add_dividends(String curr)
        {
            Match match;
            //зачисление дживидендов на брокерский счет
            if (operations != null)
            {
                string pattern_div = $"Дивиденды по акциям";
                var x = from r in operations.AsEnumerable()
                        where r.Field<String>("Currency") == curr
                        select new { Description = r.Field<String>("Description"), Amount = r.Field<double>("AmountIn") };
                foreach (var xx in x)
                {
                    match = Regex.Matches(xx.Description, pattern_div).FirstOrDefault();
                    if (match != null) { Dividends = Dividends + xx.Amount; }
                }
            }
        }

        /******************************************************************/
        private void add_repaiments(String instr, String curr)
        {
            Match match;
            // погашения
            if (operations != null)
            {
                string pattern_repaiment = $"Зачисление д/с.*погашение.*{instr}";
                var x = from r in operations.AsEnumerable()
                        where r.Field<String>("Currency") == curr && r.Field<String>("Description").Contains(instr)
                        select new { Description = r.Field<String>("Description"), Amount = r.Field<double>("AmountIn") };

                foreach (var xx in x)
                {
                    match = Regex.Matches(xx.Description, pattern_repaiment).FirstOrDefault();
                    if (match != null) { Repaiment = Repaiment + xx.Amount; }
                }
            }
        }

        /******************************************************************/
        private void add_tax(String curr)
        {
            Match match;
            // налоги
            if (operations != null)
            {
                string pattern_repaiment = $"Списание налога за налоговый период";
                var x = from r in operations.AsEnumerable()
                        where r.Field<String>("Currency") == curr
                        select new { Description = r.Field<String>("Description"), Amount = r.Field<double>("AmountOut") };

                foreach (var xx in x)
                {
                    match = Regex.Matches(xx.Description, pattern_repaiment).FirstOrDefault();
                    if (match != null) { Tax = Tax + xx.Amount; }
                }
            }
        }


        /******************************************************************/
        public void calculate()
        {
            // DataTables
            instruments = (from t in va.tables_list where t.Name == "instruments" select t.table).ToArray().FirstOrDefault();
            portfolio = (from t in va.tables_list where t.Name == "portfolio" select t.table).ToArray().FirstOrDefault();
            operations = (from t in va.tables_list where t.Name == "operations" select t.table).ToArray().FirstOrDefault();
            deals = (from t in va.tables_list where t.Name == "deals" select t.table).ToArray().FirstOrDefault();
            operations_other = (from t in va.tables_list where t.Name == "operations_other" select t.table).ToArray().FirstOrDefault();
            position = (from t in va.tables_list where t.Name == "position" select t.table).ToArray().FirstOrDefault();
            money_out = (from t in va.tables_list where t.Name == "money_out" select t.table).ToArray().FirstOrDefault();

            /*----------------------------------------------------------------------------*/
            foreach (var i in instruments.AsEnumerable().OrderBy(x => x.Field<String>("Instrument")).ToArray())
            {
                Instrument = i.Field<string>("Instrument");
                ISIN = i.Field<string>("ISIN");
                IsBond = (i.Field<string>("Type").ToLower().Contains("облигация") ? true : false);

                bool contains = position.AsEnumerable().Any(row => Instrument == row.Field<String>("Instrument"));
                if (contains)
                {
                    MinDate = DateTime.Parse("01.01.1970");
                    MaxDate = DateTime.Parse("01.01.1970");

                    // минимальная и максимальня даты по инструменту в позиции. Т.к. берем из сформированной нами ранее позиции, то точно уверены в датах
                    MinDate = (from r in position.AsEnumerable() where r.Field<String>("Instrument") == Instrument select r.Field<DateTime>("Date")).Min();
                    MaxDate = (from r in position.AsEnumerable() where r.Field<String>("Instrument") == Instrument select r.Field<DateTime>("Date")).Max();

                    /* нам не нужен зоопарк с валютами в портфеле и движениях по инструменту
                     * например инструмент купили  на вгнебирже за USD, но в портфеле Сбер оценивает его в рублях
                     * мли гипотетически в портфеле на начало оценка в рублях, потом покупки в дроугой валюте, оценка на конец тоже в рублях
                     * надо привести к одной валюте - вероятно это будет та валюта, которая изначально есть в таблице "Портфель ЦБ"
                     * конвертация к валюте портфеля потребуется в сделках, операциях, выплатах на внешний счет
                     */
                    List<string> curr_List = new List<string>();
                    List<string> list = new List<string>();
                    // собираем весь вписок валют по инструменту
                    // из Портфеля
                    if (portfolio != null)
                    {
                        list = portfolio.AsEnumerable().Where(i => i["Instrument"].ToString() == Instrument).Select(c => c["Currency"].ToString()).Distinct().ToList();
                        curr_List = curr_List.Concat(list ?? new List<string>()).Distinct().ToList();
                        this.PortfolioCurrency = (list.Count != 0 ? curr_List[0] : "");
                    }
                    // из сделок
                    if (deals != null)
                    {
                        list = deals.AsEnumerable().Where(i => i["Instrument"].ToString() == Instrument).Select(x => x["Currency"].ToString()).Distinct().ToList();
                        curr_List = curr_List.Concat(list ?? new List<string>()).Distinct().ToList();
                    }
                    // из операций
                    if (operations != null)
                    {
                        list = operations.AsEnumerable().Where(i => i["Description"].ToString().Contains(Instrument)).Select(x => x["Currency"].ToString()).Distinct().ToList();
                        curr_List = curr_List.Concat(list ?? new List<string>()).Distinct().ToList();
                    }
                    // из выводов на внешний счет
                    if (money_out != null)
                    {
                        list = money_out.AsEnumerable().Where(i => i["Description"].ToString().Contains(Instrument)).Select(x => x["Currency"].ToString()).Distinct().ToList();
                        curr_List = curr_List.Concat(list ?? new List<string>()).Distinct().ToList();
                    }

                    // итак, если у нас всего одна валюта - это хорошо
                    // если больше, то смотрим, не заполниоли ли мы ранее валлюту из портфеля, если да, то оставляем ее, если нет, то делаем основной валютой первую из доступных
                    // вариант, когда валют больше 2-х, не хочу даже рассматривать.
                    // надой найти приемлемое решение
                    if (curr_List.Count == 1)
                    {
                        this.PortfolioCurrency = curr_List[0];
                    }
                    else
                    {
                        PortfolioCurrency = (PortfolioCurrency == "" ? curr_List[0] : PortfolioCurrency);
                        Console.WriteLine($"Инструмент {Instrument} представлен в нескольких валютах: {string.Join(',', curr_List)}");
                    }

                    foreach (var c in curr_List)
                    {
                        this.Currency = c;
                        clear_vars();

                        // загрузка из Портфеля ЦБ должно идти на первом месте
                        position_from_portfolio(Instrument, Currency);

                        // !! потенциально, если случились вводы инструмента, у которого несколько валют в движениях, то это задвоит значения ввода !!
                        // вопрос нужно решить принудительной конвертацией множества валют к единой
                        add_inoperations(Instrument, "begin");
                        add_inoperations(Instrument, "all");

                        add_deals(Instrument, Currency, "begin");
                        add_deals(Instrument, Currency, "all");
                        add_coupons(Instrument, Currency);
                        add_coupons_ext(Instrument, Currency);
                        add_repaiments(Instrument, Currency);

                        add_row();
                    }

                }

            }
            /*----------------------------------------------------------------------------*/

            /*----------------------------------------------------------------------------*/
            // добавляем строки по валютам, чтобы учесть доходы/расходы, которые невозможно привязать к инструменту
            // например Дивиденды и Налоги. Если они есть.
            if (operations != null)
            {
                var c_curr = operations.AsEnumerable().Select(row => new { Currency = row.Field<string>("Currency") }).Distinct();

                foreach (var cc in c_curr)
                {
                    this.Instrument = cc.Currency;
                    this.Currency = "";
                    this.ISIN = "";
                    clear_vars();
                    add_dividends(cc.Currency);
                    add_tax(cc.Currency);
                    add_row();
                }
            }
            /*----------------------------------------------------------------------------*/
        }
    }
}
