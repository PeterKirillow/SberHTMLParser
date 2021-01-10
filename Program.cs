using ClosedXML.Excel;
using HtmlAgilityPack;
using System;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace SberHTMLParser
{
    class Program
    {
        // ссылка на класс с глобальными переменными
        public static Variables va = new Variables();
        // ссылка на класс Prices
        public static Prices prices;
        // массив из табличек документа - html текст
        static string[] htmlTables;

        /**************************************************************************/
        // создаем массив из найденных в документе табличек
        public static void ParseHtmlSplitTables()
        {
            if (!String.IsNullOrWhiteSpace(va.html))
            {
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(va.html);
                var tableNodes = doc.DocumentNode.SelectNodes("//table");
                if (tableNodes != null)
                {
                    htmlTables = Array.ConvertAll<HtmlNode, string>(tableNodes.ToArray(), n => n.OuterHtml);
                }
            }
        }

        /**************************************************************************/
        static int Main(string[] args)
        {
            HtmlDocument doc;
            string in_file = "";
            string price_file = "";
            string out_file = "";

            bool ShowAutoFilter = true;

            /*----------------------------------------------------------------------------*/
            if (args.Length == 0)
            {
                Console.WriteLine("SberHTMLParser <path to Html file> [path to prices file]");
                return 1;
            }
            else
            {
                // 
                in_file = args[0];
                if (!File.Exists(in_file))
                {
                    Console.WriteLine($"File {in_file} does not exists");
                    return 1;
                }

                // если есть файл с ценами, то будет пробовать заполнить таьлички Prices и Rates !
                // файл должен иметь определенную структуру.
                if (args.Length == 2)
                {
                    price_file = args[1];
                    if (File.Exists(price_file))
                    {
                        prices = new Prices(price_file);
                        prices.calculate();
                        if (!prices.I.IsEmpty) { va.tables_list.Add(prices.I); }
                        if (!prices.C.IsEmpty) { va.tables_list.Add(prices.C); }
                    }
                }
            }
            out_file = in_file.Replace(Path.GetExtension(in_file), ".xlsx");
            /*----------------------------------------------------------------------------*/

            // загружаем документ в строку
            // изначально всё должно быть в кодировке UTF-8 и по-русски
            // но, если файл был сохранен в кодировке windows-1251, то это надо бы явно указать
            // считываем файл в строку сначала без преобразования и попробуем найти фразу "Отчет брокера"
            // если не находим, то пробуем переключиться в windows-1251
            va.html = File.ReadAllText(in_file);
            if (!va.html.Contains("Отчет брокера", StringComparison.OrdinalIgnoreCase))
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                va.html = File.ReadAllText(in_file, Encoding.GetEncoding("windows-1251"));
                if (!va.html.Contains("Отчет брокера", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"Файл {in_file} в неизвестной кодировке");
                    return 1;
                }
            }
            // убираем из документа <TBODY> и </TBODY>. Нам при парсинге таблиц они не нужны. В разных вариантах брокерского отчета эти теги или есть или нет.
            va.html = va.html.Replace("<tbody>", "", StringComparison.OrdinalIgnoreCase).Replace("</tbody>", "", StringComparison.OrdinalIgnoreCase);

            // парсим документ и создаем массив из всех таблиц
            ParseHtmlSplitTables();

            // узнаем период отчета и дату создания. очень приблизительно предполагаем, что это строка 
            // <br>за период с dd.MM.YYYY по dd.MM.YYYY, дата создания dd.MM.YYYY</br>
            string pattern = "дата создания";
            if (Regex.IsMatch(va.html, pattern))
            {
                Match match = Regex.Matches(va.html, pattern).First();
                string s = va.html.Substring(match.Index - 30, 62);

                var regex = new Regex(@"\b\d{2}\.\d{2}.\d{4}\b");
                foreach (Match m in regex.Matches(s))
                {
                    DateTime dt;
                    if (DateTime.TryParseExact(m.Value, "dd.MM.yyyy", null, DateTimeStyles.None, out dt)) { va.report_dates.Add(dt); }
                }
            }

            /*----------------------------------------------------------------------------*/
            // парсим каждую таблицу по отдельности сверяясь со списком из конфигурационнго файла
            foreach (var t in htmlTables)
            {
                Table T = null;
                doc = new HtmlDocument();
                doc.LoadHtml(t);
                var nodes = doc.DocumentNode.SelectNodes("//table/tr");

                // список таблиц и параметры парсинга заголовков для определения имени таблицы
                string[] v = ConfigurationManager.AppSettings.AllKeys.Where(key => key.StartsWith("tables")).Select(key => ConfigurationManager.AppSettings[key]).ToArray(); ;
                foreach (string s in v)
                {
                    var e = s.Split(';');
                    string table_name = e[0];
                    int h_row = Convert.ToInt32(e[1]);
                    int h_length = Convert.ToInt32(e[2]);
                    string columns = e[3];
                    string columns_content = e[4];

                    var header = nodes[h_row].Elements("td").Select(td => td.InnerText.Trim()).ToArray();

                    if (header.Length == h_length)
                    {
                        int found = 0;
                        var c = columns.Split('#');
                        for (int ii = 1; ii <= c.Length; ii++)
                        {
                            if (header[Convert.ToInt32(c[ii - 1])].Contains(columns_content.Split('#')[ii - 1]))
                            {
                                found++;
                            }
                        }
                        if (found == c.Length)
                        {
                            T = new Table(table_name);
                        }
                    }
                }

                if (T != null && !T.IsEmpty)
                {
                    T.addrows_from_nodes(nodes);
                    va.tables_list.Add(T);
                }

            }
            /*----------------------------------------------------------------------------*/


            // worksheet "Позиция"
            Position p = new Position(va, prices);
            va.tables_list.Add(p.T);
            p.calculate();

            // worksheet "PL"
            PL pl = new PL(va, prices);
            va.tables_list.Add(pl.T);
            pl.calculate();


            /*----------------------------------------------------------------------------*/
            // export to excel
            /*----------------------------------------------------------------------------*/
            var wb = new XLWorkbook();

            // worksheet "Даты"
            IXLWorksheet ws;
            ws = wb.Worksheets.Add("Даты");
            int i = 1;
            foreach (DateTime d in va.report_dates)
            {
                if (i == 1) { ws.Cell(i, 1).Value = "Начало периода"; }
                if (i == 2) { ws.Cell(i, 1).Value = "Конец периода"; }
                if (i == 3) { ws.Cell(i, 1).Value = "Дата создания отчета"; }
                ws.Cell(i, 2).Value = d.ToString();
                i++;
            }
            ws.Columns().AdjustToContents();


            // worksheets from list of tables
            foreach (Table t in va.tables_list.OrderBy(o => o.Order))
            {

                if (!t.IsEmpty)
                {
                    ws = wb.Worksheets.Add(t.table, t.WorksheetName);

                    /*----------------------------------------------------------------------------*/
                    // добавление формул
                    int c_col = 1;
                    foreach (DataColumn c in t.table.Columns)
                    {
                        if (c.ExtendedProperties.Count != 0)
                        {
                            var rngData = ws.Range(2, c_col, t.table.Rows.Count + 1, c_col);
                            rngData.LastColumn().FormulaR1C1 = c.ExtendedProperties["Formula"].ToString();
                        }
                        c_col++;
                    }

                    /*----------------------------------------------------------------------------*/
                    // format columns
                    for (int a = 0; a <= t.table.Columns.Count - 1; a++)
                    {
                        if (t.table.Columns[a].DataType == System.Type.GetType("System.Double"))
                        {
                            ws.Columns(a + 1, a + 1).Style.NumberFormat.Format = "#,##0.00";
                        }
                    }

                    ws.Row(1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    ws.Row(1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Top);
                    ws.Tables.FirstOrDefault().ShowAutoFilter = ShowAutoFilter;
                    ws.Columns().AdjustToContents();
                }
            }

            // save file
            wb.SaveAs(out_file);
            /*----------------------------------------------------------------------------*/

            return 0;
        }
    }
}
