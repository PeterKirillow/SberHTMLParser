using HtmlAgilityPack;
using System;
using System.Data;
using System.Linq;
using ClosedXML.Excel;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Text;

namespace SberHTMLParser
{
    class Program
    {
        private static string html;

        /**************************************************************************/
        // create array of tables from html
        public static string[] ParseHtmlSplitTables()
        {
            string[] result = new string[] { };
            if (!String.IsNullOrWhiteSpace(html))
            {
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(html);
                var tableNodes = doc.DocumentNode.SelectNodes("//table");
                if (tableNodes != null)
                {
                    result = Array.ConvertAll<HtmlNode, string>(tableNodes.ToArray(), n => n.OuterHtml);
                }
            }
            return result;
        }

        /**/
        static int Main(string[] args)
        {
            HtmlDocument doc;
            string in_file = "";
            string out_file = "";
            string[] htmlTables;
            List<DateTime> report_dates = new List<DateTime>();
            List<Table> tables_list = new List<Table>();

            /*----------------------------------------------------------------------------*/
            if (args.Length == 0)
            {
                Console.WriteLine("SberHTMLParser <path to Html file>");
                return 1;
            } else
            {
                in_file = args[0];
                if ( !File.Exists(in_file) )
                {
                    Console.WriteLine($"File {in_file} does not exists");
                    return 1;
                }
            }
            out_file = in_file.Replace(Path.GetExtension(in_file), ".xlsx");
            /*----------------------------------------------------------------------------*/

            // загружаем документ в строку
            // изначально всё должно быть в кодировке UTF-8 и по-русски
            // но, если файл был сохранен в кодировке windows-1251, то это надо бы явно указать
            // считываем файл в строку сначала без преобразования и попробуем найти фразу "Отчет брокера"
            // если не находим, то пробуем переключиться в windows-1251
            html = File.ReadAllText(in_file);
            if ( ! html.Contains("Отчет брокера", StringComparison.OrdinalIgnoreCase))
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                html = File.ReadAllText(in_file, Encoding.GetEncoding("windows-1251"));
                if (!html.Contains("Отчет брокера", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"Файл {in_file} в неизвестной кодировке");
                    return 1;
                }                   
            }
            // убираем из документа <TBODY> и </TBODY>. Нам при парсинге таблиц они не нужны. В разных вариантах брокерского отчета эти теги или есть или нет.
            html = html.Replace("<tbody>", "", StringComparison.OrdinalIgnoreCase).Replace("</tbody>", "", StringComparison.OrdinalIgnoreCase);

            // парсим документ и создаем массив из всех таблиц
            htmlTables = ParseHtmlSplitTables();

            // узнаем период отчета и дату создания. очень приблизительно предполагаем, что это строка 
            // <br>за период с dd.MM.YYYY по dd.MM.YYYY, дата создания dd.MM.YYYY</br>
            string pattern = "дата создания";
            if (Regex.IsMatch(html, pattern))
            {
                Match match = Regex.Matches(html, pattern).First();
                string s = html.Substring(match.Index-30, 62);

                var regex = new Regex(@"\b\d{2}\.\d{2}.\d{4}\b");
                foreach (Match m in regex.Matches(s))
                {
                    DateTime dt;
                    if (DateTime.TryParseExact(m.Value, "dd.MM.yyyy", null, DateTimeStyles.None, out dt)) { report_dates.Add(dt); }
                }
            }
            // парсим каждую таблицу по отдельности
            foreach (var t in htmlTables)
            {
                doc = new HtmlDocument();
                doc.LoadHtml(t);
                var nodes = doc.DocumentNode.SelectNodes("//table/tr");

                // пытаемся понять, что это за табличка
                var header = nodes[0].Elements("td").Select(td => td.InnerText.Trim()).ToArray();

                // "Оценка активов"
                if (header.Length == 4 && header[0] == "Торговая площадка" && header[3] == "Оценка, руб.")
                {
                    tables_list.Add(new Table("valuation", "Оценка активов", nodes, 1));
                }
                // "Портфель Ценных Бумаг"
                else if (header.Length == 5 && header[0] == "" && header[4] == "Плановые показатели")
                {
                    tables_list.Add(new Table("portfolio", "Портфель Ценных Бумаг", nodes, 2));
                }
                // "Денежные средства"
                else if (header.Length == 9 && header[0] == "Торговая площадка" && header[8] == "Плановый исходящий остаток")
                {
                    tables_list.Add(new Table("money", "Денежные средства", nodes, 3));
                }
                // "Движение денежных средств за период"
                else if (header.Length == 6 && header[0] == "Дата" && header[5] == "Сумма списания")
                {
                    tables_list.Add(new Table("operations", "Движение денежных средств", nodes, 4));
                }
                // "Сделки купли/продажи ценных бумаг"
                else if (header.Length == 16 && header[0] == "Дата заключения" && header[15].Contains("Статус сделки"))
                {
                    tables_list.Add(new Table("deals", "Сделки купли продажи ЦБ", nodes, 5));
                }
                // "Справочник Ценных Бумаг"
                else if (header.Length == 6 && header[0] == "Наименование" && header[5] == "Выпуск, Транш, Серия")
                {
                    tables_list.Add(new Table("instruments", "Справочник ЦБ", nodes, 10));
                }
                // "Сделки РЕПО"
                else if (header.Length == 21 && header[0] == "Дата заключения" && header[11].Contains("РЕПО"))
                {
                    tables_list.Add(new Table("repo", "Сделки РЕПО", nodes, 8));
                }
                // "Выплаты дохода от эмитента на внешний счет"
                else if (header.Length == 5 && header[0] == "Дата" && header[4].Contains("Сумма"))
                {
                    tables_list.Add(new Table("money_out", "Выплаты дохода на внешний счет", nodes, 9));
                }
            }

            /*----------------------------------------------------------------------------*/
            // export to excel
            /*----------------------------------------------------------------------------*/
            var wb = new XLWorkbook();

            // worksheet "даты"
            IXLWorksheet ws;
            ws = wb.Worksheets.Add("даты");
            int i = 1;
            foreach (DateTime d in report_dates)
            {
                if (i == 1) { ws.Cell(i, 1).Value = "Начало периода";  }
                if (i == 2) { ws.Cell(i, 1).Value = "Конец периода";  }
                if (i == 3) { ws.Cell(i, 1).Value = "Дата создания отчета";  }
                ws.Cell(i, 2).Value = d.ToString();
                i++;
            }
            ws.Columns().AdjustToContents();

            // worksheets from array of tables
            foreach (Table t in tables_list.OrderBy(o => o.Order))
            {
                ws = wb.Worksheets.Add(t.table, t.WorksheetName);
                for (int a=0; a<= t.table.Columns.Count - 1; a++)
                {
                    if (t.table.Columns[a].DataType == System.Type.GetType("System.Double")) {
                        ws.Columns(a+1,a+1).Style.NumberFormat.Format = "#,##0.00";
                    }                    
                }
                ws.Row(1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                ws.Row(1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Top);
                ws.Tables.FirstOrDefault().ShowAutoFilter = false;
                ws.Columns().AdjustToContents();
            }

            // save file
            wb.SaveAs(out_file);
            /*----------------------------------------------------------------------------*/

            return 0;
        }
    }
}
