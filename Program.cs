using HtmlAgilityPack;
using System;
using System.Data;
using System.Linq;
using ClosedXML.Excel;
using System.Collections.Generic;

namespace SberHTMLParser
{
    class Program
    {

        /**************************************************************************/
        // create array of tables from gtml document
        public static string[] ParseHtmlSplitTables(string htmlString)
        {
            string[] result = new string[] { };
            if (!String.IsNullOrWhiteSpace(htmlString))
            {
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(htmlString);
                var tableNodes = doc.DocumentNode.SelectNodes("//table");
                if (tableNodes != null)
                {
                    result = Array.ConvertAll<HtmlNode, string>(tableNodes.ToArray(), n => n.OuterHtml);
                }
            }
            return result;
        }

        /**************************************************************************/
        static void Main(string[] args)
        {
            HtmlDocument doc;
            const string xlsFile = "C:\\Users\\peter\\OneDrive\\_data\\QUICK\\сбербанк\\40032HX\\40032HX_011020_311020_M.xlsx";

            // загружаем документ в строку
            htmlDocument h = new htmlDocument("C:\\Users\\peter\\OneDrive\\_data\\QUICK\\сбербанк\\40032HX\\40032HX_011020_311020_M.html");
            // парсим документ и создаем массив из всех таблиц
            string[] htmlTables = ParseHtmlSplitTables(h.html);

            var tables_list = new List<Table>();

            // парсим каждую таблицу по отдельности
            foreach (var t in htmlTables)
            {
                doc = new HtmlDocument();
                doc.LoadHtml(t);
                var nodes = doc.DocumentNode.SelectNodes("//table/tr");
                // пытаемся понять, что это за табличка
                var header = nodes[0].Elements("td").Select(td => td.InnerText.Trim()).ToArray();

                // таблица "Оценка активов"
                if (header.Length == 4 && header[0] == "Торговая площадка" && header[3] == "Оценка, руб.")
                {
                    tables_list.Add(new Table("valuation", "Оценка активов", nodes, 1));
                }
                // таблица "Портфель Ценных Бумаг"
                else if (header.Length == 5 && header[0] == "" && header[4] == "Плановые показатели")
                {
                    tables_list.Add(new Table("portfolio", "Портфель Ценных Бумаг", nodes, 2));
                }
                // таблица "Денежные средства"
                else if (header.Length == 9 && header[0] == "Торговая площадка" && header[8] == "Плановый исходящий остаток")
                {
                    tables_list.Add(new Table("money", "Денежные средства", nodes, 3));
                }
                // таблица "Движение денежных средств за период"
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
                    tables_list.Add(new Table("repo", "Сделки РЕПО", nodes, 9));
                }
            }

            // export to excel
            var wb = new XLWorkbook();
            foreach(Table t in tables_list.OrderBy(o => o.Order))
            {
                wb.Worksheets.Add(t.table, t.WorksheetName);
            }
            wb.SaveAs(xlsFile);
        }
    }
}
