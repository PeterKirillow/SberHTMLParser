using System;
using HtmlAgilityPack;
using System.Data;
using System.Linq;
using System.Configuration;

namespace SberHTMLParser
{
    class Table
    {
        public DataTable table = null;  //
        public string Name;             //
        public string WorksheetName;    //
        public int Order;               // порядок следвания таблиц при создании экселя
        int skip = 1;                   // сколько строк заголовка таблицы пропускаем

        /******************************************************************/
        public Table(string name)
        {
            this.Name = name;
            this.Order = 100;
            this.table = new DataTable(name);

            DataColumn column = null;

            // создание колонок в таблице на основании данных из конфигурационнго файла
            string s_skip = ConfigurationManager.AppSettings.Get(this.Name + ":skip");
            if (s_skip != null) { this.skip = Convert.ToInt32(s_skip); }

            string s_order = ConfigurationManager.AppSettings.Get(this.Name + ":order");
            if (s_order != null) { this.Order = Convert.ToInt32(s_order); }

            string s_wsname = ConfigurationManager.AppSettings.Get(this.Name + ":caption");
            if (s_wsname != null && s_wsname != "") { this.WorksheetName = s_wsname; } else { this.WorksheetName = this.Name;  }

            string[] v = ConfigurationManager.AppSettings.AllKeys.Where(key => key.StartsWith(this.Name + ".")).Select(key => ConfigurationManager.AppSettings[key]).ToArray(); ;
            foreach (string s in v)
            {
                var e = s.Split(';');
                column = new DataColumn
                {
                    ColumnName = e[0],
                    DataType = System.Type.GetType($"System.{e[1]}"),
                    Caption = e[2].Replace("{NewLine}", Environment.NewLine)
                };
                // добавляем формулу, если она есть в описании колонки
                if (e.Length > 3) {
                    int ii = 1;
                    foreach( string ss in v)
                    {
                        var ee = ss.Split(';');
                        e[3] = e[3].Replace("{"+ee[0]+"}","RC"+ii.ToString());
                        ii++;
                    }
                    column.ExtendedProperties.Add("Formula", "=" + e[3]);
                }
                this.table.Columns.Add(column);
            }
        }
        /*----------*/

        /******************************************************************/
        public void addrows_from_nodes(HtmlNodeCollection n)
        {
            int shift = 0;  /* если таблица разделена на части, то наименование части
                             * (например имя площадки в сделках или тип репо в табличке с репо)
                             * помещаем в дополнительную колонку в начало массива row
                             * по этой причине к самому массиву применяем сдвиг, когда конвертируем значения в тип, определенный в колонках
                             */
            String[] add_row = { }; /*массив, который будет добавлен в начало массива row, если определим, что таблица раздалена на части*/

            // получаем все строки, кроме заголовков
            var rows = n.Skip(this.skip).Select(tr => tr.Elements("td").Select(td => td.InnerText.Trim()).ToArray());

            foreach (string[] row in rows)
            {
                // все Итого и пустые строки игнорируем
                if (!row[0].Contains("Итого") && !row[0].Contains("Общий итог") && !(row.Length == 1 && row[0] == ""))
                {
                    /* предположительно, если строка состоит из одного элемента,
                    *  то это подзаголовок и нужно сформировать из этого значения дополнительную колонку
                    */
                    if (row.Length == 1)
                    {
                        add_row = new string[] { row[0].Replace("Площадка:", "").Trim() };
                        shift = 1;
                        continue;
                    }

                    // конвертация строки в double / datetime
                    for (int i = 0; i <= table.Columns.Count - 1; i++)
                    {
                        if (table.Columns[i].DataType == System.Type.GetType("System.Double"))
                        {
                            if (row[i - shift] == "") { row[i - shift] = "0"; }    // если должно быть число, то "пусто" заменяем на 0
                            row[i - shift] = row[i - shift].Replace(" ", "");       // если число разделено пробелами, убираем его
                        }
                        if (table.Columns[i].DataType == System.Type.GetType("System.DateTime"))
                        {
                            if (row[i - shift] == "") { row[i - shift] = "01.01.1970"; }    // если должна быть дата, то "пусто" заменяем на '1970-01-01'
                        }
                    }

                    // либо добавляем новую колонку в начало строки, либо нет
                    if (shift != 0) { table.Rows.Add(add_row.Concat(row).ToArray()); } else { table.Rows.Add(row); }

                }
            }
        }
        /*----------*/
    }
}
