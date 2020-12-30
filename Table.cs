using System;
using HtmlAgilityPack;
using System.Data;
using System.Linq;

namespace SberHTMLParser
{
    class Table
    {
        public DataTable table = null;
        public string Name;
        public string WorksheetName;
        public int Order;

        public Table(string name, string ws_name, HtmlNodeCollection n, int order)
        {
            this.Name = name;
            this.WorksheetName = ws_name;
            this.Order = order;
            this.table = new DataTable(name);

            // определение типов в колонках таблицы
            if (this.Name.Equals("operations"))
            {
                table.Columns.Add("Дата", typeof(DateTime));
                table.Columns.Add("Торговая площадка", typeof(String));
                table.Columns.Add("Описание операции", typeof(String));
                table.Columns.Add("Валюта", typeof(String));
                table.Columns.Add("Сумма зачисления", typeof(double));
                table.Columns.Add("Сумма списания", typeof(double));
            }
            else if (this.Name.Equals("deals"))
            {
                table.Columns.Add("Площадка", typeof(String));
                table.Columns.Add("Дата заключения", typeof(DateTime));
                table.Columns.Add("Дата расчетов", typeof(DateTime));
                table.Columns.Add("Время заключения", typeof(String));
                table.Columns.Add("Наименование ЦБ", typeof(String));
                table.Columns.Add("Код ЦБ", typeof(String));
                table.Columns.Add("Валюта", typeof(String));
                table.Columns.Add("Вид", typeof(String));
                table.Columns.Add("Количество, шт.", typeof(double));
                table.Columns.Add("Цена", typeof(double));
                table.Columns.Add("Сумма", typeof(double));
                table.Columns.Add("НКД", typeof(double));
                table.Columns.Add("Комиссия Брокера", typeof(double));
                table.Columns.Add("Комиссия Биржи", typeof(double));
                table.Columns.Add("Номер сделки", typeof(String));
                table.Columns.Add("Комментарий", typeof(String));
                table.Columns.Add("Статус сделки", typeof(String));
            }
            else if (this.Name.Equals("valuation"))
            {
                table.Columns.Add("Торговая площадка", typeof(String));
                table.Columns.Add("Оценка портфеля ЦБ, руб.", typeof(double));
                table.Columns.Add("Денежные средства, руб.", typeof(double));
                table.Columns.Add("Оценка, руб.", typeof(double));
            }
            else if (this.Name.Equals("instruments"))
            {
                table.Columns.Add("Наименование", typeof(String));
                table.Columns.Add("Код", typeof(String));
                table.Columns.Add("ISIN ценной бумаги", typeof(String));
                table.Columns.Add("Эмитент", typeof(String));
                table.Columns.Add("Вид, Категория, Тип, иная информация", typeof(String));
                table.Columns.Add("Выпуск, Транш, Серия", typeof(String));
            }
            else if (this.Name.Equals("repo"))
            {
                table.Columns.Add("Тип РЕПО", typeof(String));
                table.Columns.Add($"Дата{Environment.NewLine}заключения", typeof(DateTime));
                table.Columns.Add($"Время{Environment.NewLine}заключения", typeof(String));
                table.Columns.Add($"Наименование{Environment.NewLine}ЦБ", typeof(String));
                table.Columns.Add($"Код{Environment.NewLine}ЦБ", typeof(String));
                table.Columns.Add($"Валюта", typeof(String));
                table.Columns.Add($"Вид", typeof(String));
                table.Columns.Add($"Количество,{Environment.NewLine}шт.", typeof(double));
                table.Columns.Add($"Цена 1-й{Environment.NewLine}части", typeof(double));
                table.Columns.Add($"НКД по 1-й{Environment.NewLine}части.", typeof(double));
                table.Columns.Add($"Сумма по 1-й{Environment.NewLine}части", typeof(double));
                table.Columns.Add($"Дата{Environment.NewLine}исполнения{Environment.NewLine}1-й части", typeof(DateTime));
                table.Columns.Add($"Ставка{Environment.NewLine}РЕПО,{Environment.NewLine}%", typeof(double));
                table.Columns.Add($"Процент{Environment.NewLine}РЕПО", typeof(double));
                table.Columns.Add($"Цена 2-й{Environment.NewLine}части", typeof(double));
                table.Columns.Add($"НКД по 2-й{Environment.NewLine}части", typeof(double));
                table.Columns.Add($"Сумма по 2-й{Environment.NewLine}части", typeof(double));
                table.Columns.Add($"Дата{Environment.NewLine}исполнения{Environment.NewLine}2-й части", typeof(DateTime));
                table.Columns.Add($"Комиссия{Environment.NewLine}Брокера", typeof(double));
                table.Columns.Add($"Комиссия{Environment.NewLine}Биржи", typeof(double));
                table.Columns.Add($"Номер{Environment.NewLine}сделки", typeof(String));
                table.Columns.Add($"Статус{Environment.NewLine}сделки", typeof(String));
            }
            else if (this.Name.Equals("money"))
            {
                table.Columns.Add("Торговая площадка", typeof(String));
                table.Columns.Add("Валюта", typeof(String));
                table.Columns.Add("Курс на конец периода", typeof(double));
                table.Columns.Add("Начало периода", typeof(double));
                table.Columns.Add("Изменение за период", typeof(double));
                table.Columns.Add("Конец периода", typeof(double));
                table.Columns.Add($"Плановые зачисления{Environment.NewLine}по сделкам", typeof(double));
                table.Columns.Add($"Плановые списания{Environment.NewLine}по сделкам", typeof(double));
                table.Columns.Add($"Плановый исходящий{Environment.NewLine}остаток", typeof(double));
            }
            else if (this.Name.Equals("portfolio"))
            {
                table.Columns.Add("Площадка", typeof(String));

                table.Columns.Add("Наименование", typeof(String));
                table.Columns.Add($"ISIN ценной{Environment.NewLine}бумаги", typeof(String));
                table.Columns.Add($"Валюта{Environment.NewLine}рыночной{Environment.NewLine}цены", typeof(String));
                // Начало периода
                table.Columns.Add($"Начало периода{Environment.NewLine}Количество, шт", typeof(double));
                table.Columns.Add($"Начало периода{Environment.NewLine}Номинал", typeof(double));
                table.Columns.Add($"Начало периода{Environment.NewLine}Рыночная цена", typeof(double));
                table.Columns.Add($"Начало периода{Environment.NewLine}Рыночная стоимость,{Environment.NewLine}без НКД", typeof(double));
                table.Columns.Add($"Начало периода{Environment.NewLine}НКД", typeof(double));
                // Конец периода
                table.Columns.Add($"Конец периода{Environment.NewLine}Количество, шт", typeof(double));
                table.Columns.Add($"Конец периода{Environment.NewLine}Номинал", typeof(double));
                table.Columns.Add($"Конец периода{Environment.NewLine}Рыночная цена", typeof(double));
                table.Columns.Add($"Конец периода{Environment.NewLine}Рыночная стоимость,{Environment.NewLine}без НКД", typeof(double));
                table.Columns.Add($"Конец периода{Environment.NewLine}НКД", typeof(double));
                // Изменение за период	
                table.Columns.Add($"Изменение за период{Environment.NewLine}Количество, шт", typeof(double));
                table.Columns.Add($"Изменение за период{Environment.NewLine}Рыночная стоимость", typeof(double));
                // Плановые показатели
                table.Columns.Add($"Плановые зачисления{Environment.NewLine}по сделкам, шт", typeof(double));
                table.Columns.Add($"Плановые списания{Environment.NewLine}по сделкам, шт", typeof(double));
                table.Columns.Add($"Плановый исходящий{Environment.NewLine}остаток, шт", typeof(double));
            }
            else if (this.Name.Equals("money_out"))
            {
                table.Columns.Add("Дата", typeof(DateTime));
                table.Columns.Add("Торговая площадка", typeof(String));
                table.Columns.Add("Описание операции", typeof(String));
                table.Columns.Add("Валюта", typeof(String));
                table.Columns.Add("Сумма", typeof(double));
            }

            int shift = 0;  /* если таблица разделена на части, то наименование части
                             * (например имя площадки в сделках или тип репо в табличке с репо)
                             * помещаем в дополнительную колонку в начало массива row
                             * по этой причине к самому массиву применяем сдвиг, когда конвертируем значения в тип, определенный в колонках
                             */
            String[] add_row = { }; /*массив, который будет добавлен в начало массива row, если определим, что таблица раздалена на части*/

            int skip = 1;   // сколько строк заголовка таблицы пропускаем
            if (this.Name.Equals("portfolio")) { skip = 2; }

            // получаем все строки, кроме заголовков
            var rows = n.Skip(skip).Select(tr => tr.Elements("td").Select(td => td.InnerText.Trim()).ToArray());

            foreach (string[] row in rows)
            {
                // все Итого и пустые строки игнорируем
                if (!row[0].Contains("Итого") || (row.Length == 1 && row[0] == "") )
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

                    // конвертация строки в double
                    for (int i = 0; i <= table.Columns.Count - 1; i++)
                    {
                        if (table.Columns[i].DataType == System.Type.GetType("System.Double")) {
                            if (row[i - shift] == "") { row[i - shift] = "0";  }    // если должно быть число, то пусто заменяем на 0
                            row[i - shift] = row[i - shift].Replace(" ", "");       // если число разделено пробелами, убираем его
                        }
                    }

                    // либо добавляем новую колонку в начало строки, либо нет
                    if ( shift != 0 )
                    {
                        table.Rows.Add(add_row.Concat(row).ToArray());
                    } else
                    {
                        table.Rows.Add(row);
                    }
                }
            }
        }
    }
}
