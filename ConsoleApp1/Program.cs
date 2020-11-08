using System;
using System.Linq;
using System.Collections.Generic;
using System.Data;
using System.IO;
using NPOI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Newtonsoft.Json;


namespace ConsoleApp1
{
    class ContractPoint
    {
        public string ContractPointName { get; set; } // Название пункта таблицы
        public string ContractPointData { get; set; } // Значение пункта таблицы 
    }

    class Program
    {
        static void Main(string[] args)
        {
            string filename = "testDoc_text_format.txt";
            string[] necessary_point_names = new string[] { "Номер договора", "Наименование договора", "Регистрационный номер сделки", "Адрес контрагента", "Счет контрагента", };

            // Конвертация rtf файла в формат txt
            var document = new Aspose.Words.Document("testDoc.rtf");
            document.Save(filename, Aspose.Words.SaveFormat.Text);

            // Список для хранения пунктов таблицы и их значений
            List<ContractPoint> list = new List<ContractPoint>();

            // Заполнение списка
            using (StreamReader reader = new StreamReader(filename, System.Text.Encoding.Default))
            {
                int i = 0;
                while (!reader.EndOfStream)
                {
                    string str = reader.ReadLine();
                    foreach (string point_name in necessary_point_names)
                    {
                        if (str.Contains(point_name))
                        {
                            list.Add(new ContractPoint() { ContractPointName = point_name, ContractPointData = reader.ReadLine() });
                        }
                    }

                }
            }

            // Конвертация данных в DataTable
            DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(list), (typeof(DataTable)));
            
            // Создание и запись excel файла
            using (var filestream = new FileStream("result.xlsx", FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet excelsheet = workbook.CreateSheet("Sheet1");
                List<String> columns = new List<string>();

                int columnIndex = 0;
                foreach (System.Data.DataColumn column in table.Columns)
                {

                    columns.Add(column.ColumnName);
                    columnIndex++;
                }


                IRow row = excelsheet.CreateRow(0);
                int cellIndex = 0;

                foreach (String col in columns) 
                {
                    foreach (DataRow dsrow in table.Rows)
                    {
                        row.CreateCell(cellIndex).SetCellValue(dsrow[col].ToString());
                        cellIndex++;
                    }

                    if (row.RowNum >= 1)
                    {
                        break;
                    }

                    cellIndex = 0;
                    row = excelsheet.CreateRow(1);
                }
                workbook.Write(filestream);
            }
        }
    }
}
