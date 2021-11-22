using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace Les_2011
{
    class Program
    {
        static string path_in = "data.txt";
        static string path_out = "output.txt";
        static string path_excel1 = "table1.xlsx";
        static string path_excel2 = "table2.xlsx";
        static Random random;
        public static Dictionary<int, Student> data;

        static Program()
        {
            random = new Random();
            data = new Dictionary<int, Student>();
        }

        static void Main(string[] args)
        {
            Console.WriteLine("Билетная лотерея.");

            ReadFromPath();

            foreach (int key1 in data.Keys)
            {
                Console.WriteLine($"{key1}. {data[key1]}");
            }

            LimitQueue<Draw> draws = new LimitQueue<Draw>();

            Console.WriteLine("\n1 - начать, 0 - выйти:");
            while (Console.ReadLine() == "1")
            {
                Console.Write("Название: ");
                string name = Console.ReadLine();
                Console.Write("Количество билетов: ");
                int num_tickets;
                if (!int.TryParse(Console.ReadLine(), out num_tickets))
                {
                    num_tickets = 0;
                }

                Console.Write("Количество студентов: ");
                int num_students;
                if (!int.TryParse(Console.ReadLine(), out num_students))
                {
                    num_students = 0;
                }

                List<int> exist = new List<int>();

                Console.WriteLine($"Номера студентов через интер (в количестве {num_students}):");
                for (int i = 0; i < num_students; i++)
                {
                    int number;
                    if (int.TryParse(Console.ReadLine(), out number) && !exist.Contains(number))
                    {
                        exist.Add(number);
                    }
                    else
                    {
                        Console.WriteLine("Неккоректный ввод.");
                        i = num_tickets;
                    }
                }
                Console.WriteLine("\nНажмите интер, чтобы провести розыгрыш.");
                if (Console.ReadKey().Key == ConsoleKey.Enter)
                {
                    foreach (Draw draw in draws)
                    {
                        foreach (int index in draw.Winners)
                        {
                            data[index].Ratio *= 0.5;
                        }
                    }

                    double ratio_sum = 0;
                    foreach (int index in exist)
                    {
                        ratio_sum += data[index].Ratio;
                    }

                    Stack<int> winNumbers = new Stack<int>();
                    for (int i = 0; i < num_tickets; i++)
                    {
                        Dictionary<double, int> pool = new Dictionary<double, int>();

                        double sum = 0;
                        foreach (int index in exist)
                        {
                            sum += data[index].Ratio / ratio_sum;
                            pool.Add(sum, index);
                        }

                        double winNumber = random.NextDouble();
                        foreach (double range in pool.Keys)
                        {
                            if (winNumber < range)
                            {
                                winNumbers.Push(pool[range]);
                                exist.Remove(winNumbers.Peek());
                                ratio_sum -= data[winNumbers.Peek()].Ratio;
                                break;
                            }
                        }
                    }
                    draws.Enqueue(new Draw(name, num_tickets, winNumbers));
                    WriteToPath(draws.Peek().ToString());
                }

                Console.WriteLine("Информация по трём последним розыгрышам:");
                foreach(Draw draw in draws)
                {
                    Console.WriteLine(draw);                    
                }

                Console.WriteLine("1 - продолжить, 0 - выйти:");
            }
            
            Console.WriteLine("Работа с файлами");

            int columnOfKey = 0;
            string key = "диагноз";
            object[,] table = ReadExcel(path_excel1, "A1:B11");
            string[] match1 = GetColumnFromExcelTable(table, 1);
            string[] match2 = GetColumnFromExcelTable(table, 2);
            string[,] result = new string[36, 4];
            table = ReadExcel(path_excel2, "A1:H36");
            string[] titles = GetRowFromExcelTable(table, 1);
            string[] id = GetColumnFromExcelTable(table, 1);
            for (int i = 1; i < titles.Length; i++)
            {
                if (titles[i].ToLower().Equals(key))
                {
                    columnOfKey = i;
                    i = titles.Length;
                }
            }
            if (columnOfKey == 0) Environment.Exit(0);

            string[] match3 = GetColumnFromExcelTable(table, columnOfKey);
            for (int i = 2; i < match3.Length; i++)
            {
                bool done = false;
                for (int j = 2; j < match1.Length; j++)
                {
                    if (match3[i].ToLower().Contains(match1[j].ToLower()))
                    {
                        result[i, 1] = id[i];
                        result[i, 2] = match3[i];
                        if (result[i, 3] != null) result[i, 3] += "\n" + match2[j];
                        else result[i, 3] = match2[j];
                        done = true;
                    }                    
                }
                if(!done)
                {
                    result[i, 1] = id[i];
                    result[i, 2] = match3[i];
                    result[i, 3] = "";
                }
            }
            WriteToExcel(result);
        }

        static void ReadFromPath()
        {
            StreamReader sr = new StreamReader(path_in);
            string[] sin = sr.ReadToEnd().Trim().Split(Convert.ToChar("\n"));
            sr.Close();
            int k = 0;
            foreach (string line in sin)
            {
                try
                {
                    data.Add(k++, new Student(line.Trim().Split()[0], Convert.ToInt32(line.Trim().Split()[1])));
                }
                catch 
                {
                    k--;
                }
            }
        }

        static void WriteToPath(string str)
        {
            string sout = "";
            if (File.Exists(path_out))
            {
                StreamReader sr = new StreamReader(path_out);
                sout = sr.ReadToEnd();
                sr.Close();
            }
            StreamWriter sw = new StreamWriter(path_out);
            sw.Write(sout + str);
            sw.Close();
        }

        static object[,] ReadExcel(string path, string area)
        {
            Excel.Application excel = new Excel.Application();
            excel.DisplayAlerts = false;
            Excel.Workbook book = excel.Workbooks.Open($@"{Environment.CurrentDirectory}\{path}", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Excel.Worksheet sheet = (Excel.Worksheet)book.Sheets[1];
            object[,] table = (object[,])sheet.Range[area].Value;
            excel.Quit();
            return table;
        }

        static string[] GetColumnFromExcelTable(object[,] table, int column)
        {           
            string[] column_array = new string[table.GetLength(0)];
            for (int i = 1; i < table.GetLength(0); i++)
            {
                if (table[i, column] != null)
                {
                    column_array[i] = table[i, column].ToString();
                }
                else
                {
                    i = table.GetLength(0);
                }
            }
            
            return column_array;
        }

        static string[] GetRowFromExcelTable(object[,] table, int row)
        {
            string[] row_array = new string[table.GetLength(1)];
            for (int i = 1; i < table.GetLength(1); i++)
            {
                if (table[row, i] != null)
                {
                    row_array[i] = table[row, i].ToString();
                }
                else
                {
                    i = table.GetLength(1);
                }
            }

            return row_array;
        }

        static void WriteToExcel(string[,] table)
        {            
            Excel.Application excel = new Excel.Application();
            excel.Visible = true;
            excel.SheetsInNewWorkbook = 2;
            Excel.Workbook workBook = excel.Workbooks.Add(Type.Missing);
            excel.DisplayAlerts = false;
            Excel.Worksheet sheet = (Excel.Worksheet)excel.Worksheets.get_Item(1);
            sheet.Name = "Результат";
            sheet.Cells[1, 1] = "ID";
            sheet.Cells[1, 2] = "Диагноз";
            sheet.Cells[1, 3] = "Лекарство(а)";
            for (int i = 2; i < table.GetLength(0); i++)
            {
                for (int j = 1; j < table.GetLength(1); j++)
                {
                    sheet.Cells[i, j] = table[i, j];
                }
            }
            excel.Application.ActiveWorkbook.SaveAs($@"{Environment.CurrentDirectory}\table3.xlsx", Type.Missing,
  Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
  Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            excel.Quit();
        }
    }
}
