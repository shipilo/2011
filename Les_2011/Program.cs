using System;
using System.Collections.Generic;
using System.IO;

namespace Les_2011
{
    class Program
    {
        static string path_in = "data.txt";
        static string path_out = "output.txt";
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

            foreach (int key in data.Keys)
            {
                Console.WriteLine($"{key}. {data[key]}");
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
    }
}
