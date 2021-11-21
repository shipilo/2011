using System;

namespace Met_2011
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Упражнение 11.");
            BankFabric.DeleteAccount(BankFabric.CreateAccount());

            Console.WriteLine("Домашнее задание 11.");
            Creator.DeleteAccount(BankFabric.CreateAccount());

            Console.ReadLine();
        }
    }
}
