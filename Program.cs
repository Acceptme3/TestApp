global using ClosedXML.Excel;
using TestApp.Models;
using TestApp.Services;
using TestApp.Services.Interfaces;
using System;

namespace TestApp
{
    class Program
    {
        public static ExcelDocument document = new ExcelDocument();
        static void Main()
        {
            Console.WriteLine("Добро пожаловать в TestApp");

            while (true)
            {
                Console.WriteLine("\nВыберите действие:");
                Console.WriteLine("1. Ввести путь до файла с данными");
                Console.WriteLine("2. Вывести информацию о клиентах, заказавших товар");
                Console.WriteLine("3. Изменить контактное лицо клиента");
                Console.WriteLine("4. Определить золотого клиента");
                Console.WriteLine("5. Выйти из приложения");

                int choice;
                if (!int.TryParse(Console.ReadLine(), out choice))
                {
                    Console.WriteLine("Некорректный ввод. Пожалуйста, введите число от 1 до 5.");
                    continue;
                }

                switch (choice)
                {
                    case 1:
                        document = FileLoader.SetDataFilePath();
                        break;
                    case 2:
                        DisplayInfo();
                        ITableRepository rep = new TableRepository(document);
                        ProductManager p = new ProductManager(document, rep);
                        p.GetInformation();
                        break;
                    case 3:
                        ITableRepository clientRep = new TableRepository(document);
                        ClientManager c = new ClientManager(document, clientRep);
                        c.ClientUpdate();
                        break;
                    /*case 4:
                        FindGoldenCustomer(); 
                        break;*/
                    case 5:
                        document.Dispose();
                        Environment.Exit(0);
                        break;
                    default:
                        Console.WriteLine("Некорректный ввод. Пожалуйста, выберите число от 1 до 5.");
                        break;
                }
            }
        }
        private static void DisplayInfo()
        {

            if (document.Workbook == null)
            {
                Console.WriteLine("Данные не загружены. Введите путь до файла с данными.");
            }

        }

    }
}


