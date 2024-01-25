using TestApp.Models;
using TestApp.Services.Interfaces;

namespace TestApp.Services
{
    internal class ClientManager
    {
        private readonly ITableRepository _repository;
        private ExcelDocument _document;


        public ClientManager(ExcelDocument document, ITableRepository repository)
        {
            _document = document;
            _repository = repository;
        }

        public void ClientUpdate() 
        {
            Console.WriteLine("Введите название организации: ");
            Console.WriteLine();
            try 
            {
                string? name = Console.ReadLine();
                Console.WriteLine();

                if (string.IsNullOrEmpty(name))
                {
                    throw new Exception("Вы ввели пустую строку");
                }

                int rowIndex = _repository.FindRowIndexByValue("Клиенты", name);
                if (rowIndex<1)
                {
                    throw new Exception("Такой организации не существует");
                }
                Console.WriteLine("Клиент найден! Введите индекс поля которое хотите изменить (1 - имя организации, 2 - Адрес, 3 - Контактное лицо):");
                Console.WriteLine();

                name = Console.ReadLine();
                Console.WriteLine();

                if (string.IsNullOrEmpty(name))
                {
                    throw new Exception("Вы ввели пустую строку");
                }
                int columnIndex = int.Parse(name);

                Console.WriteLine("Введите новое значение:");
                Console.WriteLine();

                name = Console.ReadLine();
                Console.WriteLine();

                if (string.IsNullOrEmpty(name))
                {
                    throw new Exception("Вы ввели пустую строку");
                }
               _repository.SetCellValue("Клиенты", rowIndex, columnIndex, name);

                _repository.SaveChanges(FileLoader.FilePathField!);

                Console.WriteLine("Данные успешно обновлены");
                Console.WriteLine();
            }
            catch (Exception ex) 
            {
                Console.WriteLine($"{ex.Message}");
                Console.WriteLine();
            }
        }
    }
}
