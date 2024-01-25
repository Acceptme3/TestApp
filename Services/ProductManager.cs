using TestApp.Models;
using TestApp.Services.Interfaces;

namespace TestApp.Services
{
    internal class ProductManager
    {
        private readonly ITableRepository _repository;
        private ExcelDocument _document;

        

        public ProductManager(ExcelDocument document, ITableRepository repository)
        {
            _document = document;
            _repository = repository;
        }

        public void GetInformation()
        {

            Console.WriteLine("Введите наименование товара");
            Console.WriteLine();

            try
            {
                string? name = Console.ReadLine();
                if (string.IsNullOrEmpty(name))
                {
                    throw new Exception("Вы ввели пустую строку");
                }

                if (name == "0")
                {
                    throw new Exception();
                }

                List<string> product = _repository.Search("Товары", name);

                if (product.Count == 0)
                {
                    throw new Exception("Такого товара не существует");
                }

                List<string> requisition = _repository.Search("Заявки", product[0]);

                if (product.Count == 0)
                {
                    throw new Exception("Заявок на данный товар не существует");
                }

                List<string> client = _repository.Search("Клиенты", requisition[2]);

                if (product.Count == 0)
                {
                    throw new Exception("Ошибка поиска клиента.");
                }

                Console.WriteLine($"Код товара ( {name} ) {product[0]}");
                Console.WriteLine($"Номер заявки {requisition[3]}, колличество {requisition[4]}, стоимость заказа{(int.Parse(product[3]) * int.Parse(requisition[4])).ToString()}");
                Console.WriteLine($"Заказчик: {client[1]}, Адрес: {client[2]}, Контактное лицо: {client[3]}");

            }
            catch (Exception ex)
            {
                Console.WriteLine($"{ex.Message}");
            }

        }

    }
}
