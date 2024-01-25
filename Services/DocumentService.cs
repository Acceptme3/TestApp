using TestApp.Services.Interfaces;

namespace TestApp.Services
{
    internal class DocumentService : IDocumentService
    {
        public IXLWorkbook? Workbook { get; set; }     
            
        public void LoadDocument()
        {
            SetDataFilePath();
        }
        public void SetDataFilePath()
        {
            Console.Write("Введите путь к файлу Excel: ");
            try
            {
                string? excelFilePath = Console.ReadLine();
                Workbook = LoadExcelFile(excelFilePath!);

                if (Workbook != null)
                {
                    Console.WriteLine("Файл Excel успешно загружен.");
                }
                else
                {
                    Console.WriteLine("Не удалось загрузить файл Excel.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Произошла ошибка: {ex.Message}");
            }
        }

        IXLWorkbook LoadExcelFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Файл {filePath} не найден.");
                return null!;
            }
            try
            {
                var workbook = new XLWorkbook(filePath);
                return workbook;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при загрузке файла Excel: {ex.Message}");
                return null!;
            }
        }

    }

}
