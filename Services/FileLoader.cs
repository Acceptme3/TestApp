
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using TestApp.Models;

namespace TestApp.Services
{
    internal class FileLoader
    {
       public static string? FilePathField { get; private set; }

        public static ExcelDocument SetDataFilePath()
        {
            //Console.Write("Введите путь к файлу Excel: ");
            try
            {
                string? excelFilePath = "c:\\1.xlsx"; //////Console.ReadLine();
                FilePathField = excelFilePath;
                ExcelDocument workbook = new ExcelDocument (FileLoader.LoadExcelFile(excelFilePath!));

                if (workbook.Workbook != null)
                {
                    Console.WriteLine("Файл Excel успешно загружен.");
                    return workbook;
                }
                else
                {
                    Console.WriteLine("Не удалось загрузить файл Excel.");
                    return new ExcelDocument();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Произошла ошибка: {ex.Message}");
                return new ExcelDocument();
            }
        }
    
        static IXLWorkbook LoadExcelFile(string filePath)
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
