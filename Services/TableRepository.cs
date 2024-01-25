using TestApp.Models;
using TestApp.Services.Interfaces;

namespace TestApp.Services
{
   internal class TableRepository : ITableRepository
   {
        private IXLWorkbook _workbook;
        private ExcelDocument _document;

        public TableRepository (ExcelDocument document)
        {
            _document = document;
            _workbook = _document.Workbook!;
        }

        public List<string> GetRowData(string sheetName, int rowIndex)
        {
            try
            {
                var worksheet = _workbook.Worksheet(sheetName);
                var rowData = worksheet.Row(rowIndex).CellsUsed().Select(cell => cell.Value.ToString()).ToList();
                return rowData;
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка получения данных из строки: {ex.Message}");
            }
        }

        public List<string> GetColumnData(string sheetName, int columnIndex)
        {
            try
            {
                var worksheet = _workbook.Worksheet(sheetName);
                var columnData = worksheet.Column(columnIndex).CellsUsed().Select(cell => cell.Value.ToString()).ToList();
                return columnData;
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка получения данных из столбца: {ex.Message}");
            }
        }

        public List<List<string>> SearchValues(string sheetName, string searchValue)
        {
            try
            {
                var worksheet = _workbook.Worksheet(sheetName);
                var matchingRows = worksheet.RowsUsed().Where(row => row.CellsUsed().Any(cell => cell.Value.ToString().Contains(searchValue)))
                                                      .Select(row => row.CellsUsed().Select(cell => cell.Value.ToString()).ToList())
                                                      .ToList();
                return matchingRows;
            }
            catch (Exception ex)
            {

                throw new Exception($"Ошибка поиска значения: {ex.Message}");
            }
        }

        public List<string> Search(string sheetName, string searchValue)
        {
            try
            {
                var worksheet = _workbook.Worksheet(sheetName);
                var matchingRows = worksheet.RowsUsed()
                    .Where(row => row.CellsUsed().Any(cell => cell.Value.ToString().Contains(searchValue)))
                    .Select(row => row.CellsUsed().Select(cell => cell.Value.ToString()).ToList())
                    .FirstOrDefault();

                return matchingRows ?? new List<string>();
            }
            catch (Exception ex)
            {
                return new List<string>();
            }
        }

        public int FindRowIndexByValue(string sheetName, string searchValue)
        {
            try
            {
                var worksheet = _workbook.Worksheet(sheetName);
                var matchingRow = worksheet.RowsUsed()
                    .FirstOrDefault(row => row.CellsUsed().Any(cell => cell.Value.ToString().Contains(searchValue)));

                return matchingRow?.RowNumber() ?? -1; 
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка поиска значения в листе: {ex.Message}");
            }
        }
        public void SetCellValue(string sheetName, int rowIndex, int columnIndex, string cellValue)
        {
            try
            {
                var worksheet = _workbook.Worksheet(sheetName);
                var cellToModify = worksheet.Cell(rowIndex, columnIndex);
                cellToModify.Value = cellValue;
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка установки значения в ячейке: {ex.Message}");
            }
        }

        public void SaveChanges(string filePath)
        {
            try
            {
                _workbook.SaveAs(filePath);
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка сохранения изменений: {ex.Message}");
            }
        }
    }

}
