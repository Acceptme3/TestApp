namespace TestApp.Services.Interfaces
{
    internal interface ITableRepository
    {
        List<string> GetRowData(string sheetName, int rowIndex);
        List<string> GetColumnData(string sheetName, int columnIndex);
        List<List<string>> SearchValues(string sheetName, string searchValue);
        List<string> Search(string sheetName, string searchValue);
        void SetCellValue(string sheetName, int rowIndex, int columnIndex, string cellValue);
        public int FindRowIndexByValue(string sheetName, string searchValue);
        void SaveChanges(string filePath);
    }
}
