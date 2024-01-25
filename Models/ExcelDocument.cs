using ClosedXML.Excel;


namespace TestApp.Models
{
    internal class ExcelDocument : IDisposable
    {
        public IXLWorkbook? Workbook { get; set; }

        public ExcelDocument()
        {
                
        }

        public ExcelDocument(IXLWorkbook? workbook) 
        {
            Workbook = workbook;
        }

        public ExcelDocument GetInstance() 
        {
            return this;
        }
        public void Dispose()
        {
            Workbook?.Dispose();
        }
    }
}
