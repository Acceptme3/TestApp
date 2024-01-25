using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestApp.Services.Interfaces
{
    internal interface IDocumentService
    {
        public IXLWorkbook? Workbook { get; set; }
        void LoadDocument();
        
    }
}
