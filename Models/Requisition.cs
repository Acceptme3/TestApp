using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestApp.Models
{
    internal class Requisition 
    {
        public string? Id { get; set; }
        public string? NumRequisition { get; set; }
        public string? ProductId { get; set; }
        public string? ClientId { get; set; }
        public string? CountProduct { get; set; }
        public string? PostingDate { get; set; }
    }
}
