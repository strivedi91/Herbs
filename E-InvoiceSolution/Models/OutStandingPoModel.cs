using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace E_InvoiceSolution.Models
{
    public class OutStandingPoModel
    {
        public int POID { get; set; }
        public string POCode { get; set; }
        public DateTime orderdate { get; set; }
        public string po_comments { get; set; }
    }
}