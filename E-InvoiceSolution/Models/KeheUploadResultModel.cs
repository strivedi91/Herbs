using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace E_InvoiceSolution.Models
{
    public class KeheUploadResultModel
    {
        public KeheUploadResultModel()
        {
            this.ExtraShipped = new List<ItemDetails>();
            this.PurchaseOrderDetails = new List<PurchaseOrderDetails>();
            this.OrderedButNotReceived = new List<ItemDetails>();
            this.ShippedButNotOrdered = new List<ItemDetails>();
        }

        public int? POID { get; set; }
        public int? RowsDeletedFromDailyPO { get; set; }
        public int? RowsInExcelSheet { get; set; }
        public string RepatedSKUsFoundInExcelSheet { get; set; }
        public int? RowsImportedFromExcelSheet { get; set; }
        public int? RowsDeltedFromNaturesBestSupportingTable { get; set; }
        public int? SkusFoundIntblPurchaseOrderDetailsForPOID { get; set; }
        public int? QuantityFoundIntblPurchaseOrderDetailsForPOID { get; set; }
        public int? TotalSkusReceived { get; set; }
        public int? TotalQuantityReceived { get; set; }

        public List<PurchaseOrderDetails> PurchaseOrderDetails { get; set; }
        public List<ItemDetails> ExtraShipped { get; set; }
        public List<ItemDetails> ShippedButNotOrdered { get; set; }
        public List<ItemDetails> OrderedButNotReceived { get; set; }

        public string DownloadPath { get; set; }
        public string DownloadFileName { get; set; }
    }

    public class PurchaseOrderDetails
    {
        public int? POID { get; set; }
        public int? POCode { get; set; }
        public string InvoiceNumber { get; set; }
        public int? ProductCount { get; set; }
        public int? POQuantity { get; set; }
        public decimal POAmount { get; set; }
    }

    public class ItemDetails
    {
        public int? ProductTypeId { get; set; }
        public int? SKU { get; set; }
        public int? OrderedQty { get; set; }
        public int? ReceivedQty { get; set; }
        public int? ShippedQty { get; set; }
        public decimal UnitCost { get; set; }
        public decimal WholesalePrice { get; set; }
        public string InvoiceNumber { get; set; }
    }
}