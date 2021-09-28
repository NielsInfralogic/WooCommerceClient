using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WooCommerceClient.Models.Visma
{
    // Gr12:
    public enum OrderStatusCodeVisma { OrderPending = 0, OrderProcessing = 1, OrderOnHold = 3, PaymentOK = 5, PaymentFailed = 9, Transporting = 10, ReadyForPickup = 11, OrderRefunded = 90, OrderCompleted = 100, OrderCollected = 101, OrderCancelled = 999 };

    public class VismaOrder
    {
        public int CustomerNo { get; set; } = 0;
        public int OrderNo { get; set; } = 0;
        public int OrderDate { get; set; } = 0;

        public int DeliveryMethod { get; set; } = 0;
        public int DeliveryTerms { get; set; } = 0;
        public int RequiredDeliveryDate { get; set; } = 0;

        public string DeliveryName { get; set; } = "";
        public string DeliveryAddress1 { get; set; } = "";
        public string DeliveryAddress2 { get; set; } = "";
        public string DeliveryAddress3 { get; set; } = "";
        public string DeliveryAddress4 { get; set; } = "";
        public string DeliveryPostCode { get; set; } = "";
        public string DeliveryPostalArea { get; set; } = "";
        public int DeliveryCountryNumber { get; set; } = 45;

        // CustNo properties
        public string CompanyName { get; set; } = "";
        public string Name { get; set; } = "";
        public string AddressLine1 { get; set; } = "";
        public string AddressLine2 { get; set; } = "";
        public string AddressLine3 { get; set; } = "";
        public string AddressLine4 { get; set; } = "";
        public string PostCode { get; set; } = "";
        public string PostalArea { get; set; } = "";
        public int CountryNumber { get; set; } = 45;
        public string Phone { get; set; } = "";
        public string Email { get; set; } = "";

        public string YourReference { get; set; } = "";

        public int CurrencyNo { get; set; } = 0;

        public int LanguageNo { get; set; } = 45;

        public string CardType { get; set; } = "";
        public int PaymentState { get; set; } = 0;
        public int PaymentMethod { get; set; } = 0;

        public string Information1 { get; set; } = "";
        public string Information2 { get; set; } = "";
        public string CustomerOrSupplierOrderNo { get; set; } = "";

        public string Information3 { get; set; } = "";

        public string Information4 { get; set; } = "";
        public string Information5 { get; set; } = "";
         public string Information6 { get; set; } = "";
        public string UserName { get; set; } = "";      //

        public string WebPage { get; set; } = "";

        public List<VismaOrderLine> OrderLines;

        public int Group1 { get; set; } = 1;
        public int Group12 { get; set; } = 1;
        public string InvoiceNo { get; set; } = "";
        public int InvoiceDate { get; set; } = 0;

        public int WarehouseNo { get; set; } = 30;

        // 20180831
        public decimal Freight1 { get; set; } = 0.0M;

        public string Memo { get; set; } = "";

        public decimal Free1 { get; set; } = 0.0M;

        public string CardNm { get; set; } = "";
        public string CardAc { get; set; } = "";

        public int Status { get; set; } = 0;
        public int OrderPreferences { get; set; } = 0;
    }

    public class VismaOrderLine
    {
        public int LineNo { get; set; } = 0;
        public string ProductNo { get; set; } = "";
        public string Description { get; set; } = "";
        public decimal Quantity { get; set; } = 0.0M;
        public int Units { get; set; } = 1;
        public decimal PriceInCurrency { get; set; } = 0.0M;
        public decimal LinePrice { get; set; } = 0.0M;
        public int Currency { get; set; } = 0;
        public decimal CostPrice { get; set; } = 0.0M;
        public decimal VatAmount { get; set; } = 0.0M;
        public decimal Discount { get; set; } = 0.0M;
        public decimal DiscountPercent { get; set; } = 0.0M;
        public string Comment { get; set; } = "";
        public string CategoryName { get; set; } = "";
        public int VatCode { get; set; } = 0;

        public decimal PickedQuantity { get; set; } = 0.0M;
        public decimal InvoicedAmountTotalInCurrency { get; set; } = 0.0M;

        public int WarehouseNo { get; set; } = 0;

        public bool SetProcessingMethod { get; set; } = false;
    }
}
