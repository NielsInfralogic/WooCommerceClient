using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WooCommerceClient.Models.Visma
{
    public class VismaActor
    {
        public int CustomerNo { get; set; } = 0;
        public string Name { get; set; } = string.Empty;
        public string AddressLine1 { get; set; } = string.Empty;
        public string AddressLine2 { get; set; } = string.Empty;
        public string AddressLine3 { get; set; } = string.Empty;
        public string PostCode { get; set; } = string.Empty;
        public string PostalArea { get; set; } = string.Empty;
        public int CountryNo { get; set; } = 45;
        public string Phone { get; set; } = string.Empty;
        public string Mobile { get; set; } = string.Empty;

        public string EmailAddress { get; set; } = string.Empty;

        public string DeliveryName { get; set; } = string.Empty;
        public string DeliveryAddressLine1 { get; set; } = string.Empty;
        public string DeliveryAddressLine2 { get; set; } = string.Empty;
        public string DeliveryAddressLine3 { get; set; } = string.Empty;
        public string DeliveryPostCode { get; set; } = string.Empty;
        public string DeliveryPostalArea { get; set; } = string.Empty;
        public int DeliveryCountryNo { get; set; } = 45;
        public int AssociateNo { get; set; } = 0;

        public int LanguageNo { get; set; } = 45;
        public int CurrencyNo { get; set; } = 45;
    }
    
}
