using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WooCommerceClient.Models;
using WooCommerceClient.Models.Visma;

namespace WooCommerceClient.Services.SpecificHandling
{
    public class Lieu_Dit : IOrderConvert
    {
        private static readonly string defaultEmptyProdNo = Utils.ReadConfigString("EmptyProductNumber", "Diverse");

        public void ConvertOrderLine(OrderLineItem line, Order order, VismaOrderLine vline, VismaOrder vorder)
        {
            if (line != null && vline != null && 
                !string.IsNullOrEmpty(vline.Description) && 
                !string.IsNullOrEmpty(vline.ProductNo) &&
                vline.ProductNo != defaultEmptyProdNo)
                vline.Description = "";
        }
    }
}
