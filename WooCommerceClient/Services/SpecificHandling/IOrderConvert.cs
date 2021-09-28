using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WooCommerceClient.Services.SpecificHandling
{
    public interface IOrderConvert
    {
        void ConvertOrderLine(Models.OrderLineItem line, Models.Order order, Models.Visma.VismaOrderLine vline, Models.Visma.VismaOrder vorder);
    }
}
