using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WooCommerceClient.Services.SpecificHandling
{
    public class VBSConvert
    {
        public static List<IOrderConvert> OrderConverts = new List<IOrderConvert>();

        internal static void OrderLineConvert(Models.OrderLineItem line, Models.Order order, Models.Visma.VismaOrderLine vline, Models.Visma.VismaOrder vorder)
        {
            foreach (var orderConvert in OrderConverts)
            {
                orderConvert.ConvertOrderLine(line, order, vline, vorder);
            }
        }
    }
}
