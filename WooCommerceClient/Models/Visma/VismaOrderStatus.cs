using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WooCommerceClient.Models.Visma
{

    public class VismaOrderStatus
    {
        public int OrderNumber { get; set; } = 0;
        public int OrderDocNumber { get; set; } = 0;
        public string CustomerOrSupplierOrderNo { get; set; } = "";
    }
}


