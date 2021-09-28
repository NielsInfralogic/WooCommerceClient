using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WooCommerceClient.Models.Visma
{
    public enum SyncType { Products = 1, Stock = 2, Campaigns = 3, Orders = 4, Categories = 5 };

    [Serializable]
    public class Sync
    {
        public DateTime LastestSync { get; set; } = DateTime.MinValue;
    }
}
