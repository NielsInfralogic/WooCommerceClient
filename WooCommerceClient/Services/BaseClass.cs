using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WooCommerceClient.Services.SpecificHandling;

namespace WooCommerceClient.Services
{
    public class BaseClass
    {
        internal static readonly bool IsLieu_Dit = Utils.ReadConfigInt32("Lieu_Dit", 0) > 0;
        internal static readonly string TestProductNo = Utils.ReadConfigString("TestProductNo", null);
        internal static readonly string SyncAttributeOnly = Utils.ReadConfigString("SyncAttributeOnly", "");
     //   internal static readonly string SyncAttributeTypeOnly = Utils.ReadConfigString("SyncAttributeTypeOnly", null);
    }
}
