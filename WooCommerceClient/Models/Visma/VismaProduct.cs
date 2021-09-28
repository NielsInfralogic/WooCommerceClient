using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WooCommerceClient.Models.Visma
{
    public class VismaBomItem
    {
        public string ProdNo { get; set; } = "";
        public int Qty { get; set; } = 1;
        public string Descr { get; set; } = "";
    }

    public class MetaDataEx
    {
        public string title { get; set; }
    }
    public class VismaProductDetail
    {
        public string ProdNo { get; set; } = "";
        public string ProductType { get; set; } = "";
        public string Year { get; set; } = "";
        public string Producer { get; set; } = "";
        public string Country { get; set; } = "";
        public string Region { get; set; } = "";

        public string County { get; set; } = "";
        
        public string Eco { get; set; } = "";
        public string Volume { get; set; } = "";
        public string Alcohol { get; set; } = "";

        public List<string> Grapes { get; set; }

        public bool HasBOM { get; set; } = false;

        public bool NoDiscount { get; set; } = false;

        public List<VismaBomItem> VismaBomItems;

        public VismaProductDetail()
        {
            Grapes = new List<string>();
            VismaBomItems = new List<VismaBomItem>();
        }
    }
}
    