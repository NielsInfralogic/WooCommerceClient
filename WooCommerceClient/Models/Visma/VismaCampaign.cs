using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WooCommerceClient.Models.Visma
{
    public class CampaignCountOrMoreDiscountList
    {
        public List<CampaignCountOrMoreDiscount> campaigns;

        public CampaignCountOrMoreDiscountList()
        {
            campaigns = new List<CampaignCountOrMoreDiscount>();
        }
    }

    public class CampaignCountOrMoreDiscountPercent
    {
        public string id { get; set; } = "";
        public string type { get; } = "percentage_discount-count_or_more-single_product";
        public string product_id { get; set; } = "";
        public decimal percentage { get; set; } = 0.0M;
        public string name { get; set; } = "";
        public string display_name { get; set; } = "";
        public int count { get; set; } = 1;
        public int priority { get; set; } = 50;
    }

    public class CampaignCountOrMoreDiscount
    {
        public string id { get; set; } = "";
        public string type { get; } = "new_price_discount-count_or_more-single_product";
        public string product_id { get; set; } = "";
        public decimal new_price_per_item { get; set; } = 0.0M;
        public string name { get; set; } = "";
        public string display_name { get; set; } = "";
        public int count { get; set; } = 1;
        public int priority { get; set; } = 50;
    }

    public class CampaignPercentDiscountTag
    {
        public string id { get; set; } = "";
        public string type { get; } = "percentage_discount-tag";
        public string tag { get; set; } = "";
        public decimal percentage { get; set; } = 0.0M;
        public string name { get; set; } = "";
        public string display_name { get; set; } = "";
        public int priority { get; set; } = 50;
    }

    public class CampaignPriceDiscountProduct
    {
        public string id { get; set; } = "";
        public string type { get; } = "new_price_discount-single_product";
        public string product_id { get; set; } = "";
        public decimal new_price_per_item { get; set; } = 0.0M;
        public string name { get; set; } = "";
        public string display_name { get; set; } = "";
        public int priority { get; set; } = 50;
    }

    public class VismaDiscount
    {
        public decimal Discount { get; set; } = 0.0M;
        public int MinNo { get; set; } = 0;
    }

}
