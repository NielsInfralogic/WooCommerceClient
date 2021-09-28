using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using WooCommerceClient;
using WooCommerceClient.Models;
using WooCommerceClient.Services;
using WooCommerceClient.Services.SpecificHandling;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {
        //[TestMethod]
        public void TestAttributes()
        {
            bool res1 = Attributes.SyncAttributes().Result;
            Assert.IsTrue(res1);
        }

        [TestMethod]
        public void TestOrderLine()
        {
            var json = "{\"id\":67208,\"parent_id\":0,\"number\":\"67208\",\"order_key\":\"wc_order_scK16mGGLQGmA\",\"created_via\":\"checkout\",\"version\":\"5.1.0\",\"status\":\"processing\",\"currency\":\"EUR\",\"date_created\":\"2021 - 04 - 09T21: 24:57\",\"date_created_gmt\":\"2021 - 04 - 09T19: 24:57\",\"date_modified\":\"2021 - 04 - 09T21: 24:57\",\"date_modified_gmt\":\"2021 - 04 - 09T19: 24:57\",\"discount_total\":0.00,\"discount_tax\":0.00,\"shipping_total\":0.00,\"shipping_tax\":0.00,\"cart_tax\":15.08,\"total\":\"75.40\",\"total_tax\":\"15.08\",\"prices_include_tax\":true,\"customer_id\":1,\"customer_ip_address\":\"2.107.101.243\",\"customer_user_agent\":\"Mozilla / 5.0(Macintosh; Intel Mac OS X 10_15_7) AppleWebKit / 537.36(KHTML, like Gecko) Chrome / 89.0.4389.114 Safari / 537.36\",\"customer_note\":\"\",\"billing\":{\"first_name\":\"Mats\",\"last_name\":\"Kruger\",\"company\":\"\",\"address_1\":\"Moltkesvej 59 st tv\",\"address_2\":\"\",\"city\":\"Frederiksberg\",\"state\":\"\",\"postcode\":\"2000\",\"country\":\"DK\",\"email\":\"support @shift-it.dk\",\"phone\":\"23370102\"},\"shipping\":{\"first_name\":\"Mats\",\"last_name\":\"Kruger\",\"company\":\"\",\"address_1\":\"Moltkesvej 59 st tv\",\"address_2\":\"\",\"city\":\"Frederiksberg\",\"state\":\"\",\"postcode\":\"2000\",\"country\":\"DK\"},\"payment_method\":\"cod\",\"payment_method_title\":\"Efterkrav\",\"transaction_id\":\"\",\"cart_hash\":\"ab7cfc51ca2397a9a578d1158861393c\",\"meta_data\":[{\"id\":1905379,\"key\":\"is_vat_exempt\",\"value\":\"no\"},{\"id\":1905380,\"key\":\"_order_total_base_currency\",\"value\":\"580\"},{\"id\":1905381,\"key\":\"_cart_discount_base_currency\",\"value\":\"0\"},{\"id\":1905382,\"key\":\"_order_shipping_base_currency\",\"value\":\"0\"},{\"id\":1905383,\"key\":\"_order_tax_base_currency\",\"value\":\"116\"},{\"id\":1905384,\"key\":\"_order_shipping_tax_base_currency\",\"value\":\"0\"},{\"id\":1905385,\"key\":\"_cart_discount_tax_base_currency\",\"value\":\"0\"},{\"id\":1905386,\"key\":\"_base_currency_exchange_rate\",\"value\":\"7.692308\"},{\"id\":1905389,\"key\":\"wpml_language\",\"value\":\"da\"},{\"id\":1905394,\"key\":\"_new_order_email_sent\",\"value\":\"true\"}],\"line_items\":[{\"id\":10682,\"name\":\"Vino Bianco Vignarola\",\"product_id\":0,\"variation_id\":0,\"quantity\":1.0,\"tax_class\":\"\",\"subtotal\":\"18.72\",\"subtotal_tax\":\"4.68\",\"total\":\"18.72\",\"total_tax\":\"4.68\",\"taxes\":[{\"id\":1,\"total\":\"4.68\",\"subtotal\":\"4.68\"}],\"meta_data\":[{\"id\":121429,\"key\":\"_line_subtotal_base_currency\",\"value\":\"144\"},{\"id\":121431,\"key\":\"_line_subtotal_tax_base_currency\",\"value\":\"36\"},{\"id\":121433,\"key\":\"_line_total_base_currency\",\"value\":\"144\"},{\"id\":121435,\"key\":\"_line_tax_base_currency\",\"value\":\"36\"},{\"id\":121465,\"key\":\"_reduced_stock\",\"value\":\"1\"}],\"price\":18.72},{\"id\":10683,\"name\":\"Vino Bianco G05\",\"product_id\":0,\"variation_id\":0,\"quantity\":2.0,\"tax_class\":\"\",\"subtotal\":\"41.60\",\"subtotal_tax\":\"10.40\",\"total\":\"41.60\",\"total_tax\":\"10.40\",\"taxes\":[{\"id\":1,\"total\":\"10.4\",\"subtotal\":\"10.4\"}],\"meta_data\":[{\"id\":121442,\"key\":\"_line_subtotal_base_currency\",\"value\":\"320\"},{\"id\":121444,\"key\":\"_line_subtotal_tax_base_currency\",\"value\":\"80\"},{\"id\":121446,\"key\":\"_line_total_base_currency\",\"value\":\"320\"},{\"id\":121448,\"key\":\"_line_tax_base_currency\",\"value\":\"80\"},{\"id\":121466,\"key\":\"_reduced_stock\",\"value\":\"2\"}],\"price\":20.8}],\"tax_lines\":[{\"id\":\"10685\",\"rate_code\":\"DK - MOMS - 1\",\"rate_id\":\"1\",\"label\":\"Moms\",\"compound\":false,\"tax_total\":15.08,\"shipping_tax_total\":0.00,\"meta_data\":[{\"id\":121460,\"key\":\"tax_amount_base_currency\",\"value\":\"116\"},{\"id\":121462,\"key\":\"shipping_tax_amount_base_currency\",\"value\":\"0\"}]}],\"shipping_lines\":[{\"id\":10684,\"method_title\":\"Afhentning på lager\",\"method_id\":\"free_shipping\",\"instance_id\":\"7\",\"total\":\"0.00\",\"total_tax\":\"0.00\",\"taxes\":[],\"meta_data\":[{\"id\":121456,\"key\":\"Varer\",\"value\":\"Vino Bianco Vignarola & times; 1, Vino Bianco G05 & times; 2\"}]}],\"fee_lines\":[],\"coupon_lines\":[],\"refunds\":[]}";
            var order = JsonConvert.DeserializeObject<Order>(json);
            Assert.IsNotNull(order);
            //Orders.CheckOrderLineProduct(new DBaccess(), order);

            //test Lieu-Dit
            if (!VBSConvert.OrderConverts.Any(x => x is Lieu_Dit))
                VBSConvert.OrderConverts.Add(new Lieu_Dit());
            VBSconnection vbs = new VBSconnection(Utils.ReadConfigString("VBSsiteID", "standard"),
                                             Utils.ReadConfigString("VBSuser", "system"),
                                             Utils.ReadConfigString("VBSpassword", ""), // "Fredensvej17"),
                                             Utils.ReadConfigInt32("VBSCompanyNo", 9999), // 4182),
                                             Utils.ReadConfigString("Vismauser", "system"));
            vbs.ConvertToVBSOrder(order, 0);
        }

        //[TestMethod]
        public void TestAttributeTerms()
        {
            var idAttribute = 167;
            List<ProductAttributeTerm> wooCommerceAttributeTerms = WooCommerceHelpers.GetWooCommerceAttributeTerms(idAttribute).Result;
            Assert.IsNotNull(wooCommerceAttributeTerms);
            Utils.WriteLog($"Existing attribute terms {wooCommerceAttributeTerms.Count}");
        }
    }
}
