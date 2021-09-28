using RestSharp;
using RestSharp.Authenticators;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WooCommerceClient.Models;
using WooCommerceClient.Services;

namespace WooCommerceClient
{
    public static class WooCommerceHelpers
    {
        public static async Task<List<ProductCategory>> GetWooCommerceCategories()
        {
            MyRestAPI rest = new MyRestAPI(Utils.ReadConfigString("WooCommerceUrl", ""),
                       Utils.ReadConfigString("WooCommerceKey", ""),
                       Utils.ReadConfigString("WooCommerceSecret", ""));
            WCObject wc = new WCObject(rest);

            //Get all products categoires.
            List<ProductCategory> categoires = new List<ProductCategory>();

            int per_page = 100;
            int offset = 0;

            while (true)
            {
                var categoriespage = await wc.Category.GetAll(new Dictionary<string, string>() { { "per_page", per_page.ToString() }, { "offset", offset.ToString() }, { "lang", "da" } });
                if (categoriespage == null || categoriespage?.Count == 0)
                    break;
                foreach (ProductCategory c in categoriespage)
                    categoires.Add(c);
                if (categoriespage.Count < per_page)
                    break;
                offset += per_page;
            }

            offset = 0;
            if (Utils.ReadConfigInt32("DoTranslation", 0) > 0)
            {
                while (true)
                {
                    var categoriespage = await wc.Category.GetAll(new Dictionary<string, string>() { { "per_page", per_page.ToString() }, { "offset", offset.ToString() }, { "lang", "en" } });
                    if (categoriespage == null || categoriespage?.Count == 0)
                        break;
                    foreach (ProductCategory c in categoriespage)
                        categoires.Add(c);
                    if (categoriespage.Count < per_page)
                        break;
                    offset += per_page;
                }
            }

            return categoires;
        }

        public static async Task<List<ProductTag>> GetWooCommerceTags()
        {
            MyRestAPI rest = new MyRestAPI(Utils.ReadConfigString("WooCommerceUrl", ""),
                       Utils.ReadConfigString("WooCommerceKey", ""),
                       Utils.ReadConfigString("WooCommerceSecret", ""));
            WCObject wc = new WCObject(rest);
          
            //Get all products categoires.
            List<ProductTag> tags = new List<ProductTag>();

            int per_page = 100;
            int offset = 0;

            while (true)
            {
                var tagspage = await wc.Tag.GetAll(new Dictionary<string, string>() { { "per_page", per_page.ToString() }, { "offset", offset.ToString() }, { "lang", "da" } });
                if (tagspage == null || tagspage?.Count == 0)
                    break;
                foreach (ProductTag t in tagspage)
                    tags.Add(t);
                offset += per_page;
            }

            offset = 0;
            if (Utils.ReadConfigInt32("DoTranslation", 0) > 0)
            {
                while (true)
                {
                    var tagspage = await wc.Tag.GetAll(new Dictionary<string, string>() { { "per_page", per_page.ToString() }, { "offset", offset.ToString() }, { "lang", "en" } });
                    if (tagspage == null || tagspage?.Count == 0)
                        break;
                    foreach (ProductTag t in tagspage)
                        tags.Add(t);
                    offset += per_page;
                }
            }


            return tags;
        }


        public static async Task<List<ProductAttribute>> GetWooCommerceAttributes()
        {
            MyRestAPI rest = new MyRestAPI(Utils.ReadConfigString("WooCommerceUrl", ""),
                       Utils.ReadConfigString("WooCommerceKey", ""),
                       Utils.ReadConfigString("WooCommerceSecret", ""));
            WCObject wc = new WCObject(rest);

            //Get all products attributes.
            List<ProductAttribute> attributes = new List<ProductAttribute>();

            /*    int per_page = 100;
            int offset = 0;

                    while (true)
                        {
                            var attributespage = await wc.Attribute.GetAll(new Dictionary<string, string>() { { "per_page", per_page.ToString() }, { "offset", offset.ToString() } });
                            if (attributespage == null || attributespage?.Count == 0)
                                break;
                            foreach (ProductAttribute a in attributespage)
                            {
                                Utils.WriteLog($"Attribute {a.name} {a.slug}");
                                attributes.Add(a);
                            }
                            if (attributespage.Count < per_page)
                                break;
                            offset += per_page;
                        }
                        */
            var attributespage = await wc.Attribute.GetAll();
            foreach (ProductAttribute a in attributespage)
            {
                Utils.WriteLog($"Attribute  {a.name} {a.slug}");
                attributes.Add(a);
            }

            return attributes;
        }


        public static async Task<bool> DeleteAllWooCommerceAttributes()
        {
            MyRestAPI rest = new MyRestAPI(Utils.ReadConfigString("WooCommerceUrl", ""),
                       Utils.ReadConfigString("WooCommerceKey", ""),
                       Utils.ReadConfigString("WooCommerceSecret", ""));
            WCObject wc = new WCObject(rest);

            //Get all products attributes.
            List<ProductAttribute> attributes = await GetWooCommerceAttributes();

            foreach (ProductAttribute a in attributes)
            {
                Utils.WriteLog($"Attribute {a.name} {a.slug}");
                await wc.Attribute.Delete(a.id.Value);
            }

            return true;
        }

        public static async Task<List<ProductAttributeTerm>> GetWooCommerceAttributeTerms(int attributeId)
        {
            MyRestAPI rest = new MyRestAPI(Utils.ReadConfigString("WooCommerceUrl", ""),
                       Utils.ReadConfigString("WooCommerceKey", ""),
                       Utils.ReadConfigString("WooCommerceSecret", ""));
            WCObject wc = new WCObject(rest);

            //Get all products attributes.
            List<ProductAttributeTerm> terms = new List<ProductAttributeTerm>();

            int per_page = 100;
            int offset = 0;

            while (true)
            {
                var attributespage = await wc.Attribute.Terms.GetAll(attributeId, new Dictionary<string, string>() { { "per_page", per_page.ToString() }, { "offset", offset.ToString() }, { "lang", "da" } });
                if (attributespage == null || attributespage?.Count == 0)
                    break;
                foreach (ProductAttributeTerm a in attributespage)
                {
                     Utils.WriteLog($"Attribute {a.name} {a.slug} id:{a.id}");
                    terms.Add(a);
                }
                if (attributespage.Count < per_page)
                    break;
                offset += per_page;
            }
            offset = 0;
            if (Utils.ReadConfigInt32("DoTranslation", 0) > 0)
            {
                while (true)
                {
                    var attributespage = await wc.Attribute.Terms.GetAll(attributeId, new Dictionary<string, string>() { { "per_page", per_page.ToString() }, { "offset", offset.ToString() }, { "lang", "en" } });
                    if (attributespage == null || attributespage?.Count == 0)
                        break;
                    foreach (ProductAttributeTerm a in attributespage)
                    {
                        //  Utils.WriteLog($"Attribute {a.name} {a.slug} id:{a.id}");
                        terms.Add(a);
                    }
                    if (attributespage.Count < per_page)
                        break;
                    offset += per_page;
                }
            }

            return terms;
        }

        public static async Task<List<Product>> GetWooCommerceProducts()
        {
            MyRestAPI rest = new MyRestAPI(Utils.ReadConfigString("WooCommerceUrl", ""),
                       Utils.ReadConfigString("WooCommerceKey", ""),
                       Utils.ReadConfigString("WooCommerceSecret", ""));
            WCObject wc = new WCObject(rest);

            //Get all products .

            int per_page = 100;
            int offset = 0;
            List<Product> products = new List<Product>(); ;

            while (true)
            {
                try
                {
                    Utils.WriteLog($"Product.GetAll - offset {offset}..");

                    var productspage = await wc.Product.GetAll(new Dictionary<string, string>() { { "per_page", per_page.ToString() }, { "offset", offset.ToString() }, { "lang", "da" } } );
                    if (productspage == null || productspage?.Count == 0)
                        break;
                    foreach (Product p in productspage)
                        products.Add(p);
                    offset += per_page;
                }
                catch(Exception ex)
                {
                    Utils.WriteLog($"Error in wc.Product.GetAll - {ex.Message}");                   
                    return null;
                }
            }
            offset = 0;
            if (Utils.ReadConfigInt32("DoTranslation", 0) > 0)
            {
                //per_page = 1;
                while (true)
                {
                    try
                    {
                        

                        var productspage = await wc.Product.GetAll(new Dictionary<string, string>() { { "per_page", per_page.ToString() }, { "offset", offset.ToString() }, { "lang", "en" } });
                        if (productspage == null || productspage?.Count == 0)
                            break;
                        foreach (Product p in productspage)
                        {
                            Utils.WriteLog($"Product.GetAll - offset {offset} {p.id}..");
                            products.Add(p);
                        }

                        offset += per_page;
                    }
                    catch (Exception ex)
                    {
                        Utils.WriteLog($"Error in wc.Product.GetAll - {ex.Message}");
                        offset += per_page;
                        // return null;
                    }
                }

                foreach (Product p in products)
                {
                    if (p.lang == "en")
                        p.sku = p.sku.Replace("-en", "");
                }
            }
            Utils.WriteLog($"Products in WooCommerce: {products.Count}.");
            return products;
        }

        public static async Task<List<Order>> GetWooCommerceOrders(DateTime lastSyncTime)
        {
            MyRestAPI rest = new MyRestAPI(Utils.ReadConfigString("WooCommerceUrl", ""),
                       Utils.ReadConfigString("WooCommerceKey", ""),
                       Utils.ReadConfigString("WooCommerceSecret", ""));
            WCObject wc = new WCObject(rest);



            //Get all orders .
            int per_page = 100;
            int offset = 0;
            List<Order> orders = new List<Order>();
            while (true)
            {
                Dictionary<string, string> p = new Dictionary<string, string>();
                p.Add("per_page", per_page.ToString());
                p.Add("offset", offset.ToString());
                if (lastSyncTime != DateTime.MinValue)
                {
                    //  p.Add("after", Utils.DateTimeToISO8601(lastSyncTime));
                    p.Add("after", Utils.GenerateTimeStampT(lastSyncTime));
                }

                Utils.WriteLog($"Date used in filter: {Utils.DateTimeToISO8601(lastSyncTime)}   {Utils.GenerateTimeStampT(lastSyncTime)}");
                //   var ordersspage = await wc.Order.GetAll(new Dictionary<string, string>() { { "per_page", per_page.ToString() }, { "offset", offset.ToString() } });
                var ordersspage = await wc.Order.GetAll(p);
                if (ordersspage == null || ordersspage?.Count == 0)
                    break;
                //foreach (Order o in ordersspage)
                //    orders.Add(o);
                orders.AddRange(ordersspage);

                if (ordersspage.Count < per_page)
                    break;  //last page reached

                offset += per_page;
            }

            foreach (Order o in orders)
            {
                Utils.WriteLog($"Orders {o.id} {o.number} {o.line_items.Count} lines.");
            }

            return orders;
        }

        public static async Task<Order> GetWooCommerceOrder(int id)
        {
            MyRestAPI rest = new MyRestAPI(Utils.ReadConfigString("WooCommerceUrl", ""),
                       Utils.ReadConfigString("WooCommerceKey", ""),
                       Utils.ReadConfigString("WooCommerceSecret", ""));
            WCObject wc = new WCObject(rest);

            Order order = await wc.Order.Get(id);                         

            return order;
        }

        public static async Task<bool> UpdateOrderStatus(int id, Order order)
        {
            MyRestAPI rest = new MyRestAPI(Utils.ReadConfigString("WooCommerceUrl", ""),
                       Utils.ReadConfigString("WooCommerceKey", ""),
                       Utils.ReadConfigString("WooCommerceSecret", ""));
            WCObject wc = new WCObject(rest);

            try
            {
                Order returnedOrder = await wc.Order.Update(id, order);
            }
            catch (Exception ex)
            {
                Utils.WriteLog($"Error : wc.Order.Update() - {ex.Message}");
                return false;
            }

            return true;

        }


        public static bool UpdateOrderStatusV1(string orderID, string status)
        {
            try
            {
                var client = new RestClient(Utils.ReadConfigString("WooCommerceUrl", ""))
                {
                    Authenticator = new HttpBasicAuthenticator(Utils.ReadConfigString("WooCommerceKey", ""), Utils.ReadConfigString("WooCommerceSecret", ""))
                };

                var request = new RestRequest("orders/" + orderID, Method.PUT);
                request.AddHeader("Accept", "application/json");
                request.Parameters.Clear();
                request.RequestFormat = DataFormat.Json;
                OrderStatus orderStatus = new OrderStatus() { status = status };
                request.AddJsonBody(orderStatus);

                var response = client.Execute(request);

                if (response.StatusCode == System.Net.HttpStatusCode.Accepted ||
                    response.StatusCode == System.Net.HttpStatusCode.OK ||
                    response.StatusCode == System.Net.HttpStatusCode.Created)
                {
                    Utils.WriteLog("Woocommerce order " + orderID.ToString() + " updated to status '" + status + "'");
                    return true;
                }
                else
                {

                    var content = response.Content; // raw content as string  
                    Utils.WriteLog("Error response: " + content);

                    return false;
                }
            }
            catch (Exception ex)
            {
                Utils.WriteLog("Error:" + ex.Message);

            }

            return false;

        }

        public static async Task<List<OrderNote>> GetWooCommerceOrderNotes(int id)
        {
            MyRestAPI rest = new MyRestAPI(Utils.ReadConfigString("WooCommerceUrl", ""),
                       Utils.ReadConfigString("WooCommerceKey", ""),
                       Utils.ReadConfigString("WooCommerceSecret", ""));
            WCObject wc = new WCObject(rest);

            //Get all orders .

            Utils.WriteLog($"Retrieving order notes..");
            int per_page = 100;
            int offset = 0;
            List<OrderNote> orderNotes = new List<OrderNote>();
            while (true)
            {
                var ordernotessspage = await wc.Order.Notes.GetAll(id, new Dictionary<string, string>() { { "per_page", per_page.ToString() }, { "offset", offset.ToString() } });
                if (ordernotessspage == null || ordernotessspage?.Count == 0)
                    break;
                foreach (OrderNote o in ordernotessspage)
                {
                    orderNotes.Add(o);
                }
                break; // on purpose!
                //offset += per_page;
            }

            foreach (OrderNote o in orderNotes)
            {
                Utils.WriteLog($"Order ID {id} - Note {o.note}");
            }
            Utils.WriteLog($"Got order notes..");

            return orderNotes;
        }

        public static async Task<List<OrderRefund>> GetWooCommerceOrderRefunds(int id)
        {
            MyRestAPI rest = new MyRestAPI(Utils.ReadConfigString("WooCommerceUrl", ""),
                       Utils.ReadConfigString("WooCommerceKey", ""),
                       Utils.ReadConfigString("WooCommerceSecret", ""));
            WCObject wc = new WCObject(rest);

            //Get all orders .

            int per_page = 100;
            int offset = 0;
            List<OrderRefund> orderRefunds = new List<OrderRefund>();
            while (true)
            {
                var orderRefundspage = await wc.Order.Refunds.GetAll(id, new Dictionary<string, string>() { { "per_page", per_page.ToString() }, { "offset", offset.ToString() } });
                if (orderRefundspage == null || orderRefundspage?.Count == 0)
                    break;
                foreach (OrderRefund o in orderRefundspage)
                {
                    orderRefunds.Add(o);
                }
                offset += per_page;
            }

            foreach (OrderRefund o in orderRefunds)
            {
                if (o.amount.HasValue)
                    Utils.WriteLog($"Order ID {id} - Refund {o.amount}");
            }

            return orderRefunds;
        }


        public static async Task<bool> TestWooCommerceConnection()
        {
            MyRestAPI rest = new MyRestAPI(Utils.ReadConfigString("WooCommerceUrl", ""),
                        Utils.ReadConfigString("WooCommerceKey", ""),
                        Utils.ReadConfigString("WooCommerceSecret", ""));
            WCObject wc = new WCObject(rest);
            List<Setting> settings = null;
            try
            {
                settings = await wc.Setting.GetAll();
            }
            catch(Exception  ex)
            {
                Utils.WriteLog($"Unable to connect to webshop - {ex.Message}");
            }

            return settings != null;
        }
    }
}
