using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WooCommerceClient.Models;
using WooCommerceClient.Models.Visma;

namespace WooCommerceClient.Services
{
    class Products
    {
        internal static string DeleteSync(DBaccess db, Sync sync)
        {
            string errmsg = "";
            List<string> prodNoList = new List<string>();
 
            db.GetProductsToDelete(ref prodNoList,
                Utils.ReadConfigInt32("DeleteDisabledProducts", 0) > 1 ? DateTime.MinValue : sync.LastestSync, out errmsg);

            Utils.WriteLog("Products to delete: " + prodNoList.Count());
            if (prodNoList.Count > 0)
            {
                List<Product> wooProducts = GetOnlineProducts();
                // force: false will only update Val1=10 if Val1 was 9 (delete request)
                bool _ = SyncDeleteProducts(prodNoList, 44, wooProducts, false).Result;
                _ = SyncDeleteProducts(prodNoList, 45, wooProducts,false).Result;

            }

            return errmsg;
        }

        private static async Task<bool> SyncDeleteProducts(List<string> prodNoListToDelete, int langNo, List<Product> wooProducts, bool force)
        {
            if (wooProducts == null)
                return false;

            string lang = Utils.LangNoToString(langNo);
            DBaccess db = new DBaccess();
            try
            {
                var wcApi = new WC_API();

                foreach (string sku in prodNoListToDelete)
                {
                    Product wooProduct = wooProducts.FirstOrDefault(p => p.sku.Replace("-en", "") == sku.Replace("-en", "") && (p.lang == lang));
                    if (wooProduct != null)
                    {
                        try
                        {
                            Utils.WriteLog($"Deleting product {sku} with woocommerce  id:{wooProduct.id.Value}..");
                            var ok = await wcApi.Delete<Product>(wooProduct.id.Value, true);
                            if (ok == null)
                                Utils.WriteLog($"Error:  wc.Product.Delete returned null");
                            else
                            {
                                if (db.UpdateDeletedProduct(sku, force, out string errmsg) == false)
                                    Utils.WriteLog($"Error:  db.UpdateDeletedProduct() - {errmsg}");
                                wooProducts.Remove(wooProduct);
                            }
                        }
                        catch (Exception ex)
                        {
                            Utils.WriteLog($"Exception  wc.Attribute.Terms.Delete() - {ex.Message}");
                            Utils.WriteLog($"{ex.StackTrace}");
                            continue;
                        }
                    }
                    else
                    {
                        // not present in shop - mark as deleted..

                        if (db.UpdateDeletedProduct(sku, force, out string errmsg) == false)
                            Utils.WriteLog($"Error:  db.UpdateDeletedProduct() - {errmsg}");
                    }
                }
            }
            catch (Exception ex)
            {
                Utils.WriteLog($"Error SyncDeleteProducts()", ex);
                return false;
            }

            Utils.WriteLog($"SyncDeleteProducts done.");
            return true;
        }

        private static List<Product> GetOnlineProducts()
        {
            List<Product> wooProducts = null;
            try
            {
                wooProducts = WooCommerceHelpers.GetWooCommerceProducts().Result;
            }
            catch (Exception ex)
            {
                Utils.WriteLog($"Error: WooCommerceHelpers.GetWooCommerceProducts()", ex);
            }
            if (wooProducts == null)
            {
                Utils.WriteLog($"Error: WooCommerceHelpers.GetWooCommerceProducts(2) - GetWooCommerceProducts returned null");
            }

            return wooProducts;
        }

        public static List<int> GetUnrelatedEnglishProducts()
        {
            List<Product> wooProducts = null;
            try
            {
                wooProducts = WooCommerceHelpers.GetWooCommerceProducts().Result;
            }
            catch (Exception ex)
            {
                Utils.WriteLog($"Error: WooCommerceHelpers.GetWooCommerceProducts()", ex);
            }
            if (wooProducts == null)
            {
                Utils.WriteLog($"Error: WooCommerceHelpers.GetWooCommerceProducts(2) - GetWooCommerceProducts returned null");
            }
            List<int> wooProductsUnrelated = new List<int>();
            foreach(Product product in wooProducts)
            {
                if (product.lang == "en")
                    if (product.translations != null)
                        if (product.translations.da.HasValue == false || product.translations.da.Value == 0)
                          wooProductsUnrelated.Add(product.id.Value);
            }

            return wooProductsUnrelated;
        }


        internal static string Sync(DBaccess db, Sync sync)
        {
            string errmsg = "";
            List<ProductAttribute> wooCommerceAttributes = WooCommerceHelpers.GetWooCommerceAttributes().Result;
            List<ProductTag> wooCommerceTags = WooCommerceHelpers.GetWooCommerceTags().Result;
            List<ProductCategory> wooCommerceCategories = WooCommerceHelpers.GetWooCommerceCategories().Result;

            List<Product> products = new List<Product>();
            if (db.GetProducts(Utils.ReadConfigInt32("SyncProducts", 0) > 0 ? SyncType.Products : SyncType.Stock,
                ref products, sync.LastestSync, wooCommerceAttributes, wooCommerceTags, wooCommerceCategories, 45, out errmsg) == false)
            {
                Utils.WriteLog("ERROR db.GetProducts(da) - " + errmsg);
            }
            if (Utils.ReadConfigInt32("DoTranslation", 0) > 0)
            {
                List<Product> products_en = new List<Product>();
                if (db.GetProducts(Utils.ReadConfigInt32("SyncProducts", 0) > 0 ? SyncType.Products : SyncType.Stock, ref products_en, sync.LastestSync,
                                                        wooCommerceAttributes, wooCommerceTags, wooCommerceCategories, 44, out errmsg) == false)
                {
                    Utils.WriteLog("ERROR db.GetProducts(en) - " + errmsg);
                }


                // Ensure we have an english version of the product
                foreach (Product product in products)
                {
                    Product product_en = products_en.FirstOrDefault(p => p.sku == product.sku);
                    if (product_en == null)
                    {
                        product_en = new Product();
                        product_en.Clone(product);
                        products_en.Add(product_en);
                    }
                }

                // Copy to main array
                foreach (Product product_en in products_en)
                    products.Add(product_en);
            }

            bool ret = true;
            if (products.Count > 0)
            {
                ret = SyncProducts(products).Result;

            }

            if (ret)
            {
                sync.LastestSync = DateTime.Now.AddMinutes(-10);
                if (Utils.ReadConfigInt32("SyncProducts", 0) > 0)
                    Utils.WriteSyncTime(sync, Utils.ReadConfigString("ProductSyncFile", ""));
                else
                    Utils.WriteSyncTime(sync, Utils.ReadConfigString("StockSyncFile", ""));
            }

            // Remove unused producers
            /*  List<ProductAttributeTerm> attributeTermsProducers = new List<ProductAttributeTerm>();
              if (db.GetAttributeTermsForProducer(ref attributeTermsProducers, false, out errmsg) == false)
                  Utils.WriteLog("ERROR db.GetAttributeTermsForProducer() - " + errmsg);
              ret = DeleteUnusedProducers(attributeTermsProducers).Result;*/

            return errmsg;
        }

        internal static string DeleteStockSync(DBaccess db)
        {
            string errmsg = "";
            List<string> prodNoListZeroStock = new List<string>();
            if (db.GetProductsZeroStock(ref prodNoListZeroStock, out errmsg) == false)
            {
                Utils.WriteLog($"ERROR: db.GetProductsZeroStock() - {errmsg}");
                return errmsg;
            }

            Utils.WriteLog("Products stock=0 to delete: " + prodNoListZeroStock.Count());
            if (prodNoListZeroStock.Count > 0)
            {
                List<Product> wooProducts = GetOnlineProducts();
                // force: tru will update  Val1=10 regardless of previous state..
                bool _ = SyncDeleteProducts(prodNoListZeroStock, 44, wooProducts, true).Result;
                _ = SyncDeleteProducts(prodNoListZeroStock, 45, wooProducts, true).Result;
            }

            return errmsg;
        }

        private static async Task<bool> SyncProducts(List<Product> products)
        {
            try
            {
                bool hasNewProduct = false;
                DBaccess db = new DBaccess();

                //MyRestAPI rest = new MyRestAPI(Utils.ReadConfigString("WooCommerceUrl", ""),
                //         Utils.ReadConfigString("WooCommerceKey", ""),
                //         Utils.ReadConfigString("WooCommerceSecret", ""));
                //WCObject wc = new WCObject(rest);
                var wcApi = new WC_API(debug: true);

                List<ProductTag> wooTags = await WooCommerceHelpers.GetWooCommerceTags();
                List<ProductCategory> wooCategoires = await WooCommerceHelpers.GetWooCommerceCategories();
                List<ProductAttribute> wooAttributes = await WooCommerceHelpers.GetWooCommerceAttributes();
                // 

                List<Product> wooProducts = GetOnlineProducts();

                if (wooProducts == null)
                    return false;

                Utils.WriteLog($"{wooProducts.Count} existing products in WooCommerce..");

                Utils.WriteLog($"Add/update products..({products.Count})");

                foreach (Product product in products)
                {
                    if (Utils.ReadConfigInt32("SyncEnglishOnly", 0) > 0 && product.lang == "da")
                        continue;

                    Utils.WriteLog($"Testing product {product.sku} {product.lang}..");
                    int langNo = product.lang == "en" ? 44 : 45;

                    product.manage_stock = true;

                    if (Utils.ReadConfigInt32("ExcludeZeroStockProducts", 0) > 0 && product.stock_quantity == 0)
                    {
                        Utils.WriteLog($"Skipping product {product.sku} - stock=0");
                        continue;
                    }

                    GetWebShopInformations(wooTags, wooCategoires, wooAttributes, product);

                    Product existingProduct = wooProducts.FirstOrDefault(
                        p => p.sku.Replace("-en", "") == product.sku.Replace("-en", "") && p.lang == product.lang);
                    Product newUpdatedProduct = null;
                    if (existingProduct == null)
                    {
                        Utils.WriteLog($"Adding product {product.sku}..");
                        try
                        {
                            // 20210528 - leave slug undefined
                            hasNewProduct = true;

                            product.slug = null;

                            if (product.lang == "en" && product.sku.IndexOf("-en") == -1)
                                product.sku += "-en";

                            Utils.WriteLog($"wcApi.Add({JsonConvert.SerializeObject(product)})");
                            newUpdatedProduct = await wcApi.Add(product, product.lang);
                            product.id = newUpdatedProduct.id;
                        }
                        catch (Exception ex)
                        {
                            Utils.WriteLog($"Error : wc.Product.Add() - {ex.Message}");
                            continue;
                        }
                        if (newUpdatedProduct == null)
                        {
                            Utils.WriteLog("Error : wc.Product.Add() returned null");
                            continue;
                        }

                    }
                    else
                    {
                        try
                        {
                            Utils.WriteLog($"Updating product {product.sku}");

                            product.id = existingProduct.id;
                            //product.slug = existingProduct.slug;  //20210923 send no more slug to webshop
                            if (product.slug != null)
                                product.slug = null;
                            DisableSlug(product.categories);
                            DisableSlug(product.tags);

                            if (product.lang == "en" && product.sku.IndexOf("-en") == -1)
                                product.sku += "-en";


                            product.upsell_ids = new List<int>();
                            // Only set upsell on danish products!!
                            if (product.lang == "da" && product.vismaRelatedProduct != null &&
                                product.vismaRelatedProduct.Count > 0)
                            {
                                foreach (string r in product.vismaRelatedProduct)
                                {

                                    //  if (product.lang == "da")
                                    //  {
                                    Product product_da = wooProducts.FirstOrDefault(p => p.sku == r && p.lang == "da");
                                    if (product_da != null && product_da.id.HasValue)
                                    {
                                        //20210923 send no more slug to webshop
                                        if (product_da.slug != null)
                                            product_da.slug = null;

                                        DisableSlug(product_da.categories);
                                        DisableSlug(product_da.tags);

                                       // Utils.WriteLog($"wcApi.Add(2)({JsonConvert.SerializeObject(product)})");
                                        (product.upsell_ids as List<int>).Add(product_da.id.Value);
                                    }
                                    //   }
                                    //   else if (product.lang == "en")
                                    //   {
                                    //      Product product_en = wooProducts.FirstOrDefault(p => p.sku == r + "-en" && p.lang == "en");
                                    //      if (product_en != null)
                                    //          if (product_en.id.HasValue)
                                    //                product.upsell_ids.Add(product_en.id.Value);
                                    //   }
                                }
                            }

                            if (product.lang == "en")
                            {
                                Product product_da = wooProducts.FirstOrDefault(p => p.sku == product.sku.Replace("-en", "") && p.lang == "da");
                                if (product_da != null)
                                    product.translation_of = product_da.id.ToString();
                            }

                           // if ((product.lang == "en") ||
                          //      (product.lang == "da" && product.upsell_ids != null && (product.upsell_ids as List<int>).Count > 0))
                          //  {
                               // Utils.WriteLog($"wcApi.Update({JsonConvert.SerializeObject(product)})");
                                newUpdatedProduct = await wcApi.Update(product, product.lang);
                          //  }
                        }
                        catch (Exception ex)
                        {
                            Utils.WriteLog($"Error : wc.Product.Update() - {ex.Message}");
                            System.Threading.Thread.Sleep(2000);
                            try
                            {
                               // Utils.WriteLog($"wcApi.Update({JsonConvert.SerializeObject(product)})");
                                newUpdatedProduct = await wcApi.Update(product);
                            }
                            catch (Exception ex2)
                            {
                                Utils.WriteLog($"Error : wc.Product.Update() - {ex2.Message}");
                            }

                            continue;
                        }

                        if (newUpdatedProduct == null)
                        {
                            if (/*product.lang == "da" && */(product.upsell_ids as List<int>).Count == 0)
                                Utils.WriteLog($"product {product.sku} is not updated.");
                            else
                                Utils.WriteLog("Error : wc.Product.Update() returned null");
                            continue;
                        }
                    }



                }

             
                if (hasNewProduct)
                {

                    Utils.WriteLog($"Adjusting related product id's...");
                    // resolve related produc ids
                    wooProducts = await WooCommerceHelpers.GetWooCommerceProducts();

                    foreach (Product product in products)
                    {
                        if (Utils.ReadConfigInt32("SyncEnglishOnly", 0) > 0 && product.lang == "da")
                            continue;

                        if (product.lang == "en" && product.sku.IndexOf("-en") == -1)
                            product.sku += "-en";

                        product.upsell_ids = new List<int>();
                        if (product.lang == "da")
                        {
                            foreach (string r in product.vismaRelatedProduct)
                            {
                                Product product_da = wooProducts.FirstOrDefault(p => p.sku.Replace("-en", "") == r && p.lang == "da");
                                if (product_da != null && product_da.id.HasValue)
                                {
                                    (product.upsell_ids as List<int>).Add(product_da.id.Value);
                                }
                                /*                            if (product.lang == "da")
                                                            {
                                                                Product product_da = wooProducts.FirstOrDefault(p => p.sku == r && p.lang == "da");
                                                                if (product_da != null)
                                                                    if (product_da.id.HasValue)
                                                                        product.upsell_ids.Add(product_da.id.Value);
                                                            }
                                                            else if (product.lang == "en")
                                                            {
                                                                Product product_en = wooProducts.FirstOrDefault(p => p.sku == r + "-en" && p.lang == "en");
                                                                if (product_en != null)
                                                                    if (product_en.id.HasValue)
                                                                        product.upsell_ids.Add(product_en.id.Value);
                                                            }*/
                            }
                        }
                        Utils.WriteLog($"Updating product with adjusted upsell_ids {product.sku}..");


                        if (product.lang == "en")
                        {
                            Product product_da = wooProducts.FirstOrDefault(p => p.sku == product.sku.Replace("-en", "") && p.lang == "da");
                            if (product_da != null)
                                product.translation_of = product_da.id.ToString();
                        }



                      //  Utils.WriteLog($"wcApi.Update(2)({JsonConvert.SerializeObject(product)})");
                        await wcApi.Update(product, product.lang);
                    }
                }
            }
            catch (Exception ex)
            {
                Utils.WriteLog($"Error SyncProducts() - {ex.Message}");
                Utils.WriteLog($"Stacktrace: {ex.StackTrace}");
                return false;

            }

            Utils.WriteLog($"SyncProducts done.");

            return true;
        }

        private static void DisableSlug(IEnumerable<Islug> items)
        {
            if (items != null &&
                items.Any(x => x.slug != null))
                foreach (var cat in items.Where(x => x.slug != null))
                    cat.slug = null;
        }

        private static void GetWebShopInformations(List<ProductTag> wooTags, List<ProductCategory> wooCategoires, List<ProductAttribute> wooAttributes, Product product)
        {
            // Resolve woo ids for tags
            foreach (ProductTagLine tagLineInProduct in product.tags)
            {
                ProductTag wooTag = wooTags.FirstOrDefault(p => (p.slug == tagLineInProduct.slug || p.name == tagLineInProduct.name) && p.lang == product.lang);
                if (wooTag != null)
                    tagLineInProduct.id = wooTag.id;
            }

            // Resolve woo ids for catagoires
            if (product.categories != null)
            {
                foreach (ProductCategoryLine categoryLineInProduct in product.categories)
                {
                    ProductCategory wooCategory = wooCategoires.FirstOrDefault(p => (p.slug == categoryLineInProduct.slug || p.name == categoryLineInProduct.name) && p.lang == product.lang);
                    if (wooCategory != null)
                        categoryLineInProduct.id = wooCategory.id;
                }
            }

            // Resolve woo ids for attributes (no translation)
            foreach (ProductAttributeLine attributeLineInProduct in product.attributes)
            {
                ProductAttribute wooAttribute = wooAttributes.FirstOrDefault(p => p.name == attributeLineInProduct.name);
                if (wooAttribute != null)
                {
                    attributeLineInProduct.id = wooAttribute.id;
                }

            }
        }

        internal static bool DeleteAllProducts()
        {
            MyRestAPI rest = new MyRestAPI(Utils.ReadConfigString("WooCommerceUrl", ""),
                      Utils.ReadConfigString("WooCommerceKey", ""),
                      Utils.ReadConfigString("WooCommerceSecret", ""));
            WCObject wc = new WCObject(rest);

            List<Product> wooProducts = GetOnlineProducts();
            if (wooProducts == null)
            {
                return false;
            }

            List<string> productsToDelete = new List<string>();
            foreach (Product wooProduct in wooProducts)
            {

               /* if (wooProduct.lang == "en")
                {
                    try
                    {
                        if (wooProduct.upsell_ids != null)
                        {

                            List<int> xx = wooProduct.upsell_ids as List<int>;
                            int x = xx.Count;
                        }
                    }
                    catch (Exception ex)
                    {
                        Utils.WriteLog($"{ex.Message}");
                        Utils.WriteLog($"Funny id: {wooProduct.id}");
                    }
                }*/
                try
                {
                    var ok = wc.Product.Delete(wooProduct.id.Value, true).Result;
                    if (ok == null)
                    {
                        Utils.WriteLog($"Error:  wc.Product.Delete returned null");
                        return false;
                    }
                    else
                        Utils.WriteLog($"Product id {wooProduct.id.Value} deleted.");
                }
                catch (Exception ex)
                {
                    Utils.WriteLog($"Exception  wc.Attribute.Terms.Delete() - {ex.Message}");
                    return false;
                }
            }

            return true;
        }

        public static bool DeleteProduct(int id)
        {
            MyRestAPI rest = new MyRestAPI(Utils.ReadConfigString("WooCommerceUrl", ""),
                      Utils.ReadConfigString("WooCommerceKey", ""),
                      Utils.ReadConfigString("WooCommerceSecret", ""));
           

            try
            {
                WCObject wc = new WCObject(rest);
                var ok = wc.Product.Delete(id, true).Result;
                if (ok == null)
                {
                    Utils.WriteLog($"Error:  wc.Product.Delete returned null");
                    return false;
                }
                else
                    Utils.WriteLog($"Product id {id} deleted.");
            }
            catch (Exception ex)
            {
                Utils.WriteLog($"Exception  wc.Attribute.Terms.Delete() - {ex.Message}");
                return false;
            }
         

            return true;
        }

    }
}
