using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Visma.BusinessModel;
using WooCommerceClient.Models;
using WooCommerceClient.Models.Visma;
using WooCommerceClient.Services;

namespace WooCommerceClient
{
    class Program
    {
        static void Main(string[] args)
        {
            CultureInfo ci = new CultureInfo("en-US");
            Thread.CurrentThread.CurrentCulture = ci;
            Thread.CurrentThread.CurrentUICulture = ci;

            DBaccess db = new DBaccess();

            Utils.WriteLog("Starting..");
            if (db.TestConnection(out string errmsg) == false)
            {
                Utils.WriteLog("ERROR db.TestConnection() - " + errmsg);
                return;
            }
            Utils.WriteLog("Database connection OK.");

            if (WooCommerceHelpers.TestWooCommerceConnection().Result == false)
            {
                Utils.WriteLog("ERROR TestWooCommerceConnection");
                return;
            }
            Utils.WriteLog("WooCommerce connection OK.");

            //bool res1 =  DeleteAllWooCommerceAttributes().Result;

            // Check if any products to process at all..
            List<string> productsTest = new List<string>();
            Sync sync1 = Utils.ReadSyncTime(Utils.ReadConfigString("ProductSyncFile", ""));
            if (db.GetProductsToProcess(ref productsTest, sync1.LastestSync, out errmsg) == false)
            {
                Utils.WriteLog($"ERROR: db.GetProductsToProcess() - {errmsg}");
                return;
            }

            //////////////////////////////
            /// Special run - DeleteAllProducts
            //////////////////////////////
            if (Utils.ReadConfigInt32("DeleteAllProducts", 0) > 0)
            {
                Products.DeleteAllProducts();
                if (Utils.ReadConfigInt32("DeleteAllProducts", 0) > 1)
                {
                    Attributes.DeleteAllAttributes();
                    Tags.DeleteAllTags();
                }

                return;
            }
            //////////////////////////////
            /// Special run - Delete any english products without link to danish mother product.
            //////////////////////////////
            if (Utils.ReadConfigInt32("DeleteAllUnrelatedProducts", 0) > 0)
            {

                List<int> unrelatedIDList = Products.GetUnrelatedEnglishProducts();
                foreach (int id in unrelatedIDList)
                {
                    Utils.WriteLog($"Found unrelated en-product {id}");
                    Products.DeleteProduct(id);
                }

                return;
            }


            
            //////////////////////////////
            /// Special run - Delete any english products without link to danish mother product.
            //////////////////////////////
            if (Utils.ReadConfigInt32("ListAllUnrelatedAttributes", 0) > 0)
            {

               // Attributes.ListUnrelatedEnglishAttributes();
               

                return;
            }



            //////////////////////////////
            /// DeleteZeroStockProducts
            //////////////////////////////
            if (Utils.ReadConfigInt32("DeleteZeroStockProducts", 0) > 0)
            {
                errmsg += Products.DeleteStockSync(db);
                return;
            }

            // skip tags,attributes if no products to update
            //if (productsTest.Count > 0 || (Utils.ReadConfigInt32("SyncProducts", 0) == 0 && Utils.ReadConfigInt32("SyncStocks", 0) == 0))
           // {
                //////////////////////////////
                /// SyncTags
                //////////////////////////////
                if (Utils.ReadConfigInt32("SyncTags", 0) > 0)
                {
                    Tags.Sync(db, out errmsg);
                }

                //////////////////////////////
                /// SyncAttributes (actually only terms are sync'ed
                //////////////////////////////
                if (Utils.ReadConfigInt32("SyncAttributes", 0) > 0)
                {
                    Sync sync = Utils.ReadSyncTime(Utils.ReadConfigString("AttributeSyncFile", ""));

                    errmsg += Attributes.SyncAttributTerms(db, sync.LastestSync);

                    sync.LastestSync = DateTime.Now;
                    Utils.WriteSyncTime(sync, Utils.ReadConfigString("AttributeSyncFile", ""));
                }

                //////////////////////////////
                /// SyncCategories
                //////////////////////////////
                if (Utils.ReadConfigInt32("SyncCategories", 0) > 0)
                {
                    Categories.Syns();
                }
            //}

            //////////////////////////////
            /// SyncProducts/SyncStocks
            //////////////////////////////
            if (Utils.ReadConfigInt32("SyncProducts", 0) > 0 || Utils.ReadConfigInt32("SyncStocks", 0) > 0)
            {
                Utils.WriteLog("Synchronizing products..");

                Sync sync;
                if (Utils.ReadConfigInt32("SyncProducts", 0) > 0)
                    sync = Utils.ReadSyncTime(Utils.ReadConfigString("ProductSyncFile", ""));
                else
                    sync = Utils.ReadSyncTime(Utils.ReadConfigString("StockSyncFile", ""));
                if (Utils.ReadConfigInt32("SyncAllProducts", 0) > 0)
                    sync.LastestSync = DateTime.MinValue;

                if (Utils.ReadConfigInt32("SyncProducts", 0) > 0 && Utils.ReadConfigInt32("DeleteDisabledProducts", 0) > 0)
                {
                    errmsg += Products.DeleteSync(db, sync);
                }

                errmsg += Products.Sync(db, sync);

            }

            //////////////////
            /// SyncOrders
            //////////////////
            if (Utils.ReadConfigInt32("SyncOrders", 0) > 0)
            {
                Orders.Syns();
            }
        }
    }
}
