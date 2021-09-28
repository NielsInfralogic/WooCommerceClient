using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using WooCommerceClient.Models;
using WooCommerceClient.Models.Visma;
using WooCommerceClient.Services;
using WooCommerceClient.Services.SpecificHandling;

namespace WooCommerceClient.Services
{
    public class Orders : BaseClass
    {
        private static bool SyncOrders(List<Order> orders)
        {
            DBaccess db = new DBaccess();
            DateTime changeTime = DateTime.Now;
            foreach (Order order in orders)
            {
                if (order.id.HasValue == false)
                    continue;
                if (order.line_items == null)
                    continue;
                // Only look for status pending?
                //pending, processing, on - hold, completed, cancelled, refunded, failed and trash.Default is pending.

                if (order.status != "processing")
                    continue;
                // See if order already registred
                int vismaOrdNo = db.ExistingVismaWebOrder(order.id.Value.ToString(), ref changeTime, out string errmsg);

                if (vismaOrdNo < 0)
                {
                    Utils.WriteLog($"Error db.ExistingVismaWebOrder() - {errmsg}");
                    continue;
                }
                if (vismaOrdNo > 0)
                {
                    Utils.WriteLog($"Web-order {order.id.Value} already exists - visma order number is {vismaOrdNo}");
                    continue;
                }

                Utils.WriteLog($"NEW Order {order.id.Value}..");
                Utils.WriteLog($"{JsonConvert.SerializeObject(order)}");
                if (order.billing.email == "" && Utils.ReadConfigInt32("OrderMustHaveEmail", 1) > 0)
                {
                    Utils.WriteLog($"Order {order.id.Value} rejected - Email must be provided in order header");
                    continue;
                }

                List<OrderNote> orderNotes = WooCommerceHelpers.GetWooCommerceOrderNotes(order.id.Value).Result;
                try
                {
                    List<OrderRefund> orderRefunds = WooCommerceHelpers.GetWooCommerceOrderRefunds(order.id.Value).Result;
                }
                catch (Exception e)
                {
                    Utils.WriteLog("OrderRefunds", e);
                }


                //  bool hasUnknownProdNo = false;

                //  if (Utils.ReadConfigBoolean("CheckOrderProducts", true))
                //   {

                if (order.line_items != null && order.line_items.Count > 0)
                    errmsg += CheckOrderLineProduct(db, order);
                //}

                Models.Visma.VismaActor existingcustomer = new Models.Visma.VismaActor();
                if (db.GetCustomer(order.billing.email, 0, ref existingcustomer, out errmsg) == false)
                {
                    Utils.WriteLog($"Error db.GetCustomer() - {errmsg}");
                    continue;
                }


                if (existingcustomer.CustomerNo > 0)
                    Utils.WriteLog($"Found customer in visma (CustNo {existingcustomer.CustomerNo})");

                VBSconnection vbs = new VBSconnection(Utils.ReadConfigString("VBSsiteID", "standard"),
                                                 Utils.ReadConfigString("VBSuser", "system"),
                                                 Utils.ReadConfigString("VBSpassword", ""), // "Fredensvej17"),
                                                 Utils.ReadConfigInt32("VBSCompanyNo", 9999), // 4182),
                                                 Utils.ReadConfigString("Vismauser", "system"));

                int retries = 20;
                if (existingcustomer.CustomerNo == 0)
                {
                    Models.Visma.VismaActor actor = new Models.Visma.VismaActor
                    {
                        Name = (order.billing.first_name + " " + order.billing.last_name).Trim(),
                        AddressLine1 = order.billing.address_1 ?? "",
                        AddressLine2 = order.billing.address_2 ?? "",
                        AddressLine3 = "",
                        PostalArea = order.billing.city ?? "",
                        PostCode = order.billing.postcode ?? "",
                        Phone = order.billing.phone,
                        EmailAddress = order.billing.email ?? "",

                        DeliveryName = (order.shipping.first_name + " " + order.shipping.last_name).Trim(),
                        DeliveryAddressLine1 = order.shipping.address_1 ?? "",
                        DeliveryAddressLine2 = order.shipping.address_2 ?? "",
                        DeliveryAddressLine3 = "",
                        DeliveryPostalArea = order.shipping.city ?? "",
                        DeliveryPostCode = order.shipping.postcode ?? "",

                    };

                    actor.CurrencyNo = db.LookupCurrencyCode(order.currency ?? "DKK", out errmsg);

                    actor.CountryNo = db.LookupCountryCode(order.billing.country ?? "", out errmsg);
                    if (actor.CountryNo <= 0)
                        actor.CountryNo = Utils.ReadConfigInt32("DefaultCountryCode", 45);
                    actor.DeliveryCountryNo = db.LookupCountryCode(order.shipping.country ?? "", out errmsg);
                    if (actor.DeliveryCountryNo <= 0)
                        actor.DeliveryCountryNo = Utils.ReadConfigInt32("DefaultCountryCode", 45);


                    var languageMetaData = order.meta_data?.Count > 0
                        ? order.meta_data.FirstOrDefault(x => x.key == "wpml_language")
                        : null;
                    string language = languageMetaData != null ? (string)languageMetaData.value : "da";
                    actor.LanguageNo = language.ToLower() == "en" ? 44 : 45;

                    //int dummycustNoToUse = -1;
                    //dbOrders.GetNextFreeCustomerNo(ref dummycustNoToUse, out errmsg);
                    // Create customer
                    Utils.WriteLog("Calling AddVBSCustomer()..");
                    existingcustomer.CustomerNo = 0;
                    int ActNoCreated = 0;
                    int CustNoCreated = 0;
                    if (vbs.AddVBSCustomer(actor, -1, ref ActNoCreated, ref CustNoCreated, out errmsg) == false)
                        Utils.WriteLog("ERROR AddVBSCustomer() failed - " + errmsg);
                    else
                    {
                        Thread.Sleep(1000);
                        int vismaCustNumberTest = 0;
                        while (--retries > 0)
                        {
                            vismaCustNumberTest = db.CheckCustomer(order.billing.email, out errmsg);
                            if (vismaCustNumberTest > 0)
                                break;

                            if (vismaCustNumberTest < 0)
                                Utils.WriteLog("ERROR CheckCustomer() failed - " + errmsg);

                            Thread.Sleep(1000);
                        }
                        if (vismaCustNumberTest == 0)
                        {
                            Utils.WriteLog("ERROR: Unable to verify customer creation with email " + order.billing.email);
                        }
                        actor.CustomerNo = CustNoCreated;
                        actor.AssociateNo = ActNoCreated;

                        existingcustomer.CustomerNo = CustNoCreated;

                        // if (actor.CountryNo != 45)
                        //     db.UpdateCustomerCity(CustNoCreated, actor.DeliveryPostalArea, actor.DeliveryCountryNo, out errmsg);
                    }
                }

                Models.Visma.VismaOrder vorder = vbs.ConvertToVBSOrder(order, existingcustomer.CustomerNo);

                if (vorder == null)
                {
                    errmsg = "Error in ConvertToVBSOrder() ";
                    continue;
                }
                int vismaOrderNumber = vbs.AddVBSOrder(vorder, out errmsg);


                if (vismaOrderNumber <= 0)
                {
                    errmsg = "Error adding order - " + errmsg;
                    continue;
                }

                Utils.WriteLog("Added order " + vismaOrderNumber);

                // Verify...
                retries = 20;
                int vismaOrderNumberTest = 0;
                while (--retries > 0)
                {
                    vismaOrderNumberTest = db.ExistingVismaWebOrder(vorder.CustomerOrSupplierOrderNo, ref changeTime, out errmsg);
                    if (vismaOrderNumberTest > 0)
                        break;

                    Thread.Sleep(1000);
                }

                if (vismaOrderNumberTest <= 0)
                {
                    Utils.WriteLog("VERIFY ERROR: Could not verify order creation for web order " + order.id);
                }

                if (vismaOrderNumber > 0 && vismaOrderNumberTest > 0)
                {
                    // Update CardNm,CardAc
                    if (vorder.CardNm != "")
                    {
                        if (db.UpdateOrderCardNm(vismaOrderNumber, vorder.CardNm, "0", out errmsg) == false)
                            Utils.WriteLog("ERROR: db.UpdateOrderCardNm() - " + errmsg);
                    }
                    Models.Visma.VismaOrder orderReadback = new Models.Visma.VismaOrder();

                    /*
                    if (db.GetOrder(vismaOrderNumber, ref orderReadback, out errmsg))
                    {
                        // Adjust Gr2-product for each orderline
                        foreach (Models.OrderLine line in orderReadback.OrderLines)
                        {
                            int gr2 = 0;
                            if (dbOrders.GetProductGr2(line.SKU, ref gr2, out errmsg))
                            {
                                if (gr2 == 1)
                                    dbOrders.UpdateOrderLineProccessingMethod(vismaOrderNumber, line.LineNumber, true, out errmsg);
                                if (gr2 == 3)
                                    dbOrders.UpdateOrderLineProccessingMethod(vismaOrderNumber, line.LineNumber, false, out errmsg);
                            }
                        }
                    }*/


                }
                // Update phone , mobil on actor
                /*if (existingCustNo > 0 && order.LoginStatus != 2)
                {
                    if (dbOrders.UpdateCustomer(existingCustNo, order.Phone, order.Mobile, out errmsg) == false)
                        Utils.WriteLog(false, "ERROR: db.UpdateCustomer() - " + errmsg);
                }*/

            }



            return true;
        }

        public static string CheckOrderLineProduct(DBaccess db, Order order)
        {
            string errmsg = "";
            try
            {
                foreach (OrderLineItem line in order.line_items)
                {
                    if (string.IsNullOrEmpty(line.sku))
                        continue;

                    line.sku = line.sku.Replace("-en", "");

                    if (db.HasProdID(line.sku, out errmsg) == 0)
                    {
                        Utils.WriteLog($"ProdNo {line.sku} is unknown");
                        if (Utils.ReadConfigString("DefaultUnknownProductCode", "") != "")
                        {
                            line.sku = Utils.ReadConfigString("DefaultUnknownProductCode", "");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Utils.WriteLog("CheckOrderLineProduct()", ex);
                errmsg += ex.GetBaseException().Message;
            }

            return errmsg;
        }

        private static bool SyncOrderClose()
        {
            DBaccess db = new DBaccess();
            //   RestSharpClient client = new RestSharpClient();
            OrderUpdateHttpClient client = new OrderUpdateHttpClient();

            List<VismaOrderStatus> vismaOrderStatuses = new List<VismaOrderStatus>();
            if (db.GetAllWebOrdersToClose(ref vismaOrderStatuses, out string errmsg) == false)
            {
                Utils.WriteLog("ERROR db.GetAllWebOrdersToClose()  " + errmsg);
                return false;
            }

            Utils.WriteLog($"{vismaOrderStatuses.Count} completed orders to update in WooCommerce..");
            foreach (VismaOrderStatus orderStatus in vismaOrderStatuses)
            {
                int wooOrderID = Utils.StringToInt(orderStatus.CustomerOrSupplierOrderNo);
                if (wooOrderID <= 0)
                    continue;
                //  bool res = client.UpdateOrderStatus(wooOrderID.ToString(), "completed");
                bool res = client.UpdateOrder(wooOrderID.ToString(), "completed");
                if (res)
                {
                    Utils.WriteLog("Status for order " + wooOrderID.ToString() + " (Visma " + orderStatus.OrderNumber + ")  updated to 'completed' in WooCommerce");

                    DateTime now = DateTime.Now;

                    // Indicate status sent                        
                    // db2.UpdateStatusInf6(order.ordNo, order.status, out errmsg);
                    // res = client.CreateOrderNote(order.webOrderNo, string.Format("Updated {0:0000}-{1:00}-{2:00} {3:00}:{4:00}:{5:00} by Visma OrderFeedbackService", now.Year, now.Month, now.Day, now.Hour, now.Minute, now.Second));
                    res = client.CreateOrderNote(wooOrderID.ToString(), string.Format("Updated {0:0000}-{1:00}-{2:00} {3:00}:{4:00}:{5:00} by Visma OrderFeedbackService", now.Year, now.Month, now.Day, now.Hour, now.Minute, now.Second));


                    if (db.UpdateOrderStatusFlag(orderStatus.OrderNumber, orderStatus.OrderDocNumber, out errmsg) == false)
                        Utils.WriteLog($"Error UpdateOrderStatusFlag() - {errmsg}");

                }
            }

            return true;
        }

        internal static void Syns()
        {
            Sync sync = Utils.ReadSyncTime(Utils.ReadConfigString("OrderSyncFile", ""));

            DateTime st = sync.LastestSync;
            if (st != DateTime.MinValue)
            {
                st = st.AddHours(-2);
            }

            List<Order> wooCommerceOrders = WooCommerceHelpers.GetWooCommerceOrders(st).Result;
            if (wooCommerceOrders.Count > 0)
            {
                if (IsLieu_Dit && !VBSConvert.OrderConverts.Any(x => x is Lieu_Dit))
                    VBSConvert.OrderConverts.Add(new Lieu_Dit());

                if (SyncOrders(wooCommerceOrders) == true)
                    Utils.WriteSyncTime(sync, Utils.ReadConfigString("OrderSyncFile", ""));
            }

            SyncOrderClose();
        }
    }
}
