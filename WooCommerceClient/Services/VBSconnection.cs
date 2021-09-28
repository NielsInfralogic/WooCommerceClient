using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Visma.BusinessModel;
using Visma.BusinessServices.Client;
using Visma.BusinessServices.Generic;
using Visma.BusinessServices.Wrapper;
using WooCommerceClient.Models;
using WooCommerceClient.Services.SpecificHandling;

namespace WooCommerceClient.Services
{
    public class VBSconnection : BaseClass
    {
        private readonly string vbsUserName;
        private readonly string vbsPassword;
        private readonly int vbsCompanyNumber;
        private readonly string vismaUser;
        private readonly string vbsSiteID;
        GenericServiceClient client;

        public VBSconnection(string _vbssiteID, string _vbsUserName, string _vbsPassword, int _companyNumber, string _vismaUser)
        {
            vbsSiteID = _vbssiteID;
            vbsUserName = _vbsUserName;
            vbsPassword = _vbsPassword;
            vbsCompanyNumber = _companyNumber;
            vismaUser = _vismaUser;
        }

        public bool TestConnection(out string messsage)
        {
            messsage = "";
            bool status = false;
            try
            {
                client = new GenericServiceClient();
                (new Credentials(vbsSiteID, vbsUserName, vbsPassword)).Apply(client.ClientCredentials);

                RequestBuilder requestBuilder = new RequestBuilder();
                Context context = requestBuilder.AddContext();
                context.CompanyNo = vbsCompanyNumber;
                context.CultureId = CultureId.Danish__Denmark;

                context.UserName = vismaUser;
                context.SiteId = vbsSiteID;
                ResponseReader responseReader = requestBuilder.Dispatch(client);
                if (responseReader == null)
                {
                    messsage = string.Concat("ResponseReader null");
                    status = false;
                }
                else
                {
                    if (responseReader.Status != Response.StatusCode.Ok)
                    {
                        messsage = string.Concat("responseReader.Status NOT OK - ", responseReader.Status.ToString());
                        status = false;
                    }
                    else
                        status = true;
                }
                client.Close();
            }
            catch (Exception exception)
            {
                messsage = string.Concat("VBSconnection.TestConnection() Exception ", exception.Message);
            }

            return status;
        }

        private int LookupDeliveryMethod(string shippingMethod)
        {
            if (shippingMethod.ToLower().IndexOf("postnord pickup point") != -1)
                return 20;
            if (shippingMethod.ToLower().IndexOf("postnord omdeling til privat") != -1 || shippingMethod.ToLower().IndexOf("postnord - omdeling til privat") != -1) 
                return 22;
            if (shippingMethod.ToLower().IndexOf("postnord erhverv") != -1 || shippingMethod.ToLower().IndexOf("postnord busi") != -1)
                return 21;
            if (shippingMethod.ToLower().IndexOf("gls privat") != -1)
                return 40;
            if (shippingMethod.ToLower().IndexOf("gls erhverv") != -1)
                return 41;
            if (shippingMethod.ToLower().IndexOf("gls udlev") != -1)
                return 42;       
            if (shippingMethod.ToLower().IndexOf("burd") != -1)
                return 30;
            if (shippingMethod.ToLower().IndexOf("afhentning") != -1 || shippingMethod.ToLower().IndexOf("afhentes") != -1)
                return 1;

            
            return 5;

        }

        private int LookupPaymentMethod(string method)
        {
            if (method.ToLower().IndexOf("kortbetaling") != -1 || method.ToLower().IndexOf("quickpay") != -1)
                return 6;
            if (method.ToLower().IndexOf("mobilepay") != -1 )
                return 7;
            if (method.ToLower().IndexOf("cash on delivery") != -1 ||
                method.ToLower().IndexOf("cod") != -1)
                return 4;

            return 54;
        }

        public Models.Visma.VismaOrder ConvertToVBSOrder(Models.Order order, int vismaCustNo)
        {
            Models.Visma.VismaOrder vorder = new Models.Visma.VismaOrder()
            {
                OrderPreferences = Utils.ReadConfigInt32("OrderPreferences", 256)
            };

            DBaccess db = new DBaccess();
            try
            {
                vorder.OrderLines = new List<Models.Visma.VismaOrderLine>();
                vorder.CustomerNo = vismaCustNo;

                vorder.OrderNo = 0; // auto-generated..
                vorder.OrderDate = Utils.DateTimeToVismaDate(order.date_created.HasValue ? order.date_created.Value : DateTime.MinValue);
                if (vorder.OrderDate < 20000101)
                    vorder.OrderDate = DateTime.Now.Year * 10000 + DateTime.Now.Month * 100 + DateTime.Now.Day;

                vorder.Name = order.billing.first_name + " " + order.billing.last_name;
                vorder.Name = vorder.Name.Trim();
                vorder.AddressLine1 = order.billing.address_1 ??"";
                vorder.AddressLine2 = order.billing.address_2 ?? "";
                vorder.AddressLine3 = "";
                vorder.PostalArea = order.billing.city ?? "";
                vorder.PostCode = order.billing.postcode ?? "";
                vorder.CountryNumber = db.LookupCountryCode(order.billing.country ?? "", out string errmsg);
                vorder.Phone = order.billing.phone;

                vorder.CurrencyNo = db.LookupCurrencyCode(order.currency ?? "DKK", out errmsg);
                string companyName = order.shipping.company ?? "";

                vorder.DeliveryName = order.shipping.first_name + " " + order.shipping.last_name;
                bool usesCompany = false;
                if (companyName.Trim() != "")
                {
                    vorder.DeliveryName = companyName;
                    usesCompany = true;
                }
                else
                {
                    companyName = order.billing.company ?? "";
                    if (companyName.Trim() != "")
                    {
                        vorder.DeliveryName = companyName;
                        usesCompany = true;
                    }
                }
                              
                if (usesCompany)
                {
                    vorder.DeliveryAddress1 = order.shipping.address_1 ?? "";
                    vorder.DeliveryAddress2 = order.shipping.address_2 ?? "";

                    if (order.shipping.first_name != "" && order.shipping.last_name != "")
                    {
                        vorder.DeliveryAddress1 = order.shipping.first_name + " " + order.shipping.last_name;
                        vorder.DeliveryAddress2 = order.shipping.address_1 ?? "";
                        vorder.DeliveryAddress3 = order.shipping.address_2 ?? "";
                    }

                }
                else
                {
                    vorder.DeliveryAddress1 = order.shipping.address_1 ?? "";
                    vorder.DeliveryAddress2 = order.shipping.address_2 ?? "";
                    vorder.DeliveryAddress3 = "";
                }

                vorder.DeliveryPostCode = order.shipping.postcode;
                vorder.DeliveryPostalArea = order.shipping.city ?? "";

                vorder.DeliveryCountryNumber = db.LookupCountryCode(order.shipping.country ?? "", out errmsg);

                /* vorder.DeliveryMethod = order.DeliveryMethod ?? 1;
                 if (vorder.DeliveryMethod == 2)
                     vorder.DeliveryMethod = 6;
                 vorder.RequiredDeliveryDate = Utils.DateTimeToVismaDate(Utils.StringToDate(order.DeliveryDate));
                 if (vorder.RequiredDeliveryDate < 20000101)
                     vorder.RequiredDeliveryDate = 0;*/


                vorder.Information1 = order.billing.phone ?? ""; // order.order_key;
                vorder.Information2 = order.billing.email ?? "";
                vorder.Information3 = order.transaction_id ?? "";
                vorder.Information4 = order.payment_method_title ?? "";
                //vorder.Information4 = !string.IsNullOrWhiteSpace(order.payment_method) ?
                //    order.payment_method : order.payment_method_title ?? "";

                // vorder.Information6 = order.number;
                vorder.Freight1 = order.shipping_total ?? 0.0M;

                // Pakkeshop id in Inf5
                List<OrderMeta> metaList = order.meta_data;
                string language = "da";
                if (metaList != null)
                {
                    foreach (OrderMeta orderMeta in metaList)
                    {
                        if (orderMeta.key == "Pakkeshop")
                        {
                            try
                            {
                                vorder.Information5 = orderMeta.value as string;
                                break;
                            }
                            catch { }
                        }

                        if (orderMeta.key == "Udleveringssted")
                        {
                            try
                            {
                                vorder.Information5 = orderMeta.value as string;
                                break;
                            }
                            catch { }
                        }

                        if (orderMeta.key == "wpml_language")
                        {
                            try
                            {
                                language = orderMeta.value as string;
                                break;
                            }
                            catch { }
                        }

                        if (orderMeta.key == "QUICKPAY_PAYMENT_ID" || orderMeta.key == "_quickpay_transaction_id")
                            vorder.Information3 = orderMeta.value as string;
                    }
                }

                vorder.LanguageNo = language.ToLower() == "en" ? 44 : 45;



               

                if (order.shipping_lines != null)
                {
                    Utils.WriteLog($"Shipping lines {order.shipping_lines.Count}");
                    foreach (OrderShippingLine shippingLine in order.shipping_lines)
                    {
                        if (string.IsNullOrEmpty(shippingLine.method_title) == false)
                        {
                            // PostNord Pickup Point
                            // PostNord Pickup Point (gratis)
                            // Afhentning på lager
                            // Gratis forsendelse
                            // Flat rate
                            vorder.Information6 = shippingLine.method_title;
                            break;
                        }
                    }
                    foreach (OrderShippingLine shippingLine in order.shipping_lines)
                    {
                        if (shippingLine.total.HasValue)
                        {
                            if (shippingLine.total.Value > vorder.Freight1)
                                vorder.Freight1 = shippingLine.total.Value;
                        }
                    }


                }

                vorder.DeliveryMethod = LookupDeliveryMethod(vorder.Information6);
                vorder.PaymentMethod = LookupPaymentMethod(vorder.Information4);



                vorder.Memo = order.customer_note;
                vorder.CustomerOrSupplierOrderNo = order.id.ToString();


                vorder.Group1 = (int)Utils.ReadConfigInt32("OrderGroup1", 2);
                vorder.Group12 = (int)0;


                vorder.WarehouseNo = (int)Utils.ReadConfigInt32("WarehouseNo", 1);

                vorder.CardNm = "";
                vorder.CardAc = "0";
                int lnNo = 1;
                if (order.line_items != null && order.line_items.Count > 0)
                    ConvertOrderLines(order, vorder, db, ref errmsg, ref lnNo);

                if (vorder.Freight1 > 0.0M && Utils.ReadConfigString("ProdNoFreight", "") != "")
                {
                    Models.Visma.VismaOrderLine vlinefragt = new Models.Visma.VismaOrderLine()
                    {
                        LineNo = lnNo++,
                        Quantity = 1.0M,
                        PriceInCurrency = vorder.Freight1,
                        ProductNo = Utils.ReadConfigString("ProdNoFreight", ""),
                        Description = "Fragt"
                    };
                    if (Utils.ReadConfigInt32("AddVatToPrices", 1) > 0)
                        vlinefragt.PriceInCurrency *= 1.25M;

                    Utils.WriteLog($"Line {lnNo - 1} ProdNo:{vlinefragt.ProductNo} Desc:{vlinefragt.Description} Qty:{vlinefragt.Quantity} Price:{vlinefragt.PriceInCurrency} Discount:{vlinefragt.Discount}");
                    vorder.OrderLines.Add(vlinefragt);
                    vorder.Freight1 = 0.0M;
                }


            }
            catch(Exception ex)
            {
                Utils.WriteLog($"Exception ConvertToVBS() - {ex.Message}");
                return null;
            }
            return vorder;
        }

        public static void ConvertOrderLines(Order order, Models.Visma.VismaOrder vorder, DBaccess db, ref string errmsg, ref int lnNo)
        {
            foreach (Models.OrderLineItem line in order.line_items)
            {

                Models.Visma.VismaOrderLine vline = new Models.Visma.VismaOrderLine
                {
                    LineNo = lnNo++,
                    Quantity = line.quantity ?? 0.0M
                };
                errmsg = ConvertOrderLine(db, lnNo, order, vorder, line, vline);

                vorder.OrderLines.Add(vline);
            }
        }

        public static string ConvertOrderLine(DBaccess db, int lnNo, Order order, Models.Visma.VismaOrder vorder, OrderLineItem line, Models.Visma.VismaOrderLine vline)
        {
            string errmsg = "";
            Utils.WriteLog($"Order line {lnNo}..");
            if (line.sku == null)
                vline.ProductNo = Utils.ReadConfigString("EmptyProductNumber", "Diverse");
            else
                vline.ProductNo = line.sku != "" ? line.sku : Utils.ReadConfigString("EmptyProductNumber", "Diverse");
            vline.PriceInCurrency = line.price ?? 0.0M;
            vline.DiscountPercent = 0.0M;
            vline.Discount = 0.0M;
            vline.Description = line.name ?? "";
            if (line.subtotal.HasValue && line.total.HasValue)
            {
                vline.PriceInCurrency = vline.Quantity != 0.0M ? line.subtotal.Value / vline.Quantity : line.subtotal.Value;
                vline.Discount = vline.Quantity != 0.0M ? (line.subtotal.Value - line.total.Value) / vline.Quantity : (line.subtotal.Value - line.total.Value);
                vline.DiscountPercent = vline.Discount != 0.0M ? 100.0M * (vline.Discount / vline.PriceInCurrency) : 0.0M;
            }
            vline.LinePrice = line.total ?? 0.0M; // net price..

            if (Utils.ReadConfigInt32("AddVatToPrices", 0) > 0) //Default: do not add vat
            {
                vline.LinePrice *= 1.25M;
                vline.Discount *= 1.25M;
                vline.PriceInCurrency *= 1.25M;
            }
            vline.WarehouseNo = (int)Utils.ReadConfigInt32("WarehouseNo", 1);

            int stockOnStcNo3 = 0;
            if (db != null && db.GetStock(vline.ProductNo, 3, ref stockOnStcNo3, out errmsg) == false)
                Utils.WriteLog($"Error: GetStock() {errmsg}");

            if (stockOnStcNo3 >= vline.Quantity)
                vline.WarehouseNo = 3;

            VBSConvert.OrderLineConvert(line, order, vline, vorder);

            Utils.WriteLog($"Line {lnNo - 1} ProdNo:{vline.ProductNo} Desc:{vline.Description} Qty:{vline.Quantity} Price:{vline.PriceInCurrency} Discount:{vline.Discount} Stock;{vline.WarehouseNo}");
            return errmsg;
        }

        public int AddVBSOrder(Models.Visma.VismaOrder order, out string errmsg)
        {
            errmsg = "";
            int result = 0;

            string memoFileName = GetMemoFilename(order);

            using (GenericServiceClient genericServiceClient = new GenericServiceClient())
            {
                new Credentials(vbsSiteID, vbsUserName, vbsPassword).Apply(genericServiceClient.ClientCredentials);
                RequestBuilder requestBuilder = new RequestBuilder();
                Context context = requestBuilder.AddContext();
                context.CompanyNo = vbsCompanyNumber;
                context.CultureId = CultureId.Danish__Denmark;
                context.UserName = vismaUser;
                context.SiteId = vbsSiteID;

                ResponseReader responseReader = null;

                TableHandler tableHandler = context.UseTable((long)T.Order);
                Row row = tableHandler.AddRow();

                Utils.WriteLog("Order header for order " + order.CustomerOrSupplierOrderNo);
                row.SuggestValue((long)C.Order.OrderNo);
                try
                {
                    CultureInfo invariantCulture = CultureInfo.InvariantCulture;

                    row.SetIntegerValue((long)C.Order.OrderDate, (long)order.OrderDate);
                    row.SetIntegerValue((long)C.Order.OrderType, (long)Utils.ReadConfigInt32("OrderType", 1));
                    row.SetIntegerValue((long)C.Order.TransactionType, (long)1);
                    row.SetIntegerValue((long)C.Order.CustomerNo, (long)order.CustomerNo);

                    if (Utils.ReadConfigInt32("SetOrderPreferences", 0) > 0)
                        row.SetIntegerValue((long)C.Order.OrderPreferences, (long)order.OrderPreferences);



                    //      if (order.CustomerNo != Utils.ReadConfigInt32("AnonymousCustNo", 5) || Utils.ReadConfigBoolean("UseAnonymousNameInOrder", false)
                    //{




                    if (order.Name != "")
                        row.SetStringValue((long)C.Order.Name, order.Name);

                    row.SetStringValue((long)C.Order.AddressLine1, order.AddressLine1);
                    row.SetStringValue((long)C.Order.AddressLine2, order.AddressLine2);
                    row.SetStringValue((long)C.Order.AddressLine3, order.AddressLine3);
                    row.SetStringValue((long)C.Order.PostCode, order.PostCode);
                    row.SetStringValue((long)C.Order.PostalArea, order.PostalArea);
                    if (order.CountryNumber > 0)
                        row.SetIntegerValue((long)C.Order.CountryNo, (long)order.CountryNumber);
                    Utils.WriteLog("Order.CountryNo = " + order.CountryNumber);

                    if (order.LanguageNo > 0)
                        row.SetIntegerValue((long)C.Order.LanguageNo, (long)order.LanguageNo);

                    row.SetStringValue((long)C.Order.DeliveryName, order.DeliveryName);
                    row.SetStringValue((long)C.Order.DeliveryAddress1, order.DeliveryAddress1);
                    row.SetStringValue((long)C.Order.DeliveryAddress2, order.DeliveryAddress2);
                    row.SetStringValue((long)C.Order.DeliveryAddress3, order.DeliveryAddress3);
                    row.SetStringValue((long)C.Order.DeliveryPostCode, order.DeliveryPostCode);
                    row.SetStringValue((long)C.Order.DeliveryPostalArea, order.DeliveryPostalArea);
                    if (order.DeliveryCountryNumber > 0)
                        row.SetIntegerValue((long)C.Order.DeliveryCountryNo, (long)order.DeliveryCountryNumber);
                    Utils.WriteLog("Order.DeliveryCountryNo = " + order.DeliveryCountryNumber);

                    if (order.PaymentMethod > 0)
                        row.SetIntegerValue((long)C.Order.PaymentMethod, order.PaymentMethod);

                    if (order.DeliveryMethod > 0)
                        row.SetIntegerValue((long)C.Order.DeliveryMethod, order.DeliveryMethod);
                    if (order.RequiredDeliveryDate > 0)
                        row.SetIntegerValue((long)C.Order.RequiredDeliveryDate, (long)order.RequiredDeliveryDate);

                    row.SetStringValue((long)C.Order.CustomerOrSupplierOrderNo, order.CustomerOrSupplierOrderNo);  // ShopOrderID

                    row.SetStringValue((long)C.Order.Information1, order.Information1);    // Phone
                    row.SetStringValue((long)C.Order.Information2, order.Information2);  // Email
                    row.SetStringValue((long)C.Order.Information3, order.Information3);  // Quickpay
                    row.SetStringValue((long)C.Order.Information4, order.Information4);  // DeliveryComment
                    row.SetStringValue((long)C.Order.Information5, order.Information5);  // Phone
                    row.SetStringValue((long)C.Order.Information6, order.Information6);  // Ship methood
                    row.SetIntegerValue((long)C.Order.Group1, (long)order.Group1);
                    //row.SetIntegerValue((long)C.Order.Group12, (long)order.Group12);

                    // 20180831
                    // row.SetDecimalValue((long)C.Order.Freight1, order.Freight1);           //FreightAmount

                    if (order.WebPage != "")
                    {
                        row.SetStringValue((long)C.Order.WebPage, order.WebPage);       // Comment
                    }

                    // 20180901
                    if (memoFileName != "")
                    {
                        row.SetStringValue((long)C.Order.MemoFileName, memoFileName);       // Comment
                    }

                    if (order.WarehouseNo > 0)
                        row.SetIntegerValue((long)C.Order.WarehouseNo, (long)order.WarehouseNo);

                    // 20180927
                    row.SetDecimalValue((long)C.Order.Free1, order.Free1);           //Quickpay_amount

                    // 20181107
                    /*   if (order.CardNm != "")
                           row.SetStringValue((long)C.Order.CardCompany, order.CardNm);    // InternOrderID
                       if (order.CardAc != "")
                           row.SetStringValue((long)C.Order.MaskedCardNo, order.CardAc);    // State for StatusFeedbackService to WooCommerce..                    
   */

                    if (order.CurrencyNo > 0)
                        row.SetIntegerValue((long)C.Order.CurrencyNo, (long)order.CurrencyNo);

                    TableHandler tableHandler2 = row.JoinTable((long)Fk.OrderLine.Order);

                    int lnno = 1;

                    foreach (Models.Visma.VismaOrderLine line in order.OrderLines)
                    {
                        Utils.WriteLog("Order line " + lnno.ToString() + " " + line.ProductNo + " " + line.Quantity.ToString() + " " + line.PriceInCurrency.ToString());
                        lnno++;

                        Row row2 = tableHandler2.AddRow();
                        row2.SuggestValue((long)C.OrderLine.LineNo);



                        //           if (line.Comment != "")
                        //             row2.SetStringValue((long)C.OrderLine.Free1, line.Comment);



                        //                        if (line.SetProcessingMethod)
                        //                          row2.SetIntegerValue((long)C.OrderLine.ProcessingMethod2, 16384);

                        Assignment assignment = row2.SetStringValue((long)C.OrderLine.ProductNo, line.ProductNo);
                        assignment.Operation.EscalateErrorToOperation = row;

                        if (line.WarehouseNo > 0)
                            row2.SetIntegerValue((long)C.OrderLine.WarehouseNo, line.WarehouseNo);

                        if (line.Description != "")
                        {
                            assignment = row2.SetStringValue((long)C.OrderLine.Description, line.Description); // Product name
                            assignment.Operation.EscalateErrorToOperation = row;
                        }


                        if (line.ProductNo.Trim() != "")
                        {
                            assignment = row2.SetDecimalValue((long)C.OrderLine.Quantity, line.Quantity);               // OrdLn.NoInvoAB
                            assignment.Operation.EscalateErrorToOperation = row;

                            if (order.CurrencyNo > 0)
                                row2.SetIntegerValue((long)C.OrderLine.CurrencyNo, (long)order.CurrencyNo);

                            if (line.PriceInCurrency > 0.0M)
                            {
                                assignment = row2.SetDecimalValue((long)C.OrderLine.PriceInCurrency, line.PriceInCurrency);           // OrdLn.Price
                                assignment.Operation.EscalateErrorToOperation = row;

                                assignment = row2.SetDecimalValue((long)C.OrderLine.DiscountPercent1, line.DiscountPercent);
                                assignment.Operation.EscalateErrorToOperation = row;
                            }
                            else if (line.PriceInCurrency == 0.0M)
                            {
                                assignment = row2.SetDecimalValue((long)C.OrderLine.PriceInCurrency, 0.0M);           // OrdLn.Price
                                assignment.Operation.EscalateErrorToOperation = row;
                                assignment = row2.SetDecimalValue((long)C.OrderLine.DiscountPercent1, 100.0M);
                                assignment.Operation.EscalateErrorToOperation = row;
                            }
                            // assignment = row2.SetDecimalValue((long)C.OrderLine.CostPriceInCurrency, line.CostPrice);
                            // assignment.Operation.EscalateErrorToOperation = row;
                            // assignment = row2.SetDecimalValue((long)C.OrderLine.DiscountAmount1InCurrency, line.Discount);
                            // assignment.Operation.EscalateErrorToOperation = row;
                            //assignment = row2.SetDecimalValue((long)C.OrderLine.VatAmountInCurrency, line.VatAmount);  
                            //assignment.Operation.EscalateErrorToOperation = row;
                            //assignment = row2.SetDecimalValue((long)C.OrderLine.AmountInCurrency, line.LinePrice);      // OrdLn.Am
                            //assignment.Operation.EscalateErrorToOperation = row; 

                        }


                        // }
                    }
                    Projection projection = row.ProjectColumns();
                    projection.AddColumn((long)C.Order.OrderNo);
                    Utils.WriteLog("AddVbsOrder: calling Dispatch..");
                    responseReader = requestBuilder.Dispatch(genericServiceClient);
                    bool allSucceeded = responseReader.AllSucceeded;
                    if (allSucceeded)
                    {
                        Utils.WriteLog("AddVbsOrder: Dispatch ok");
                        foreach (OperationResult res in responseReader.OperationResultDictionary.Values)
                        {
                            if (res is ProjectionResult)
                            {
                                ProjectionResult projectionResult = res as ProjectionResult;
                                ResultRow resultRow = projectionResult.ResultSet.ResultRows[0];
                                result = (int)resultRow.Values[0];
                            }
                        }
                    }
                    else
                    {
                        Utils.WriteLog("AddVbsOrder: Dispatch failed");
                        errmsg = GetErrorMessages("AddVbsOrder", errmsg, responseReader);
                    }
                }
                catch (Exception ex)
                {
                    errmsg = "Exception: " + ex.Message;
                    if (ex.InnerException != null)
                        errmsg += " Inner Exception: " + ex.InnerException.Message;
                    Utils.WriteLog("AddVbsOrder: " + errmsg);
                }

                if (context != null)
                    context = null;

                if (responseReader != null)
                    responseReader = null;

                if (requestBuilder != null)
                    requestBuilder = null;

                if (genericServiceClient != null)
                    genericServiceClient.Close();

            }


            return result;
        }

        private static string GetMemoFilename(Models.Visma.VismaOrder order)
        {
            string memoFileName = "";
            if (order.Memo != "")
            {

                string memoPath = Utils.ReadConfigString("MemoPath", @"v:\Memo\F0001");

                memoFileName = memoPath + @"\M" + Utils.GenerateTimeStamp() + ".txt";
                if (Utils.WriteToMemoFile(memoFileName, order.Memo) == false)
                    memoFileName = "";
            }

            return memoFileName;
        }

        private static string GetErrorMessages(string title, string errmsg, ResponseReader responseReader)
        {
            if (responseReader.Messages.Count > 0)
            {
                errmsg += GetMessages("1", responseReader.Messages, errmsg, $"ERROR {title}");
            }
            ReadOnlyCollection<OperationResult> operationsWithErrors = responseReader.GetOperationsWithErrors();
            foreach (OperationResult current4 in operationsWithErrors)
            {
                errmsg += GetMessages("2", current4.Messages, errmsg, $"ERROR {title}");
                ReadOnlyCollection<RowResult> rowsWithErrors = current4.GetRowsWithErrors();
                using (IEnumerator<RowResult> enumerator6 = rowsWithErrors.GetEnumerator())
                {
                    while (enumerator6.MoveNext())
                    {
                        AssignmentRowResult assignmentRowResult = (AssignmentRowResult)enumerator6.Current;
                        if (errmsg != "")
                            errmsg += "Statuscode:" + assignmentRowResult.StatusObject.ToString();
                        errmsg += GetMessages("3", assignmentRowResult.Messages, errmsg, $"ERROR {title}");
                    }
                }
            }

            return errmsg;
        }

        private static string GetMessages(string prefix, ReadOnlyCollection<Message> messages, string errmsg, string logTitle = null)
        {
            if (messages == null || messages.Count == 0)
                return "";

            var msg = errmsg != "" ? ", " : "";
            msg += string.Join("\n", messages.Select(x =>
                x.Text + $"({prefix} " + x.Severity.ToString() + ")"));

            if (!string.IsNullOrEmpty(logTitle))
                Utils.WriteLog($"{logTitle}: {msg}");

            return errmsg != "" ? ", " + msg : msg;
        }


        // Returns CustomerNo
        public bool AddVBSCustomer(Models.Visma.VismaActor actor, int custNoToUseForced, ref int ActNoCreated, ref int CustNoCreated, out string errmsg)
        {
            errmsg = "";
            CustNoCreated = 0;
            ActNoCreated = 0;
            bool allSucceeded = false;
            int seller = Utils.ReadConfigInt32("SellerNo", 47);

            using (GenericServiceClient genericServiceClient = new GenericServiceClient())
            {
                new Credentials(vbsSiteID, vbsUserName, vbsPassword).Apply(genericServiceClient.ClientCredentials);

                RequestBuilder requestBuilder = new RequestBuilder();
                Context context = requestBuilder.AddContext();
                context.CompanyNo = vbsCompanyNumber;
                context.CultureId = CultureId.Danish__Denmark;
                context.UserName = vismaUser;
                context.SiteId = vbsSiteID;

                ResponseReader responseReader = null;

                TableHandler tableHandler = context.UseTable((long)T.Associate);
                Row row = tableHandler.AddRow();
                try
                {
                    row.SuggestValue((long)C.Associate.AssociateNo);
                    if (custNoToUseForced <= 0)
                        row.SuggestValue((long)C.Associate.CustomerNo);
                    else
                        row.SetIntegerValue((long)C.Associate.CustomerNo, (long)custNoToUseForced);

                    if (Utils.ReadConfigInt32("AssociateProcessing", 0) > 0)
                        row.SetOn((long)C.Associate.AssociateProcessing, Utils.ReadConfigInt32("AssociateProcessing", 0));

                    // 
                    if (Utils.ReadConfigInt32("TemplateActor", 0) > 0)
                        row.SetIntegerValue((long)C.Associate.AssociateTemplate, (long)Utils.ReadConfigInt32("TemplateActor", 1760));

                    row.SetStringValue((long)C.Associate.Name, actor.Name);
                    row.SetStringValue((long)C.Associate.EmailAddress, actor.EmailAddress);
                    row.SetStringValue((long)C.Associate.Phone, actor.Phone);
                    row.SetStringValue((long)C.Associate.MobilePhone, actor.Mobile);

                    row.SetStringValue((long)C.Associate.AddressLine1, actor.AddressLine1);
                    row.SetStringValue((long)C.Associate.AddressLine2, actor.AddressLine2);
                    row.SetStringValue((long)C.Associate.AddressLine3, actor.AddressLine3);
                    row.SetStringValue((long)C.Associate.PostCode, actor.PostCode);
                    row.SetStringValue((long)C.Associate.PostalArea, actor.PostalArea);
                    if (actor.CountryNo > 0)
                        row.SetIntegerValue((long)C.Associate.CountryNo, (long)actor.CountryNo);

                    if (actor.LanguageNo > 0)
                        row.SetIntegerValue((long)C.Associate.LanguageNo, (long)actor.LanguageNo);


                    if (actor.CurrencyNo > 0)
                        row.SetIntegerValue((long)C.Associate.CurrencyNo, (long)actor.CurrencyNo);

                    Projection projection = row.ProjectColumns();
                    projection.AddColumn((long)C.Associate.AssociateNo);
                    projection.AddColumn((long)C.Associate.CustomerNo);
                    responseReader = requestBuilder.Dispatch(genericServiceClient);
                    Utils.WriteLog("AddVBSCustomer()..Dispatched");

                    allSucceeded = responseReader.AllSucceeded;
                    if (allSucceeded)
                    {
                        //            Utils.WriteLog(false, "AddVBSCustomer()..Reading response 1");
                        // Prefs.logger.WriteLog("Dispatching customer succeeded: (Name:" + actor.Name + ")");
                        foreach (OperationResult current in responseReader.OperationResultDictionary.Values)
                        {
                            //              Utils.WriteLog(false, "AddVBSCustomer()..Reading response 2");
                            if (current is ProjectionResult)
                            {
                                Utils.WriteLog("AddVBSCustomer()..Reading response 3");
                                ProjectionResult projectionResult = current as ProjectionResult;
                                ResultRow resultRow = projectionResult.ResultSet.ResultRows[0];
                                CustNoCreated = resultRow.GetIntegerValue((long)C.Associate.CustomerNo);
                                ActNoCreated = resultRow.GetIntegerValue((long)C.Associate.AssociateNo);
                            }
                        }
                    }
                    else
                    {
                        Utils.WriteLog("Dispatching customer failed: (Name:" + actor.Name + ") " + vbsUserName + " " + vbsPassword + " " + vbsCompanyNumber.ToString() + " " + vbsSiteID);
                        foreach (Message msg in responseReader.Messages)
                            errmsg += "Error text:" + msg.Text + " (" + msg.Severity + ")";

                        ReadOnlyCollection<OperationResult> operationsWithErrors = responseReader.GetOperationsWithErrors();
                        foreach (OperationResult msg in operationsWithErrors)
                        {
                            foreach (Message msg2 in msg.Messages)
                                errmsg += "Error text:" + msg2.Text + " (" + msg2.Severity + ")";

                            ReadOnlyCollection<RowResult> rowsWithErrors = msg.GetRowsWithErrors();
                            using (IEnumerator<RowResult> enumerator = rowsWithErrors.GetEnumerator())
                            {
                                while (enumerator.MoveNext())
                                {
                                    AssignmentRowResult assignmentRowResult = (AssignmentRowResult)enumerator.Current;
                                    Utils.WriteLog("Statuscode:" + assignmentRowResult.StatusObject.ToString());
                                    foreach (Message msg3 in assignmentRowResult.Messages)
                                        errmsg += "Error text:" + msg3.Text + " (" + msg3.Severity.ToString() + ")";

                                }
                            }
                        }
                    }
                }

                catch (Exception Ex)
                {
                    errmsg = "Exception: " + Ex.Message;
                    Utils.WriteLog("Exception in AddVBSCustomer()  - " + Ex.Message);
                }

                if (context != null)
                    context = null;

                if (responseReader != null)
                    responseReader = null;

                if (requestBuilder != null)
                    requestBuilder = null;

                if (genericServiceClient != null)
                    genericServiceClient.Close();

            }
            return allSucceeded;
        }

        public bool GetVBSCustomer(string email, ref Models.Visma.VismaActor actor, out string errmsg)
        {
            actor.CustomerNo = 0;
            errmsg = "";

            using (GenericServiceClient genericServiceClient = new GenericServiceClient())
            {
                new Credentials(vbsSiteID, vbsUserName, vbsPassword).Apply(genericServiceClient.ClientCredentials);
                //new Credentials("VBus1100", "system", "").Apply(genericServiceClient.ClientCredentials);
                //new VismaLicenseCredentials(new Guid("257564c8-c5d4-4101-9b94-97a821ea1ad3"),"system", string.Empty).Apply(genericServiceClient.ClientCredentials);

                RequestBuilder requestBuilder = new RequestBuilder();
                Context context = requestBuilder.AddContext();
                context.CompanyNo = vbsCompanyNumber;
                context.CultureId = CultureId.Danish__Denmark;

                context.UserName = vismaUser;
                context.SiteId = vbsSiteID;
                try
                {
                    TableHandler tableHandler = context.UseTable((long)T.Associate);
                    RowsSelection rowsSelection = tableHandler.SelectRows();
                    // rowsSelection.IntegerColumnValue((long)C.Associate.CustomerNo, ComparisonOperator.EqualTo, (long)actNoToTest);
                    rowsSelection.StringColumnValue((long)C.Associate.EmailAddress, ComparisonOperator.EqualTo, email);

                    Rows rows = rowsSelection.Rows;

                    Projection projection = rows.ProjectColumns();

                    projection.AddColumn((long)C.Associate.CustomerNo);
                    projection.AddColumn((long)C.Associate.Name);
                    projection.AddColumn((long)C.Associate.EmailAddress);
                    projection.AddColumn((long)C.Associate.AddressLine1);
                    projection.AddColumn((long)C.Associate.AddressLine2);
                    projection.AddColumn((long)C.Associate.AddressLine3);
                    projection.AddColumn((long)C.Associate.PostalArea);
                    projection.AddColumn((long)C.Associate.PostCode);
                    projection.AddColumn((long)C.Associate.Phone);
                    projection.AddColumn((long)C.Associate.MobilePhone);
                    projection.AddColumn((long)C.Associate.ChangedDate);
                    //projection.AddColumn((long)C.Associate.AssociateProcessing);
                    Utils.WriteLog("GetVBSCustomer() - readding response 0..");

                    ResponseReader responseReader = requestBuilder.Dispatch(genericServiceClient);
                    OperationResult operationResult = null;
                    if (responseReader.OperationResultDictionary.TryGetValue(projection.Operation.OperationNo, out operationResult))
                    {
                        Utils.WriteLog("GetVBSCustomer() - readding response..");
                        ProjectionResult projectionResult = operationResult as ProjectionResult;
                        if (projectionResult != null)
                        {
                            Utils.WriteLog("GetVBSCustomer() - readding response 2..");
                            actor = projectionResult.ResultSet.ResultRows.Select(row => new Models.Visma.VismaActor
                            {
                                CustomerNo = row.GetIntegerValue((long)C.Associate.CustomerNo),
                                Name = row.GetStringValue((long)C.Associate.Name),
                                AddressLine1 = row.GetStringValue((long)C.Associate.AddressLine1),
                                AddressLine2 = row.GetStringValue((long)C.Associate.AddressLine2),
                                PostCode = row.GetStringValue((long)C.Associate.PostCode),
                                PostalArea = row.GetStringValue((long)C.Associate.PostalArea),
                                Phone = row.GetStringValue((long)C.Associate.Phone),
                                EmailAddress = row.GetStringValue((long)C.Associate.EmailAddress)
                            }).FirstOrDefault();
                        }

                    }

                }
                catch (Exception Ex)
                {
                    errmsg = "Exception: " + Ex.Message;
                    Utils.WriteLog("GetVBSCustomer() - " + errmsg);
                    return false;
                }
            }
            Utils.WriteLog("GetVBSCustomer() - returned " + actor.CustomerNo);
            return actor.CustomerNo > 0;
        }
    }
}
