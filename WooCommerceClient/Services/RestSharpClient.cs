using Newtonsoft.Json;
using RestSharp;
using RestSharp.Authenticators;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using WooCommerceClient.Models;

namespace WooCommerceClient.Services
{
    public class RestSharpClient
    {
        public bool UpdateOrderStatus(string orderID, string status1)
        {
            try
            {
                Utils.WriteLog("Setting status " + status1 + " for order id " + orderID);
                var client = new RestClient(Utils.ReadConfigString("WooCommerceUrl", ""))
                {
                    Authenticator = new HttpBasicAuthenticator(Utils.ReadConfigString("WooCommerceKey", ""), Utils.ReadConfigString("WooCommerceSecret", ""))
                };

                var request = new RestRequest("orders/" + orderID, Method.PUT);
                request.AddHeader("Accept", "application/json");
                request.Parameters.Clear();
                request.RequestFormat = DataFormat.Json;
                OrderStatus orderStatus = new OrderStatus() { status = status1 };


                request.AddJsonBody(orderStatus);
               

                var response = client.Execute(request);


                if (response.StatusCode == System.Net.HttpStatusCode.Accepted ||
                    response.StatusCode == System.Net.HttpStatusCode.OK ||
                    response.StatusCode == System.Net.HttpStatusCode.Created)
                {
                    Utils.WriteLog("Order " + orderID.ToString() + " updated to status '" + status1 + "'");
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


        public bool CreateOrderNote(string orderID, string note)
        {
            try
            {
                var client = new RestClient(Utils.ReadConfigString("WooCommerceUrl", ""))
                {
                    Authenticator = new HttpBasicAuthenticator(Utils.ReadConfigString("WooCommerceKey", ""), Utils.ReadConfigString("WooCommerceSecret", ""))
                };


                var request = new RestRequest("orders/" + orderID, Method.POST);
                request.AddHeader("Accept", "application/json");
                request.Parameters.Clear();
                request.RequestFormat = DataFormat.Json;
                OrderNote orderNote = new OrderNote() { note = note };
                request.AddJsonBody(orderNote);

                var response = client.Execute(request);

                if (response.StatusCode == System.Net.HttpStatusCode.Accepted ||
                    response.StatusCode == System.Net.HttpStatusCode.OK ||
                    response.StatusCode == System.Net.HttpStatusCode.Created)
                {
                    Utils.WriteLog("Order " + orderID.ToString() + " note added: '" + note + "'");
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
    }


    public class OrderUpdateHttpClient
    {
        private readonly HttpClient client;
        public OrderUpdateHttpClient()
        {
            HttpClientHandler handler = new HttpClientHandler();
            // {
            //      Proxy = new System.Net.WebProxy("http://127.0.0.1:8888"),
            //      UseProxy = false,
            //  };

            client = new HttpClient
            {
                Timeout = new TimeSpan(0, 0, 300),
                BaseAddress = new Uri(Utils.ReadConfigString("WooCommerceUrl", ""))
            };
            var authenticationBytes = Encoding.ASCII.GetBytes(Utils.ReadConfigString("WooCommerceKey", "") + ":" + Utils.ReadConfigString("WooCommerceSecret", ""));
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Authorization =
                    new System.Net.Http.Headers.AuthenticationHeaderValue("Basic", Convert.ToBase64String(authenticationBytes));
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

           
        }

         
        public  bool UpdateOrder(string orderID, string status1)
        {
            Utils.WriteLog("Requesting UpdateOrder()");
            try
            {
                OrderStatus orderStatus = new OrderStatus() { status = status1 };
                string postBody = JsonConvert.SerializeObject(orderStatus);

                Utils.WriteLog($"Putting CloseOrderAsync(orders/{orderID}");

                HttpResponseMessage response = client.PutAsync($"orders/{orderID}", new StringContent(postBody, Encoding.UTF8, "application/json")).Result;
                Utils.WriteLog("Response: " + response.StatusCode.ToString());
                if (response.IsSuccessStatusCode)
                {
                    string str =  response.Content.ReadAsStringAsync().Result;
                    Utils.WriteLog(str);
                    var settings = new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore,
                        MissingMemberHandling = MissingMemberHandling.Ignore
                    };
                    return true;
                }
                else
                    return false;
            }
            catch (HttpRequestException hre)
            {
                Utils.WriteLog("Error:" + hre.Message);
            }
            catch (TaskCanceledException)
            {
                Utils.WriteLog("Request canceled");
            }
            catch (Exception ex)
            {
                Utils.WriteLog("Exception:" + ex.Message);
            }
            finally
            {
                /*if (httpClient != null)
                {
                    httpClient.Dispose();
                    httpClient = null;
                }*/
            }
            return false;
        }

        public bool CreateOrderNote(string orderID, string note)
        {
            Utils.WriteLog("Requesting UpdateOrder()");
            try
            {
                OrderNote orderNote = new OrderNote() { note = note };
                string postBody = JsonConvert.SerializeObject(orderNote);

                Utils.WriteLog($" Posting note PostAsync(orders/{orderID}");

                HttpResponseMessage response = client.PostAsync($"orders/{orderID}", new StringContent(postBody, Encoding.UTF8, "application/json")).Result;
                Utils.WriteLog("Response: " + response.StatusCode.ToString());
                if (response.IsSuccessStatusCode)
                {
                    string str = response.Content.ReadAsStringAsync().Result;
                    Utils.WriteLog(str);
                    var settings = new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore,
                        MissingMemberHandling = MissingMemberHandling.Ignore
                    };
                    return true;
                }
                else
                    return false;
            }
            catch (HttpRequestException hre)
            {
                Utils.WriteLog("Error:" + hre.Message);
            }
            catch (TaskCanceledException)
            {
                Utils.WriteLog("Request canceled");
            }
            catch (Exception ex)
            {
                Utils.WriteLog("Exception:" + ex.Message);
            }
            finally
            {
                /*if (httpClient != null)
                {
                    httpClient.Dispose();
                    httpClient = null;
                }*/
            }
            return false;
        }


 

    }
}
