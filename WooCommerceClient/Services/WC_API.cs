using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WooCommerceClient.Models;

namespace WooCommerceClient.Services
{
    class WC_API
    {
        private readonly string defaultRestAPIUrl = Utils.ReadConfigString("WooCommerceUrl", "");
        private readonly string defaultAPIKey = Utils.ReadConfigString("WooCommerceKey", "");
        private readonly string defaultAPISecret = Utils.ReadConfigString("WooCommerceSecret", "");
        private readonly WCObject wc;

        public WC_API(bool? debug = false)
        {
            MyRestAPI rest = new MyRestAPI(defaultRestAPIUrl, defaultAPIKey, defaultAPISecret);
            wc = new WCObject(rest);
            if (debug.HasValue)
                rest.Debug = debug.Value;
        }

        public async Task<T> Add<T>(T t, int? parentId = default(int?)) where T : IWCObject
        {
            return await Add<T>(t, "", parentId);
        }
        public async Task<T> Add<T>(T t, string language, int? parentId = default(int?)) where T : IWCObject
        {
            T newObject = default(T);
            try
            {
                Dictionary<string, string> parameters = string.IsNullOrEmpty(language) ? null : new Dictionary<string, string>() { { "lang", language } }; ;

                if ((t is ProductAttributeTerm) && parentId.HasValue)
                    newObject = await wc.Attribute.Terms.Add(t, parentId.Value, parameters);
                else if ((t is ProductAttribute))
                    newObject = await wc.Attribute.Add(t, parameters);
                else if ((t is Product))
                    newObject = await wc.Product.Add(t, parameters);
            }
            catch (Exception ex)
            {
                Utils.WriteLog($"Error : wc.{typeof(T).Name}.Add()", ex);
            }

            if (newObject == null)
            {
                Utils.WriteLog($"Error : wc.{typeof(T).Name}.Add() returned null");
            }
            else
            {
                t.id = newObject.id.Value;
            }

            return newObject;
        }

        public async Task<T> Update<T>(T t, int? parentId = default(int?)) where T : IWCObject
        {
            return await Update<T>(t, "", parentId);
        }

        public async Task<T> Update<T>(T t, string language, int? parentId = default(int?)) where T : IWCObject
        {
            T newObject = default(T);
            try
            {
                Dictionary<string, string> parameters = string.IsNullOrEmpty(language) ? null : new Dictionary<string, string>() { { "lang", language } }; ;

                if ((t is ProductAttributeTerm) && t.id.HasValue && parentId.HasValue)
                    newObject = await wc.Attribute.Terms.Update(t.id.Value, t, parentId.Value, parameters);
                else if ((t is ProductAttribute) && t.id.HasValue)
                    newObject = await wc.Attribute.Update(t.id.Value, t, parameters);
                else if ((t is Product) && t.id.HasValue)
                {
                    newObject = await wc.Product.Update(t.id.Value, t, parameters);
                }
            }
            catch (Exception ex)
            {
                Utils.WriteLog($"Error : wc.{typeof(T).Name}.Update()", ex);
            }

            if (newObject == null)
            {
                Utils.WriteLog($"Error : wc.{typeof(T).Name}.Update() returned null");
            }

            return newObject;
        }

        internal async Task<object> Delete<T>(int idValue, bool force) where T : IWCObject
        {
            return await Delete<T>(idValue, "", force);
        }

        internal async Task<object> Delete<T>(int idValue, string language, bool force) where T : IWCObject
        {
            if (typeof(T) == typeof(Product))
            {
                Dictionary<string, string> parameters = string.IsNullOrEmpty(language) ? null : new Dictionary<string, string>() { { "lang", language } }; ;
                return await wc.Product.Delete(idValue, force, parameters);
            }

            return null;
        }
    }
}
