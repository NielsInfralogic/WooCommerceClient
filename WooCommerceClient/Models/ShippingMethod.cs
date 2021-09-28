using System.Runtime.Serialization;

namespace WooCommerceClient.Models
{
    [DataContract]
    public class ShippingMethod
    {
        public static string Endpoint { get { return "shipping_methods"; } }

        /// <summary>
        /// Method ID. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string id { get; set; }

        /// <summary>
        /// Shipping method title. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string title { get; set; }

        /// <summary>
        /// Shipping method description.
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string description { get; set; }

    }
}
