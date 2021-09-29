using System.Runtime.Serialization;

namespace WooCommerceClient.Models
{
    [DataContract]
    public class ProductTag : IWCObject
    {
        public static string Endpoint { get { return "products/tags"; } }

        /// <summary>
        /// Unique identifier for the resource. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public int? id { get; set; }

        /// <summary>
        /// Tag name. 
        /// mandatory
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string name { get; set; }

        /// <summary>
        /// An alphanumeric identifier for the resource unique to its type.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string slug { get; set; }

        /// <summary>
        /// HTML description of the resource.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string description { get; set; }

        /// <summary>
        /// Number of published products for the resource. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public int? count { get; set; }

        /// <summary>
        /// Language code e.g. 'en'
        /// </summary>
        [DataMember(EmitDefaultValue = true)]
        public string lang { get; set; } = "da";

        /// <summary>
        /// Language code e.g. 'en'
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string translation_of { get; set; } = null;


        /// <summary>
        /// Point to master product ID
        /// </summary>
        /// [DataMember(EmitDefaultValue = false)]
     //   [IgnoreDataMember]
      //  public Translations translations { get; set; }

 

    }
}
