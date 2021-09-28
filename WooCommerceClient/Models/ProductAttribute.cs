using System.Collections.Generic;
using System.Runtime.Serialization;

namespace WooCommerceClient.Models
{
    [DataContract]
    public class ProductAttribute : IWCObject
    {
        public static string Endpoint { get { return "products/attributes"; } }

        /// <summary>
        /// Unique identifier for the resource. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public int? id { get; set; }

        /// <summary>
        /// Attribute name. 
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
        /// Type of attribute. Options: select and text. Default is select.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string type { get; set; }

        /// <summary>
        /// Default sort order. Options: menu_order, name, name_num and id. Default is menu_order.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string order_by { get; set; }

        /// <summary>
        /// Enable/Disable attribute archives. Default is false.
        /// </summary>
        [DataMember(EmitDefaultValue = true)]
        public bool? has_archives { get; set; } = true;


        /// <summary>
        /// Language code e.g. 'en'
        /// </summary>
    //    [DataMember(EmitDefaultValue = true)]
    //    public string lang { get; set; } = "da";

        /// <summary>
        /// Point to master product ID
        /// </summary>
    //    [DataMember(EmitDefaultValue = false)]
     //   public Translations translations { get; set; }

    }

    [DataContract]
    public class ProductAttributeTerm : IWCObject
    {
        public static string Endpoint { get { return "terms"; } }

        /// <summary>
        /// Unique identifier for the resource. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public int? id { get; set; }

        /// <summary>
        /// Term name. 
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
        /// Menu order, used to custom sort the resource.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public int? menu_order { get; set; }

        /// <summary>
        /// Number of published products for the resource. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public int? count { get; set; }

        
        /// <summary>
        /// CUSTOM attribute: Attribute term id for country
        /// </summary>
       [DataMember(EmitDefaultValue = false)]
       public int? _country_id { get; set; }

        /// <summary>
        /// CUSTOM attribute: Attribute term id for country
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public int? _region_id { get; set; }

        /// <summary>
        /// Unique external identifier for the resource. 
        ///  
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public int visma_id { get; set; } = -1;

        /// <summary>
        /// Unique external identifier for the resource. 
        ///  
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public List<int> _related_producers { get; set; }

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
        ///[DataMember(EmitDefaultValue = false)]
        [IgnoreDataMember]
        public Translations translations { get; set; }
    }



}
