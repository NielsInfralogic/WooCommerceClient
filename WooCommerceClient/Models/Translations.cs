using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace WooCommerceClient.Models
{
    [DataContract]
    public class Translations
    {
        /// <summary>
        /// Reference to id of translation(s)
        /// </summary>
        [DataMember(EmitDefaultValue = true)]
        public int da { get; set; } = 0;

        [DataMember(EmitDefaultValue = true)]
        public int en { get; set; } = 0;

        [IgnoreDataMember]
        public string nameforda { get; set; }

        [IgnoreDataMember]
        public string nameforen { get; set; }

    }
}
