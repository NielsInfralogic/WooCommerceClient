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
        [DataMember(EmitDefaultValue = false)]
        public int? da { get; set; } = 0;

        [DataMember(EmitDefaultValue = false)]
        public int? en { get; set; } = 0;


    }
}
