using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WooCommerceClient.Models.Visma
{
    public class VismaProducer
    {
        public int LangNo { get; set; } = 45;
        public string Producer { get; set; } = "";
        public string ProducerRegion { get; set; } = "";
        public string ProducerCountry { get; set; } = "";

        public int VismaID { get; set; } = 0;

        public bool InUse { get; set; } = true;

        public List<int> RelatedProducers { get; set; }

        public VismaProducer()
        {
            RelatedProducers = new List<int>();
        }
    }
}
