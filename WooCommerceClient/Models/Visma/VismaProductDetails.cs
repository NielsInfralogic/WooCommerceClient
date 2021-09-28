using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WooCommerceClient.Models.Visma
{
    public class VismaProductDetails
    {
        public string Type { get; set; } = "";
        public string Year { get; set; } = "";
        public string Organic { get; set; } = "";
        public string Biodynamic { get; set; } = "";
        public string Naturewine { get; set; } = "";

        public decimal Weight { get; set; } = 0.0M;
        public decimal Content { get; set; } = 0.0M;
        public decimal Alcohol { get; set; } = 0.0M;

        public int UnitsPerPackage { get; set; } = 1;

        public string Country { get; set; } = "";
        public string District { get; set; } = "";
        public string County { get; set; } = "";
        public string Producer { get; set; } = "";
        public string Classification { get; set; } = "";
        public string Mark { get; set; } = "";

        public string SubType { get; set; } = "";

        public int UnitsPerPallet { get; set; } = 0;

        public string Lagringsform { get; set; } = "";
        public string Destillationsmetode { get; set; } = "";
        public string Varegruppe { get; set; } = "";

        public string Category { get; set; } = "";
        public string Appellation { get; set; } = "";

        public string Omraade { get; set; } = "";

        public string Scores { get; set; } = "";

        // Intersport 
        public string Color { get; set; } = "";
        public string Size { get; set; } = "";
        public string Model { get; set; } = "";

        public string Division { get; set; } = "";
        public string Brand { get; set; } = "";
        public string Group { get; set; } = "";

    }

}
