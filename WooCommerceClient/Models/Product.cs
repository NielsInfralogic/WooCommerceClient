using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using WooCommerceClient.Models.Base;

namespace WooCommerceClient.Models
{
    public class ProductBatch : BatchObject<Product> { }

    [DataContract]
    public class Product : JsonObject, IWCObject
    {
        public static string Endpoint { get { return "products"; } }

        /// <summary>
        /// Unique identifier for the resource. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public int? id { get; set; }

        /// <summary>
        /// Product name.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string name { get; set; }

        /// <summary>
        /// Product slug.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string slug { get; set; }

        /// <summary>
        /// Product URL. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string permalink { get; set; }

        /// <summary>
        /// The date the product was created, in the site’s timezone. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public DateTime? date_created { get; set; }

        /// <summary>
        /// The date the product was created, as GMT. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public DateTime? date_created_gmt { get; set; }

        /// <summary>
        /// The date the product was last modified, in the site’s timezone. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public DateTime? date_modified { get; set; }

        /// <summary>
        /// The date the product was last modified, as GMT. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public DateTime? date_modified_gmt { get; set; }

        /// <summary>
        /// Product type. Options: simple, grouped, external and variable. Default is simple.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string type { get; set; }

        /// <summary>
        /// Product status (post status). Options: draft, pending, private and publish. Default is publish.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string status { get; set; }

        /// <summary>
        /// Featured product. Default is false.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public bool? featured { get; set; }

        /// <summary>
        /// Catalog visibility. Options: visible, catalog, search and hidden. Default is visible.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string catalog_visibility { get; set; }

        /// <summary>
        /// Product description.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string description { get; set; }

        /// <summary>
        /// Enable HTML for product description 
        /// write-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public bool? enable_html_description { get; set; }

        /// <summary>
        /// Product short description.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string short_description { get; set; }

        /// <summary>
        /// Enable HTML for product short description 
        /// write-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string enable_html_short_description { get; set; }

        /// <summary>
        /// Unique identifier.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string sku { get; set; }
        
        [DataMember(EmitDefaultValue = false, Name = "price")]
        protected object priceValue { get; set; }
        /// <summary>
        /// Current product price. 
        /// read-only
        /// </summary>
        public decimal? price { get; set; }
        
        [DataMember(EmitDefaultValue = false, Name = "regular_price")]
        protected object regular_priceValue { get; set; }
        /// <summary>
        /// Product regular price.
        /// </summary>
        public decimal? regular_price { get; set; }
        
        [DataMember(EmitDefaultValue = false, Name = "sale_price")]
        protected object sale_priceValue { get; set; }
        /// <summary>
        /// Product sale price.
        /// </summary>
        public decimal? sale_price { get; set; }

        /// <summary>
        /// Start date of sale price, in the site’s timezone.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public DateTime? date_on_sale_from { get; set; }

        /// <summary>
        /// Start date of sale price, as GMT.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public DateTime? date_on_sale_from_gmt { get; set; }

        /// <summary>
        /// End date of sale price, in the site’s timezone.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public DateTime? date_on_sale_to { get; set; }

        /// <summary>
        /// End date of sale price, in the site’s timezone.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public DateTime? date_on_sale_to_gmt { get; set; }

        /// <summary>
        /// Price formatted in HTML. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string price_html { get; set; }

        /// <summary>
        /// Shows if the product is on sale. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public bool? on_sale { get; set; }

        /// <summary>
        /// Shows if the product can be bought. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public bool? purchasable { get; set; }

        /// <summary>
        /// Amount of sales. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public long? total_sales { get; set; }

        /// <summary>
        /// If the product is virtual. Default is false.
        /// </summary>
        [DataMember(EmitDefaultValue = false, Name = "virtual")]
        public bool? _virtual { get; set; }

        /// <summary>
        /// If the product is downloadable. Default is false.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public bool? downloadable { get; set; }

        /// <summary>
        /// List of downloadable files. See Product - Downloads properties
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public List<ProductDownloadLine> downloads { get; set; }

        /// <summary>
        /// Number of times downloadable files can be downloaded after purchase. Default is -1.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public int? download_limit { get; set; }

        /// <summary>
        /// Number of days until access to downloadable files expires. Default is -1.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public int? download_expiry { get; set; }

        /// <summary>
        /// Product external URL. Only for external products.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string external_url { get; set; }

        /// <summary>
        /// Product external button text. Only for external products.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string button_text { get; set; }

        /// <summary>
        /// Tax status. Options: taxable, shipping and none. Default is taxable.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string tax_status { get; set; }

        /// <summary>
        /// Tax class.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string tax_class { get; set; }

        /// <summary>
        /// Stock management at product level. Default is false.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public bool? manage_stock { get; set; }

        [DataMember(EmitDefaultValue = false, Name = "stock_quantity")]
        protected object stock_quantityValue { get; set; }
        /// <summary>
        /// Stock quantity.
        /// </summary>
        public int? stock_quantity { get; set; }

        /// <summary>
        /// Controls the stock status of the product. Options: instock, outofstock, onbackorder. Default is instock.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string stock_status { get; set; }

        /// <summary>
        /// If managing stock, this controls if backorders are allowed. Options: no, notify and yes. Default is no.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string backorders { get; set; }

        /// <summary>
        /// Shows if backorders are allowed. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public bool? backorders_allowed { get; set; }

        /// <summary>
        /// Shows if the product is on backordered. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public bool? backordered { get; set; }

        /// <summary>
        /// Allow one item to be bought in a single order. Default is false.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public bool? sold_individually { get; set; }

        [DataMember(EmitDefaultValue = false, Name = "weight")]
        protected object weightValue { get; set; }
        /// <summary>
        /// Product weight (kg).
        /// </summary>
        public decimal? weight { get; set; }

        /// <summary>
        /// Product dimensions. See Product - Dimensions properties
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public ProductDimension dimensions { get; set; }

        /// <summary>
        /// Shows if the product need to be shipped. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public bool? shipping_required { get; set; }

        /// <summary>
        /// Shows whether or not the product shipping is taxable. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public bool? shipping_taxable { get; set; }

        /// <summary>
        /// Shipping class slug.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string shipping_class { get; set; }

        /// <summary>
        /// Shipping class ID. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string shipping_class_id { get; set; }

        /// <summary>
        /// Allow reviews. Default is true.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public bool? reviews_allowed { get; set; }

        /// <summary>
        /// Reviews average rating. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string average_rating { get; set; }

        /// <summary>
        /// Amount of reviews that the product have. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public int? rating_count { get; set; }

        /// <summary>
        /// List of related products IDs. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public List<int> related_ids { get; set; }

        /// <summary>
        /// List of up-sell products IDs.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
       //  public List<int> upsell_ids { get; set; }
         public object upsell_ids { get; set; }

        /// <summary>
        /// List of cross-sell products IDs.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public List<int> cross_sell_ids { get; set; }

        /// <summary>
        /// Product parent ID.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public int? parent_id { get; set; }

        /// <summary>
        /// Optional note to send the customer after purchase.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string purchase_note { get; set; }

        /// <summary>
        /// List of categories. See Product - Categories properties
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public List<ProductCategoryLine> categories { get; set; }

        /// <summary>
        /// List of tags. See Product - Tags properties
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public List<ProductTagLine> tags { get; set; }

        /// <summary>
        /// List of images. See Product - Images properties
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public List<ProductImage> images { get; set; }

        /// <summary>
        /// List of attributes. See Product - Attributes properties
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public List<ProductAttributeLine> attributes { get; set; }

        /// <summary>
        /// Defaults variation attributes. See Product - Default attributes properties
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public List<ProductDefaultAttribute> default_attributes { get; set; }

        /// <summary>
        /// List of variations IDs. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public List<int> variations { get; set; }

        /// <summary>
        /// List of grouped products ID. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public List<int> grouped_products { get; set; }

        /// <summary>
        /// Menu order, used to custom sort products.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public int? menu_order { get; set; }

        /// <summary>
        /// Meta data. See Product - Meta data properties
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public List<ProductMeta> meta_data { get; set; } = null;

        /// <summary>
        /// Language code e.g. 'en'
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string lang { get; set; } = "da";

        /// <summary>
        /// Language code e.g. 'en'
        /// </summary>
        [DataMember(EmitDefaultValue = true)]
        public string translation_of { get; set; } = null;
        

       
        /// <summary>
        /// Point to master product ID
        /// </summary>
        //[DataMember(EmitDefaultValue = false)]
       // [IgnoreDataMember]
     //   public Translations translations { get; set; }

        [IgnoreDataMember]
        public int vismaUnitsPerProdNo { get; set; }

        [IgnoreDataMember]
        public List<string> vismaRelatedProduct { get; set; }

        [IgnoreDataMember]
        public decimal vismaPriceEUR { get; set; } = 0.0M;

        // Deep copy - old style..
        public void Clone(Product copyFrom)
        {
            id = copyFrom.id;
            name = copyFrom.name;
            //slug = copyFrom.slug; //20210923 don't send slug to webshop
            permalink = copyFrom.permalink;
            date_created = copyFrom.date_created;
            date_created_gmt = copyFrom.date_created_gmt;
            date_modified = copyFrom.date_modified;
            date_modified_gmt = copyFrom.date_modified_gmt;
            type = copyFrom.type;
            status = copyFrom.status;
            featured = copyFrom.featured;
            catalog_visibility = copyFrom.catalog_visibility;
            description = copyFrom.description;
            enable_html_description = copyFrom.enable_html_description;
            short_description = copyFrom.short_description;
            enable_html_short_description = copyFrom.enable_html_short_description;
            sku = copyFrom.sku;
            price = copyFrom.price;
            regular_price = copyFrom.regular_price;
            sale_price = copyFrom.sale_price;
            date_on_sale_from = copyFrom.date_on_sale_from;
            date_on_sale_from_gmt = copyFrom.date_on_sale_from_gmt;
            date_on_sale_to = copyFrom.date_on_sale_to;
            date_on_sale_to_gmt = copyFrom.date_on_sale_to_gmt;
            price_html = copyFrom.price_html;
            on_sale = copyFrom.on_sale;
            purchasable = copyFrom.purchasable;
            total_sales = copyFrom.total_sales;
            _virtual = copyFrom._virtual;
            downloadable = copyFrom.downloadable;
            downloads = null;
            if (copyFrom.downloads != null)
            {
                downloads = new List<ProductDownloadLine>();
                foreach (ProductDownloadLine productDownloadLine in copyFrom.downloads)
                    downloads.Add(new ProductDownloadLine() { file = productDownloadLine.file, name = productDownloadLine.name, id = productDownloadLine.id });
            }  
               
            download_limit = copyFrom.download_limit;
            download_expiry = copyFrom.download_expiry;
            external_url = copyFrom.external_url;
            button_text = copyFrom.button_text;
            tax_status = copyFrom.tax_status;
            tax_class = copyFrom.tax_class;
            manage_stock = copyFrom.manage_stock;
            stock_quantity = copyFrom.stock_quantity;
            stock_status = copyFrom.stock_status;
            backorders = copyFrom.backorders;
            backorders_allowed = copyFrom.backorders_allowed;
            backordered = copyFrom.backordered;
            sold_individually = copyFrom.sold_individually;
            weight = copyFrom.weight;
            dimensions = null;
            if (copyFrom.dimensions != null)
                dimensions = new ProductDimension() { height = copyFrom.dimensions.height, length = copyFrom.dimensions.length, width = copyFrom.dimensions.width };
            shipping_required = copyFrom.shipping_required;
            shipping_taxable = copyFrom.shipping_taxable;
            shipping_class = copyFrom.shipping_class;
            shipping_class_id = copyFrom.shipping_class_id;
            reviews_allowed = copyFrom.reviews_allowed;
            average_rating = copyFrom.average_rating;
            rating_count = copyFrom.rating_count;
            related_ids = null;
            if (copyFrom.related_ids != null)
            {
                related_ids = new List<int>();
                foreach (int i in copyFrom.related_ids)
                    related_ids.Add(i);
            }
            upsell_ids = null;
            if (copyFrom.upsell_ids != null)
            {
                upsell_ids = new List<int>();
                foreach (int i in copyFrom.upsell_ids as List<int>)
                    (upsell_ids as List<int>).Add(i);
            }
            cross_sell_ids =null;
            if (copyFrom.cross_sell_ids != null)
            {
                cross_sell_ids = new List<int>();
                foreach (int i in copyFrom.cross_sell_ids)
                    cross_sell_ids.Add(i);
            }
            parent_id = copyFrom.parent_id;
            purchase_note = copyFrom.purchase_note;
            categories = null;
            if (copyFrom.categories != null)
            {
                categories = new List<ProductCategoryLine>();
                foreach (ProductCategoryLine line in copyFrom.categories)
                    categories.Add(new ProductCategoryLine() { id = line.id, name = line.name  /*, slug = line.slug */ });//20210923 don't send slug to webshop
            }

            tags = null;
            if (copyFrom.tags != null) 
            {
                tags = new List<ProductTagLine>();
                foreach (ProductTagLine line in copyFrom.tags)
                    tags.Add(new ProductTagLine() { id = line.id, name = line.name/*, slug = line.slug*/ });//20210923 don't send slug to webshop
            }

            images = null;
            if (copyFrom.images != null)
            {
                images = new List<ProductImage>();
                foreach (ProductImage image in copyFrom.images)
                    images.Add(new ProductImage() { alt = image.alt, date_created = image.date_created, date_created_gmt = image.date_created_gmt, date_modified = image.date_modified, date_modified_gmt = image.date_modified_gmt, id = image.id, name = image.name, position = image.position, src = image.src });
            }
            attributes = null;
            if (copyFrom.attributes != null)
            {
                attributes = new List<ProductAttributeLine>();
                foreach (ProductAttributeLine line in copyFrom.attributes)
                {
                    ProductAttributeLine newLine = new ProductAttributeLine() { id = line.id, name = line.name, position = line.position, variation = line.variation, visible = line.visible, options = null };
                    if (line.options != null)
                    {
                        newLine.options = new List<string>();
                        foreach (string option in line.options)
                            newLine.options.Add(option);
                    }
                    attributes.Add(newLine);
                }
            }
            default_attributes = null;
            if (copyFrom.default_attributes != null)
            {
                default_attributes = new List<ProductDefaultAttribute>();
                foreach (ProductDefaultAttribute attribute in copyFrom.default_attributes)
                    default_attributes.Add(new ProductDefaultAttribute() { id = attribute.id, name = attribute.name, option = attribute.option });
            }
            variations = null;
            if (copyFrom.variations != null)
            {
                variations = new List<int>();
                foreach (int variation in copyFrom.variations)
                    variations.Add(variation);
            }

            grouped_products = null;
            if (copyFrom.grouped_products != null)
            {
                grouped_products = new List<int>();
                foreach (int grouped_product in copyFrom.grouped_products)
                    grouped_products.Add(grouped_product);
            }
            menu_order = copyFrom.menu_order;

            meta_data = null;

            if (copyFrom.meta_data != null)
            {
                meta_data = new List<ProductMeta>();
                foreach (ProductMeta meta in copyFrom.meta_data)
                    meta_data.Add(new ProductMeta() { id = meta.id, key = meta.key, value = meta.value });
            }
            lang = copyFrom.lang;

            //translations = null;
            //if (copyFrom.translations != null)
           // {
                //translations = new Translations() { da = copyFrom.translations.da, en = copyFrom.translations.en, nameforda = copyFrom.translations.nameforda, nameforen = copyFrom.translations.nameforen };
            //
            //}

            vismaUnitsPerProdNo = copyFrom.vismaUnitsPerProdNo;

        }
    }

    [DataContract]
    public class ProductMeta : WCObject.MetaData
    {

    }

    public class ProductBomItem
    {
        public string title { get; set; }
        public string sku { get; set; }
        public int quantity { get; set; }
    }

    [DataContract]
    public class ProductDownloadLine
    {
        /// <summary>
        /// File MD5 hash. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string id { get; set; }

        /// <summary>
        /// File name.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string name { get; set; }

        /// <summary>
        /// File URL.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string file { get; set; }

    }



    [DataContract]
    public class ProductDimension
    {
        /// <summary>
        /// Product length (cm).
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string length { get; set; }

        /// <summary>
        /// Product width (cm).
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string width { get; set; }

        /// <summary>
        /// Product height (cm).
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string height { get; set; }
    }

    [DataContract]
    public class ProductCategoryLine : Islug
    {
        /// <summary>
        /// Category ID.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public int? id { get; set; }

        /// <summary>
        /// Category name. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string name { get; set; }

        /// <summary>
        /// Category slug. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string slug { get; set; }

    }

    [DataContract]
    public class ProductTagLine : Islug
    {
        /// <summary>
        /// Tag ID.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public int? id { get; set; }

        /// <summary>
        /// Tag name. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string name { get; set; }

        /// <summary>
        /// Tag slug. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string slug { get; set; }

    }

    [DataContract]
    public class ProductImage
    {
        /// <summary>
        /// Image ID.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public long? id { get; set; }

        /// <summary>
        /// The date the image was created, in the site’s timezone. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public DateTime? date_created { get; set; }

        /// <summary>
        /// The date the image was created, as GMT. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public DateTime? date_created_gmt { get; set; }

        /// <summary>
        /// The date the image was last modified, in the site’s timezone. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public DateTime? date_modified { get; set; }

        /// <summary>
        /// The date the image was last modified, as GMT. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public DateTime? date_modified_gmt { get; set; }

        /// <summary>
        /// Image URL.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string src { get; set; }

        /// <summary>
        /// Image name.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string name { get; set; }

        /// <summary>
        /// Image alternative text.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string alt { get; set; }

        /// <summary>
        /// Image position. 0 means that the image is featured.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public int? position { get; set; }
    }

    [DataContract]
    public class ProductAttributeLine
    {
        /// <summary>
        /// Attribute ID.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public int? id { get; set; }

        /// <summary>
        /// Attribute name.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string name { get; set; }

        /// <summary>
        /// Attribute position.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public int? position { get; set; }

        /// <summary>
        /// Define if the attribute is visible on the “Additional information” tab in the product’s page. Default is false.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public bool? visible { get; set; }

        /// <summary>
        /// Define if the attribute can be used as variation. Default is false.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public bool? variation { get; set; }

        /// <summary>
        /// List of available term names of the attribute.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public List<string> options { get; set; }

    }

    [DataContract]
    public class ProductDefaultAttribute
    {
        /// <summary>
        /// Attribute ID.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public int? id { get; set; }

        /// <summary>
        /// Attribute name.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string name { get; set; }

        /// <summary>
        /// Selected attribute term name.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string option { get; set; }

    }

    [DataContract]
    public class ProductReview
    {
        public static string Endpoint { get { return "products/reviews"; } }

        /// <summary>
        /// Unique identifier for the resource. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public int? id { get; set; }

        /// <summary>
        /// The date the review was created, in the site’s timezone.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public DateTime? date_created { get; set; }

        /// <summary>
        /// The date the review was created, as GMT.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public DateTime? date_created_gmt { get; set; }

        /// <summary>
        /// Unique identifier for the product that the review belongs to.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public int? product_id { get; set; }

        /// <summary>
        /// Status of the review. Options: approved, hold, spam, unspam, transh and untrash. Defauls to approved.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string status { get; set; }

        /// <summary>
        /// Reviewer name. 
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string reviewer { get; set; }

        /// <summary>
        /// Reviewer email. 
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string reviewer_email { get; set; }

        /// <summary>
        /// The content of the review. 
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string review { get; set; }

        /// <summary>
        /// Review rating (0 to 5).
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public int? rating { get; set; }

        /// <summary>
        /// Shows if the reviewer bought the product or not. 
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public bool? verified { get; set; }
    }


   

    [DataContract]
    public class ProductMultiLang : JsonObject
    {
        public static string Endpoint { get { return "products"; } }

        /// <summary>
        /// Unique identifier for the resource. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public int? id { get; set; }

        /// <summary>
        /// Product name.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string name { get; set; }

       

        /// <summary>
        /// The date the product was created, in the site’s timezone. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public DateTime? date_created { get; set; }

        /// <summary>
        /// The date the product was created, as GMT. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public DateTime? date_created_gmt { get; set; }

        /// <summary>
        /// The date the product was last modified, in the site’s timezone. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public DateTime? date_modified { get; set; }

        /// <summary>
        /// The date the product was last modified, as GMT. 
        /// read-only
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public DateTime? date_modified_gmt { get; set; }

        /// <summary>
        /// Product type. Options: simple, grouped, external and variable. Default is simple.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string type { get; set; }

        /// <summary>
        /// Product description.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string description { get; set; }


        /// <summary>
        /// Language code e.g. 'en'
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string lang { get; set; }

        /// <summary>
        /// Point to master product ID
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string translation_of { get; set; }

        /// <summary>
        /// Product short description.
        /// </summary>
        [DataMember(EmitDefaultValue = false)]
        public string short_description { get; set; }

        
        [DataMember(EmitDefaultValue = false, Name = "regular_price")]
        protected object regular_priceValue { get; set; }
        /// <summary>
        /// Product regular price.
        /// </summary>
        public decimal? regular_price { get; set; }




    }


    public class ProductMetaEur
    {
        public string EUR { get; set; } = "";
    }
}


