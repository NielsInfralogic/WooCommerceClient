using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WooCommerceClient.Models;

namespace WooCommerceClient.Services
{
    class Categories
    {
        internal static void Syns()
        {
            List<ProductCategory> categories = new List<ProductCategory>();
            // categories.Add(new ProductCategory() { vismaParentCatNo = 0, parent = null, name = Constants.NoDiscount, slug = Utils.SanitizeSlugNameNew(Constants.NoDiscount), lang = "da", translations = new Translations() { nameforda = Constants.NoDiscount, nameforen = Constants.NoDiscountEn } });
            categories.Add(new ProductCategory()
            {
                vismaParentCatNo = 0,
                parent = null,
                name = Constants.NoDiscount,
                name_da = Constants.NoDiscount,
                lang = "da"
            });
            //categories.Add(new ProductCategory() { vismaParentCatNo = 0, parent = null, name = Constants.NoDiscountEn, slug = Utils.SanitizeSlugNameNew(Constants.NoDiscount) + "-en", lang = "en", translations = new Translations() { nameforda = Constants.NoDiscount, nameforen = Constants.NoDiscountEn } });
            categories.Add(new ProductCategory()
            {
                vismaParentCatNo = 0,
                parent = null,
                name = Constants.NoDiscountEn,
                name_da = Constants.NoDiscount,
                lang = "en"
            });
            //categories.Add(new ProductCategory() { vismaParentCatNo = 0, parent = null, name = Constants.CategoryOekologisk, slug = Utils.SanitizeSlugNameNew(Constants.CategoryOekologiskSlug), lang = "da", translations = new Translations() { nameforda = Constants.CategoryOekologisk, nameforen = Constants.CategoryOekologiskEn } });
            categories.Add(new ProductCategory()
            {
                vismaParentCatNo = 0,
                parent = null,
                name = Constants.CategoryOekologisk,
                name_da = Constants.CategoryOekologisk,
                lang = "da"
            });
            //categories.Add(new ProductCategory() { vismaParentCatNo = 0, parent = null, name = Constants.CategoryOekologiskEn, slug = Utils.SanitizeSlugNameNew(Constants.CategoryOekologiskSlug) + "-en", lang = "en", translations = new Translations() { nameforda = Constants.CategoryOekologisk, nameforen = Constants.CategoryOekologiskEn } });
            categories.Add(new ProductCategory()
            {
                vismaParentCatNo = 0,
                parent = null,
                name = Constants.CategoryOekologiskEn,
                name_da = Constants.CategoryOekologisk,
                lang = "en"
            });
            //categories.Add(new ProductCategory() { vismaParentCatNo = 0, parent = null, name = Constants.ProductWithBom, slug = Utils.SanitizeSlugNameNew(Constants.ProductWithBom), lang = "da", translations = new Translations() { nameforda = Constants.ProductWithBom, nameforen = Constants.ProductWithBomEn } });
            categories.Add(new ProductCategory()
            {
                vismaParentCatNo = 0,
                parent = null,
                name = Constants.ProductWithBom,
                name_da = Constants.ProductWithBom,
                lang = "da"
            });
            //categories.Add(new ProductCategory() { vismaParentCatNo = 0, parent = null, name = Constants.ProductWithBomEn, slug = Utils.SanitizeSlugNameNew(Constants.ProductWithBom) + "-en", lang = "en", translations = new Translations() { nameforda = Constants.ProductWithBom, nameforen = Constants.ProductWithBomEn } });
            categories.Add(new ProductCategory()
            {
                vismaParentCatNo = 0,
                parent = null,
                name = Constants.ProductWithBomEn,
                name_da = Constants.ProductWithBom,
                lang = "en"
            });

            //if (db.GetCategories(ref categories, out errmsg) == false)
            //    Utils.WriteLog("ERROR db.GetCategories() - " + errmsg);

            if (categories.Count > 0)
            {
                bool res = SyncCategories(categories).Result;
            }
        }

        private static async Task<bool> SyncCategories(List<ProductCategory> categories)
        {
            try
            {
                MyRestAPI rest = new MyRestAPI(Utils.ReadConfigString("WooCommerceUrl", ""),
                         Utils.ReadConfigString("WooCommerceKey", ""),
                         Utils.ReadConfigString("WooCommerceSecret", ""));
                WCObject wc = new WCObject(rest);

                List<ProductCategory> wooProductCategories = WooCommerceHelpers.GetWooCommerceCategories().Result;

                // Level 0 categories

                List<ProductCategory> level0categories = categories.FindAll(p => p.vismaParentCatNo == 0);

                foreach (ProductCategory category in level0categories)
                {
                    ProductCategory existingCategory = wooProductCategories.FirstOrDefault(p =>  p.name == category.name && p.lang == category.lang);
                    ProductCategory newUpdatedCategory = null;
                    if (existingCategory == null)
                    {
                        try
                        {
                            Utils.WriteLog($"Adding category  {category.name}  ");
                            newUpdatedCategory = await wc.Category.Add(category, new Dictionary<string, string>() { { "lang", category.lang } });
                        }
                        catch (Exception ex)
                        {
                            Utils.WriteLog($"Error : wc.Category.Add() - {ex.Message} {category.name}  ");
                            return false;
                        }
                        if (newUpdatedCategory == null)
                        {
                            Utils.WriteLog($"Error : wc.Category.Add() returned null");
                            return false;
                        }
                        category.id = newUpdatedCategory.id;
                    }
                    else
                    {
                        try
                        {
                            category.id = existingCategory.id;
                            if (category.lang == "en")
                            {
                                ProductCategory category_da = wooProductCategories.FirstOrDefault(p => p.lang == "da" && p.name  == category.name_da);
                                if (category_da != null)
                                    category.translation_of = category_da.id.Value.ToString();
                            }

                            Utils.WriteLog($"Updating category  {category.name}  ");
                            newUpdatedCategory = await wc.Category.Update(category.id.Value, category, new Dictionary<string, string>() { { "lang", category.lang } });
                        }
                        catch (Exception ex)
                        {
                            Utils.WriteLog($"Error : wc.Category.Update() - {ex.Message} - {category.id} . {ex.InnerException} {ex.StackTrace}");
                            return false;
                        }
                        if (newUpdatedCategory == null)
                        {
                            Utils.WriteLog($"Error : wc.Category.Update() returned null");
                            return false;
                        }
                    }


                }
                /*         wooProductCategories = WooCommerceHelpers.GetWooCommerceCategories().Result;

                         // Level 1 categories

                         List<ProductCategory> level1categories = new List<ProductCategory>();
                         foreach (ProductCategory level0categorie in level0categories)
                         {
                             List<ProductCategory> l1c = categories.FindAll(p => p.vismaParentCatNo == level0categorie.vismaCatNo);
                             foreach (ProductCategory c in l1c)
                             {
                                 ProductCategory wooCat0 = wooProductCategories.FirstOrDefault(p => p.slug == level0categorie.slug || p.name == level0categorie.name);
                                 if (wooCat0 != null)
                                     c.parent = wooCat0.id;
                                 level1categories.Add(c);
                             }
                         }

                         foreach (ProductCategory category in level1categories)
                         {
                             ProductCategory existingCategory = wooProductCategories.FirstOrDefault(p => (p.slug == category.slug || p.name == category.name) && p.lang == category.lang);
                             ProductCategory newUpdatedCategory = null;
                             if (existingCategory == null)
                             {
                                 try
                                 {
                                     newUpdatedCategory = await wc.Category.Add(category);
                                 }
                                 catch (Exception ex)
                                 {
                                     Utils.WriteLog($"Error : wc.Category.Add() - {ex.Message}");
                                     return false;
                                 }
                                 if (newUpdatedCategory == null)
                                 {
                                     Utils.WriteLog($"Error : wc.Category.Add() returned null");
                                     return false;
                                 }
                                 category.id = newUpdatedCategory.id;
                             }
                             else
                             {
                                 try
                                 {
                                     ProductCategory existingCategory_en = wooProductCategories.FirstOrDefault(p => (p.slug == category.slug || p.name == category.name) && p.lang == "en");
                                     ProductCategory existingCategory_da = wooProductCategories.FirstOrDefault(p => (p.slug == category.slug || p.name == category.name) && p.lang == "da");

                                     category.id = existingCategory.id;
                                     category.translations.da = existingCategory_da != null ? existingCategory_da.id.Value : existingCategory.translations.da;
                                     category.translations.en = existingCategory_en != null ? existingCategory_en.id.Value : existingCategory.translations.en;

                                     newUpdatedCategory = await wc.Category.Update(category.id.Value, category);
                                 }
                                 catch (Exception ex)
                                 {
                                     Utils.WriteLog($"Error : wc.Category.Add() - {ex.Message}");
                                     return false;
                                 }
                                 if (newUpdatedCategory == null)
                                 {
                                     Utils.WriteLog($"Error : wc.Category.Add() returned null");
                                     return false;
                                 }
                             }
                         }

                         wooProductCategories = WooCommerceHelpers.GetWooCommerceCategories().Result;

                         // Level 2 categories

                         List<ProductCategory> level2categories = new List<ProductCategory>();
                         foreach (ProductCategory level1categorie in level0categories)
                         {
                             List<ProductCategory> l2c = categories.FindAll(p => p.vismaParentCatNo == level1categorie.vismaCatNo);
                             foreach (ProductCategory c in l2c)
                             {
                                 ProductCategory wooCat1 = wooProductCategories.FirstOrDefault(p => (p.slug == level1categorie.slug || p.name == level1categorie.name) && p.lang == level1categorie.lang);
                                 if (wooCat1 != null)
                                     c.parent = wooCat1.id;
                                 level2categories.Add(c);
                             }
                         }

                         foreach (ProductCategory category in level2categories)
                         {
                             ProductCategory existingCategory = wooProductCategories.FirstOrDefault(p => (p.slug == category.slug || p.name == category.name) && p.lang == category.lang);
                             ProductCategory newUpdatedCategory = null;
                             if (existingCategory == null)
                             {
                                 try
                                 {
                                     newUpdatedCategory = await wc.Category.Add(category);
                                 }
                                 catch (Exception ex)
                                 {
                                     Utils.WriteLog($"Error : wc.Category.Add() - {ex.Message}");
                                     continue;
                                 }
                                 if (newUpdatedCategory == null)
                                 {
                                     Utils.WriteLog($"Error : wc.Category.Add() returned null");
                                     continue;
                                 }
                                 category.id = newUpdatedCategory.id;
                             }
                             else
                             {
                                 try
                                 {
                                     ProductCategory existingCategory_en = wooProductCategories.FirstOrDefault(p => (p.slug == category.slug || p.name == category.name) && p.lang == "en");
                                     ProductCategory existingCategory_da = wooProductCategories.FirstOrDefault(p => (p.slug == category.slug || p.name == category.name) && p.lang == "da");

                                     category.id = existingCategory.id;
                                     category.translations.da = existingCategory_da != null ? existingCategory_da.id.Value : existingCategory.translations.da;
                                     category.translations.en = existingCategory_en != null ? existingCategory_en.id.Value : existingCategory.translations.en;
                                     newUpdatedCategory = await wc.Category.Update(category.id.Value, category);
                                 }
                                 catch (Exception ex)
                                 {
                                     Utils.WriteLog($"Error : wc.Category.Add() - {ex.Message}");
                                     continue;
                                 }
                                 if (newUpdatedCategory == null)
                                 {
                                     Utils.WriteLog($"Error : wc.Category.Add() returned null");
                                     continue;
                                 }
                             }
                         }*/

                foreach (ProductCategory category in categories)
                {
                    if (category.lang == "en")
                    {
                        ProductCategory category_da = wooProductCategories.FirstOrDefault(p => p.lang == "da" && p.name == category.name_da);
                        if (category_da != null)
                        {
                            category.translation_of = category_da.id.Value.ToString();
                            await wc.Category.Update(category.id.Value, category);
                            Utils.WriteLog($"Updated translation for category {category.id} {category.name} {category.lang}  {category.translation_of}");


                        }
                    }
                }

            }
            catch (Exception ex)
            {
                Utils.WriteLog($"Error SyncCategoires() - {ex.Message}");
                return false;

            }

            return true;
        }
    }
}
