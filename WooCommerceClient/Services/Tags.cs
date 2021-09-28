using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WooCommerceClient.Models;
using WooCommerceClient.Models.Visma;

namespace WooCommerceClient.Services
{
    class Tags
    {
        private static async Task<bool> SyncTags(List<ProductTag> tags)
        {
            Utils.WriteLog($"SYNC Tags..");
            try
            {
                MyRestAPI rest = new MyRestAPI(Utils.ReadConfigString("WooCommerceUrl", ""),
                         Utils.ReadConfigString("WooCommerceKey", ""),
                         Utils.ReadConfigString("WooCommerceSecret", ""));
                WCObject wc = new WCObject(rest);

                List<ProductTag> wooCommerceTags = await WooCommerceHelpers.GetWooCommerceTags();

                foreach (ProductTag tag in tags)
                {
                    ProductTag existingTag = wooCommerceTags.FirstOrDefault(p => 
                        (p.slug == tag.slug || GetAlcoholValueWithoutPrefix(p.name) == GetAlcoholValueWithoutPrefix(tag.name)) 
                        && p.lang == tag.lang);
                    ProductTag newTag = null;
                    if (existingTag == null)
                    {
                        try
                        {
                            Utils.WriteLog($"Sync'ing new tag {tag.name} {tag.slug} {tag.lang}..");
                            newTag = await wc.Tag.Add(tag);

                        }
                        catch (Exception ex)
                        {
                            Utils.WriteLog($"Error : wc.Tag.Add() - {ex.Message} {tag.name} {tag.slug} {tag.lang}");
                            return false;
                        }
                        if (newTag == null)
                        {
                            Utils.WriteLog($"Error : wc.Tag.Add() returned null");
                            return false;
                        }
                        tag.id = newTag.id;

                    }
                    else
                    {
                        Utils.WriteLog($"Sync'ing existing tag id:{tag.id} {tag.name} {tag.slug} {tag.lang}..");
                        if (tag.lang == "en")
                        {
                            ProductTag tag_da = wooCommerceTags.FirstOrDefault(p => p.lang == "da" &&
                                GetAlcoholValueWithoutPrefix(p.name) == GetAlcoholValueWithoutPrefix(tag.name));
                            if (tag_da != null)
                                tag.translation_of = tag_da.id.Value.ToString();
                        }
                        tag.id = existingTag.id;
                        CheckAndUpdateAlcoholPrefix(tag);
                        await wc.Tag.Update(existingTag.id.Value, tag);
                    }
                }

                // SET LANGUAGE LINKS.
                wooCommerceTags = await WooCommerceHelpers.GetWooCommerceTags();
                if (wooCommerceTags != null)
                {

                    foreach (ProductTag tag in tags)
                    {
                        if (tag.lang == "en")
                        {
                            ProductTag tag_da = wooCommerceTags.FirstOrDefault(p => p.lang == "da" &&
                                GetAlcoholValueWithoutPrefix(p.name) == GetAlcoholValueWithoutPrefix(tag.name));
                            if (tag_da != null)
                                tag.translation_of = tag_da.id.Value.ToString();
                        }


                    }

                    foreach (ProductTag tag in tags)
                    {
                        if (tag.lang == "en")
                        {
                            await wc.Tag.Update(tag.id.Value, tag);
                            Utils.WriteLog($"Updating lang attributes on  existing tag {tag.id} {tag.name} {tag.slug} {tag.lang}  {tag.translation_of}..");
                        }
                    }


                }
            }
            catch (Exception ex)
            {
                Utils.WriteLog($"Error SyncTags() - {ex.Message}");
                return false;

            }

            return true;
        }

        private static void CheckAndUpdateAlcoholPrefix(ProductTag tag)
        {
            string old_prefix = tag.lang != "da" ? 
                Constants.Old_AlcoholPrefixEn : Constants.Old_AlcoholPrefix;
            string prefix = tag.lang != "da" ? 
                Constants.AlcoholPrefixEn : Constants.AlcoholPrefix;
            if (tag.name.StartsWith(old_prefix))
                tag.name = tag.name.Replace(old_prefix, prefix);
        }

        private static string GetAlcoholValueWithoutPrefix(string name)
        {
            return name.Replace(Constants.AlcoholPrefix, "").Replace(Constants.AlcoholPrefixEn, "").
                Replace(Constants.Old_AlcoholPrefix, "").Replace(Constants.Old_AlcoholPrefixEn, "");
        }

        internal static void Sync(DBaccess db, out string errmsg)
        {
            List<ProductTag> tags = new List<ProductTag>();
            //    if (db.GetTagsForYear(ref tags, out errmsg) == false)
            //        Utils.WriteLog("ERROR db.GetTagsForYear() - " + errmsg);

            // adds to list..
            //    if (db.GetTagsForVolume(ref tags, out errmsg) == false)
            //        Utils.WriteLog("ERROR db.GeGetTagsForVoumetTags() - " + errmsg);

            // adds to list..
            if (db.GetTagsForAlcohol(ref tags, out errmsg) == false)
                Utils.WriteLog("ERROR db.GeGetTagsForVoumetTags() - " + errmsg);

            //  tags.Add(new ProductTag() { name = Constants.NoDiscount, slug = Utils.SanitizeSlugName(Constants.NoDiscount) });

            bool res = SyncTags(tags).Result;
        }
    }
}
