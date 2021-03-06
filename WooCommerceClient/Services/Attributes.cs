using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WooCommerceClient.Models;
using WooCommerceClient.Models.Visma;

namespace WooCommerceClient.Services
{
    public class Attributes : BaseClass
    {

        internal static string SyncAttributTerms(DBaccess db, DateTime latestSyncTime)
        {
            AttributeChangeDates attributeChangeDates = new AttributeChangeDates();
            if (db.GetAttributeChangeDates(ref attributeChangeDates, out string errmsg) == false)
                Utils.WriteLog($"ERROR: db.GetAttributeChangeDates() - {errmsg}");

            List <ProductAttribute> wooCommerceAttributes = WooCommerceHelpers.GetWooCommerceAttributes().Result;

            ///////////////
            // Aargange
            ///////////////
            ///
           if (string.IsNullOrEmpty(SyncAttributeOnly) || SyncAttributeOnly.Contains("SyncAttributeForYear") || attributeChangeDates.ChDtYear >= latestSyncTime)
               SyncAttributeTermsForYear(db, out errmsg, wooCommerceAttributes);

            //////////////////
            // Vol
            //////////////////
            if (string.IsNullOrEmpty(SyncAttributeOnly) || SyncAttributeOnly.Contains("SyncAttributeForVolume") || attributeChangeDates.ChDtVolume >= latestSyncTime)
                SyncAttributeTermsForVolume(db, out errmsg, wooCommerceAttributes);


            /////////////////////
            // Type
            /////////////////////
            if (string.IsNullOrEmpty(SyncAttributeOnly) || SyncAttributeOnly.Contains("SyncAttributeForType") || attributeChangeDates.ChDtType >= latestSyncTime)
                SyncAttributeTermsForType(db, out errmsg, wooCommerceAttributes);

            ////////////////
            // Grapes
            ////////////////
            if (string.IsNullOrEmpty(SyncAttributeOnly) || SyncAttributeOnly.Contains("SyncAttributeForGrape") || attributeChangeDates.ChDtGrapes>= latestSyncTime)
                SyncAttributeTermsForGrape(db, out errmsg, wooCommerceAttributes);

            ///////////////
            // Land
            ///////////////              
            int idCountryAttribute = 0;
            if (string.IsNullOrEmpty(SyncAttributeOnly) || SyncAttributeOnly.Contains("SyncAttributeForCountry") || attributeChangeDates.ChDtCountry >= latestSyncTime)
            {
                idCountryAttribute = SyncAttributeTermsForCountry(db, out errmsg, wooCommerceAttributes);

                // Read back for region -> country relationship
                List<ProductAttributeTerm> existingWooCommerceTermsForCountry = WooCommerceHelpers.GetWooCommerceAttributeTerms(idCountryAttribute).Result;

                //////////////////////
                // Region
                //////////////////////
                int idRegionAttribute = 0;
                if (idCountryAttribute > 0 && (string.IsNullOrEmpty(SyncAttributeOnly) || SyncAttributeOnly.Contains("SyncAttributeForRegion") || attributeChangeDates.ChDtRegion >= latestSyncTime))
                    idRegionAttribute = SyncAttributeTermsForRegion(db, out errmsg, wooCommerceAttributes, existingWooCommerceTermsForCountry);

                //////////////////////
                ///// Producers
                //////////////////////
                if (idRegionAttribute > 0 && (string.IsNullOrEmpty(SyncAttributeOnly) || SyncAttributeOnly.Contains("SyncAttributeForProducer") || attributeChangeDates.ChDtProducerDa >= latestSyncTime || attributeChangeDates.ChDtProducerEn >= latestSyncTime))
                {
                    int idProducentAttribute = SyncAttributeTermsForProducer(idRegionAttribute, db, out errmsg, wooCommerceAttributes, existingWooCommerceTermsForCountry);

                    // Delete unused producers
                    if (string.IsNullOrEmpty(SyncAttributeOnly) || SyncAttributeOnly.Contains("DeleteUnusedAttributeTerms"))
                        DeleteUnusedAttributeTerms(db, out errmsg, idProducentAttribute);

                    if (Utils.ReadConfigInt32("DeleteIllegalProducers", 0) > 0)
                    {
                        var _ = DeleteIllegalProducers();
                    }
                }
            }
                             
            return errmsg;
        }

        private static bool SyncAttributeTermsForType(DBaccess db, out string errmsg, List<ProductAttribute> wooCommerceAttributes)
        {
            Utils.WriteLog("Sync'ing terms for Product types..");

            List<ProductAttributeTerm> attributeTerms = new List<ProductAttributeTerm>();
            if (db.GetAttributeTermsForType(ref attributeTerms, 45, out errmsg) == false)
            {
                Utils.WriteLog("ERROR db.GetAttributeTermsForType() - " + errmsg);
                return false;
            }
            List<ProductAttributeTerm> attributeTerms_en = new List<ProductAttributeTerm>();
            if (Utils.ReadConfigInt32("DoTranslation", 0) > 0)
            {
                if (db.GetAttributeTermsForType(ref attributeTerms_en, 44, out errmsg) == false)
                {
                    Utils.WriteLog("ERROR db.GetAttributeTermsForType() - " + errmsg);
                    return false;
                }
                // Ensure we have entry for 'en' for all..
                EnsureEnglishAttributTerms(ref attributeTerms, ref attributeTerms_en);
            }

            // add english to full list
            foreach (ProductAttributeTerm term in attributeTerms_en)
                attributeTerms.Add(term);

            // Now sync these types
            ProductAttribute a = wooCommerceAttributes.FirstOrDefault(p => p.name == Constants.AttributeNameType);
            int idTypeAttribute = a.id.Value;

            return SyncAttributeTerms(idTypeAttribute, attributeTerms).Result;
        }

        private static bool SyncAttributeTermsForGrape(DBaccess db, out string errmsg, List<ProductAttribute> wooCommerceAttributes)
        {
            Utils.WriteLog("Sync'ing terms for Grapes..");

            var attributeTerms = new List<ProductAttributeTerm>();
            var attributeTerms_en = new List<ProductAttributeTerm>();
            if (db.GetAttributeTermsForGrapes(ref attributeTerms, 45, out errmsg) == false)
            {
                Utils.WriteLog("ERROR db.GetAttributeTermsForGrapes() - " + errmsg);
                return false;
            }
            if (Utils.ReadConfigInt32("DoTranslation", 0) > 0)
            {
                if (db.GetAttributeTermsForGrapes(ref attributeTerms_en, 44, out errmsg) == false)
                {
                    Utils.WriteLog("ERROR db.GetAttributeTermsForGrapes() - " + errmsg);
                    return false;
                }

                EnsureEnglishAttributTerms(ref attributeTerms, ref attributeTerms_en);
            }
            // add english to full list
            foreach (ProductAttributeTerm term in attributeTerms_en)
                attributeTerms.Add(term);

            Utils.WriteLog("Syncing grapes..");
            var a = wooCommerceAttributes.FirstOrDefault(p => p.name == Constants.AttributeNameGrape);
            int idDrueAttribute = a.id.Value;
            var res = SyncAttributeTerms(idDrueAttribute, attributeTerms).Result;

            return res;
        }

        private static int SyncAttributeTermsForCountry(DBaccess db, out string errmsg, List<ProductAttribute> wooCommerceAttributes)
        {
            Utils.WriteLog("Sync'ing terms for Countries..");
            int idCountryAttribute = 0;
            var attributeTerms = new List<ProductAttributeTerm>();
            var attributeTerms_en = new List<ProductAttributeTerm>();
            if (db.GetAttributeTermsForCountry(ref attributeTerms, 45, out errmsg) == false)
                Utils.WriteLog("ERROR db.GetAttributeTermsForCountry() - " + errmsg);
            if (Utils.ReadConfigInt32("DoTranslation", 0) > 0)
            {
                if (db.GetAttributeTermsForCountry(ref attributeTerms_en, 44, out errmsg) == false)
                {
                    Utils.WriteLog("ERROR db.GetAttributeTermsForCountry() - " + errmsg);
                    return 0;
                }

                EnsureEnglishAttributTerms(ref attributeTerms, ref attributeTerms_en);
            }

            // add english to full list
            if (Utils.ReadConfigInt32("DoTranslation", 0) > 0)
            {
                foreach (ProductAttributeTerm term in attributeTerms_en)
                    attributeTerms.Add(term);
            }
            // Now sync countries
            Utils.WriteLog("Syncing countries..");
            var a = wooCommerceAttributes.FirstOrDefault(p => p.name == Constants.AttributeNameCountry);
            idCountryAttribute = a.id.Value;
            var res = SyncAttributeTerms(idCountryAttribute, attributeTerms).Result;

            return idCountryAttribute;
        }

        private static void LinkRegionsToCountries(List<ProductAttributeTerm> existingWooCommerceTermsForCountry, List<RegionCountryRelation> regionCountryRelationList, ref List<ProductAttributeTerm> attributeTerms, string lang)
        {
            DBaccess db = new DBaccess();
            // Add country id to region structure..
            foreach (ProductAttributeTerm term in attributeTerms)
            {
                RegionCountryRelation r = regionCountryRelationList.FirstOrDefault(p => p.Region == term.name);
                if (r != null)
                {
                    ProductAttributeTerm termCountry = existingWooCommerceTermsForCountry.FirstOrDefault(p => p.name == r.Country && p.lang == lang);
                    term._country_id = termCountry != null ? termCountry.id : 0;
                }
                if (term._country_id == 0)
                {
                    // Cant find country for region - try alternative way,,
                    string ctry = "";
                    if (db.GetAlternativeCountryForRegion(term.name, lang == "en" ? 44 : 45, ref ctry, out string errmsg) == false)
                        Utils.WriteLog("ERROR db.GetAlternativeCountryForRegion() - " + errmsg);
                    Utils.WriteLog($"Found alternative {lang} country relation {ctry}");
                    if (ctry != "")
                    {
                        ProductAttributeTerm termCountry = existingWooCommerceTermsForCountry.FirstOrDefault(p => p.name == ctry && p.lang == lang);
                        term._country_id = termCountry != null ? termCountry.id : 0;
                    }
                }
            }

        }

        private static int SyncAttributeTermsForRegion(DBaccess db, out string errmsg, List<ProductAttribute> wooCommerceAttributes, List<ProductAttributeTerm> existingWooCommerceTermsForCountry)
        {
            Utils.WriteLog("Sync'ing terms for Regions..");
            int idRegionAttribute = 0;
            var attributeTerms = new List<ProductAttributeTerm>();
            var attributeTerms_en = new List<ProductAttributeTerm>();
            List<RegionCountryRelation> regionCountryRelationList = new List<RegionCountryRelation>();
            if (db.GetAttributeTermsForRegion(ref attributeTerms, ref regionCountryRelationList, 45, out errmsg) == false)
            {
                Utils.WriteLog("ERROR db.GetAttributeTermsForRegion() - " + errmsg);
                return 0;
            }

            // Add country id to region structure..
            LinkRegionsToCountries(existingWooCommerceTermsForCountry, regionCountryRelationList, ref attributeTerms, "da");

            if (Utils.ReadConfigInt32("DoTranslation", 0) > 0)
            {
                List<RegionCountryRelation> regionCountryRelationList_en = new List<RegionCountryRelation>();
                if (db.GetAttributeTermsForRegion(ref attributeTerms_en, ref regionCountryRelationList_en, 44, out errmsg) == false)
                {
                    Utils.WriteLog("ERROR db.GetAttributeTermsForRegion() - " + errmsg);
                    return 0;
                }

                EnsureEnglishAttributTerms(ref attributeTerms, ref attributeTerms_en);

                // Add country id to region structure..
                LinkRegionsToCountries(existingWooCommerceTermsForCountry, regionCountryRelationList_en, ref attributeTerms_en, "en");

                // add english to full list
                foreach (ProductAttributeTerm term in attributeTerms_en)
                    attributeTerms.Add(term);
            }

            Utils.WriteLog("Syncing regions..");
            var a = wooCommerceAttributes.FirstOrDefault(p => p.name == Constants.AttributeNameRegion); // = område
            idRegionAttribute = a.id.Value;
            var res = SyncAttributeTerms(idRegionAttribute, attributeTerms).Result;

            return idRegionAttribute;
        }

        private static List<ProductAttributeTerm> GetAttributeTermsForProducer(int idRegionAttribute, DBaccess db,
                            List<ProductAttributeTerm> existingWooCommerceTermsForCountry, out string errmsg)
        {
            // Read back for producer -> region and country relationships.
         
            List<ProductAttributeTerm> existingWooCommerceTermsForRegion = WooCommerceHelpers.GetWooCommerceAttributeTerms(idRegionAttribute).Result;
            var attributeTerms = new List<ProductAttributeTerm>();
            var attributeTerms_en = new List<ProductAttributeTerm>();
            if (db.GetAttributeTermsForProducer(ref attributeTerms, false, 45, out errmsg) == false)
                Utils.WriteLog("ERROR db.GetAttributeTermsForProducer() - " + errmsg);
            if (Utils.ReadConfigInt32("DoTranslation", 0) > 0)
            {
             
                if (db.GetAttributeTermsForProducer(ref attributeTerms_en, false, 44, out errmsg) == false)
                    Utils.WriteLog("ERROR db.GetAttributeTermsForProducer() - " + errmsg);

                EnsureEnglishAttributTerms(ref attributeTerms, ref attributeTerms_en);
            }
            // Fetch producer-country-region relationships (language dependent!)
            List<VismaProducer> producerInfo = new List<VismaProducer>();
            List<VismaProducer> producerInfo_en = new List<VismaProducer>();
            if (db.GetProducerInfo(ref producerInfo, 45, out errmsg) == false)
            {
                Utils.WriteLog("ERROR db.GetProducerInfo(45) - " + errmsg);
            }

            if (Utils.ReadConfigInt32("DoTranslation", 0) > 0)
            {
                if (db.GetProducerInfo(ref producerInfo_en, 44, out errmsg) == false)
                {
                    Utils.WriteLog("ERROR db.GetProducerInfo(44) - " + errmsg);
                }
            }

            foreach (ProductAttributeTerm term in attributeTerms)
            {

                VismaProducer vp = producerInfo.FirstOrDefault(p => p.Producer == term.name.Replace("&amp;", "&"));
                if (vp != null)
                {
                    ProductAttributeTerm termCountry = existingWooCommerceTermsForCountry.FirstOrDefault(p => p.name == vp.ProducerCountry && p.lang == "da");
                    ProductAttributeTerm termRegion = existingWooCommerceTermsForRegion.FirstOrDefault(p => p.name == vp.ProducerRegion && p.lang == "da");
                    term._country_id = termCountry != null ? termCountry.id : 0;
                    term._region_id = termRegion != null ? termRegion.id : 0;
                }

                Utils.WriteLog($"INFO: Producer {term.name} links to regionid:{term._region_id} countryid:{term._country_id}");
            }

            if (Utils.ReadConfigInt32("DoTranslation", 0) > 0)
            {
                foreach (ProductAttributeTerm term in attributeTerms_en)
                {

                    VismaProducer vp = producerInfo_en.FirstOrDefault(p => p.Producer == term.name.Replace("&amp;", "&"));
                    if (vp != null)
                    {
                        ProductAttributeTerm termCountry = existingWooCommerceTermsForCountry.FirstOrDefault(p => p.name == vp.ProducerCountry && p.lang == "en");
                        ProductAttributeTerm termRegion = existingWooCommerceTermsForRegion.FirstOrDefault(p => p.name == vp.ProducerRegion && p.lang == "en");
                        term._country_id = termCountry != null ? termCountry.id : 0;
                        term._region_id = termRegion != null ? termRegion.id : 0;
                    }

                    Utils.WriteLog($"INFO: Producer {term.name} regionid:{term._region_id} countryid:{term._country_id}");
                }

                foreach (ProductAttributeTerm term in attributeTerms_en)
                    attributeTerms.Add(term);
            }

            return attributeTerms;
        }

        private static List<ProductAttribute> GetInitialAttributes()
        {
            List<ProductAttribute> attributes = new List<ProductAttribute>();
            var types = new List<string>()
            {
                Constants.AttributeNameType,
                Constants.AttributeNameGrape,
                Constants.AttributeNameProducer,
                Constants.AttributeNameCountry,
                Constants.AttributeNameRegion,
                Constants.AttributeNameYear,
                Constants.AttributeNameVolume
            };

            foreach (var type in types)
                attributes.Add(new ProductAttribute()
                {
                    name = type,
                    slug = Utils.SanitizeSlugNameNew(type)
                });


            return attributes;
        }

  

        private static int SyncAttributeTermsForProducer(int idRegionAttribute, DBaccess db, out string errmsg,
            List<ProductAttribute> wooCommerceAttributes, List<ProductAttributeTerm> existingWooCommerceTermsForCountry)
        {
            var attributeTerms = GetAttributeTermsForProducer(idRegionAttribute, db, existingWooCommerceTermsForCountry, out errmsg);

            Utils.WriteLog("Sync'ing terms for Producers..");

            var a = wooCommerceAttributes.FirstOrDefault(p => p.name == Constants.AttributeNameProducer);
            var idProducentAttribute = a.id.Value;
            var res = SyncAttributeTerms(idProducentAttribute, attributeTerms).Result;

            // Special link between related producers..
            List<ProductAttributeTerm> attributeTerms2 = GetAttributesForRelatedProducers(idProducentAttribute, db, ref errmsg, attributeTerms);
            if (attributeTerms2.Count > 0)
            {
                Utils.WriteLog($"Syncing producer relations..");
                res = SyncAttributeTerms(idProducentAttribute, attributeTerms2).Result;
            }

            return idProducentAttribute;
        }

        private static List<ProductAttributeTerm> DeleteUnusedAttributeTerms(DBaccess db, out string errmsg, int idProducentAttribute)
        {
            Utils.WriteLog("Deling unused Producers..");

            var attributeTerms = new List<ProductAttributeTerm>();
            if (db.GetAttributeTermsForProducerToDelete(ref attributeTerms, out errmsg) == false)
                Utils.WriteLog("ERROR db.GetAttributeTermsForProducerToDelete() - " + errmsg);

            var res = SyncDeletedAttributeTerms(idProducentAttribute, attributeTerms).Result;
            return attributeTerms;
        }

        private static void SyncAttributeTermsForVolume(DBaccess db, out string errmsg, List<ProductAttribute> wooCommerceAttributes)
        {
            Utils.WriteLog("Sync'ing terms for Volumes..");

            var attributeTerms = new List<ProductAttributeTerm>();
            var attributeTerms_en = new List<ProductAttributeTerm>();
            if (db.GetAttributeTermsForVolume(ref attributeTerms, 45, out errmsg) == false)
                Utils.WriteLog("ERROR db.GetAttributeTermsForVolume() - " + errmsg);
            if (Utils.ReadConfigInt32("DoTranslation", 0) > 0)
            {
                if (db.GetAttributeTermsForVolume(ref attributeTerms_en, 44, out errmsg) == false)
                    Utils.WriteLog("ERROR db.GetAttributeTermsForVolume() - " + errmsg);

                EnsureEnglishAttributTerms(ref attributeTerms, ref attributeTerms_en);

                // add english to full list
                foreach (ProductAttributeTerm term in attributeTerms_en)
                    attributeTerms.Add(term);
            }

            Utils.WriteLog("Syncing volume..");
            var a = wooCommerceAttributes.FirstOrDefault(p => p.name == Constants.AttributeNameVolume);
            int idVolumeAttribute = a.id.Value;
            var res = SyncAttributeTerms(idVolumeAttribute, attributeTerms).Result;
        }

        private static List<ProductAttributeTerm> SyncAttributeTermsForYear(DBaccess db, out string errmsg, List<ProductAttribute> wooCommerceAttributes)
        {
            Utils.WriteLog("Sync'ing terms for Years..");

            var attributeTerms = new List<ProductAttributeTerm>();
            var attributeTerms_en = new List<ProductAttributeTerm>();
            if (db.GetAttributeTermsForYear(ref attributeTerms, 45, out errmsg) == false)
                Utils.WriteLog("ERROR db.GetAttributeTermsForYear() - " + errmsg);

            if (Utils.ReadConfigInt32("DoTranslation", 0) > 0)
            {
                if (db.GetAttributeTermsForYear(ref attributeTerms_en, 44, out errmsg) == false)
                    Utils.WriteLog("ERROR db.GetAttributeTermsForYear() - " + errmsg);

                EnsureEnglishAttributTerms(ref attributeTerms, ref attributeTerms_en);

                // add english to full list
                foreach (ProductAttributeTerm term in attributeTerms_en)
                    attributeTerms.Add(term);
            }

            Utils.WriteLog("Syncing years..");
            var a = wooCommerceAttributes.FirstOrDefault(p => p.name == Constants.AttributeNameYear);
            int idYearAttribute = a.id.Value;
            var res = SyncAttributeTerms(idYearAttribute, attributeTerms).Result;

            return attributeTerms;
        }
 
        private static List<ProductAttributeTerm> GetAttributesForRelatedProducers(int idProducerAttribute, DBaccess db, ref string errmsg, List<ProductAttributeTerm> attributeTerms)
        {
            List<ProductAttributeTerm> attributeTerms2 = new List<ProductAttributeTerm>();
            var wooCommerceAttributeTermsForProducers = WooCommerceHelpers.GetWooCommerceAttributeTerms(idProducerAttribute).Result;

            foreach (ProductAttributeTerm term in attributeTerms)
            {
                List<int> producerRelations = new List<int>();
                if (db.GetAttributeTermsForRelatedProducers(term.visma_id, ref producerRelations, out errmsg) == false)
                    Utils.WriteLog("ERROR db.GetAttributeTermsForRelatedProducers() - " + errmsg);

                if (producerRelations.Count > 0)
                {
                    term._related_producers = new List<int>();
                    foreach (int visma_id in producerRelations)
                    {
                        ProductAttributeTerm existingTermForProducer = wooCommerceAttributeTermsForProducers.FirstOrDefault(
                            p => p.visma_id == visma_id && p.lang == term.lang);
                        if (existingTermForProducer != null)
                        {
                            term._related_producers.Add(existingTermForProducer.id.Value);
                        }
                    }

                    attributeTerms2.Add(term);
                }
            }

            return attributeTerms2;
        }

        private static void EnsureEnglishAttributTerms(ref List<ProductAttributeTerm> terms_da, ref List<ProductAttributeTerm> terms_en)
        {
            // Ensure we have entry for 'en' for all..
            foreach (ProductAttributeTerm term_da in terms_da)
            {
                ProductAttributeTerm term_en = terms_en.FirstOrDefault(p => p.visma_id == term_da.visma_id);
                if (term_en == null)
                {
                    term_en = new ProductAttributeTerm()
                    {
                        name = term_da.name,
                        visma_id = term_da.visma_id,
                        lang = "en",
                        slug = null // term_da.slug + "-en",
                        //translations = new Translations()
                    };

                    terms_en.Add(term_en);
                }

            }
        }

        private static async Task<bool> SyncDeletedAttributeTerms(int idAttribute, List<ProductAttributeTerm> attributeTermsToDelete)
        {
            try
            {
                MyRestAPI rest = new MyRestAPI(Utils.ReadConfigString("WooCommerceUrl", ""),
                         Utils.ReadConfigString("WooCommerceKey", ""),
                         Utils.ReadConfigString("WooCommerceSecret", ""));
                WCObject wc = new WCObject(rest);

                List<ProductAttributeTerm> wooCommerceAttributeTerms = await WooCommerceHelpers.GetWooCommerceAttributeTerms(idAttribute);
                Utils.WriteLog($"Existing attribute terms {wooCommerceAttributeTerms.Count}");

                // bool readbackAttributes = true;
                foreach (ProductAttributeTerm term in attributeTermsToDelete)
                {
                    ProductAttributeTerm existingAttribute = wooCommerceAttributeTerms.FirstOrDefault(p => p.visma_id == term.visma_id && p.lang == term.lang);

                    if (existingAttribute != null)
                    {
                        try
                        {
                            Utils.WriteLog($"Sync'ing delete attribute term {term.name} .");
                            term.id = existingAttribute.id;

                            string ret = await wc.Attribute.Terms.Delete(existingAttribute.id.Value, idAttribute, true);
                            Utils.WriteLog(ret);
                        }
                        catch (Exception ex)
                        {
                            Utils.WriteLog($"Error : wc.Attribute.Delete() - {ex.Message}");
                            //                            return false;
                        }

                    }

                }


            }
            catch (Exception ex)
            {
                Utils.WriteLog($"Error SyncDeletedAttributeTerms() - {ex.Message}");
                return false;

            }

            return true;
        }

        private static async Task<bool> SyncAttributeTerms(int idAttribute, List<ProductAttributeTerm> attributeTerms)
        {
            try
            {
                bool resyncAfterCreate = false;
                var wcApi = new WC_API();

                List<ProductAttributeTerm> wooCommerceAttributeTerms = await WooCommerceHelpers.GetWooCommerceAttributeTerms(idAttribute);
                Utils.WriteLog($"Existing attribute terms {wooCommerceAttributeTerms.Count} for attribure id {idAttribute}");

                foreach (ProductAttributeTerm term in attributeTerms)
                {
                    if (SyncAttributeTerm(idAttribute, wcApi, wooCommerceAttributeTerms, term) == true)
                        resyncAfterCreate = true;
                }

                // reread all
                wooCommerceAttributeTerms = await WooCommerceHelpers.GetWooCommerceAttributeTerms(idAttribute);

                // establish language links (if new elements added)
                if (resyncAfterCreate && wooCommerceAttributeTerms != null)
                {
                  
                    // Set default if no translations set..
                    foreach (ProductAttributeTerm term in attributeTerms)
                    {
                        term.slug = null;
                        if (term.lang == "en")
                        {
                            ProductAttributeTerm term_da = wooCommerceAttributeTerms.FirstOrDefault(p => p.lang == "da" && p.visma_id == term.visma_id);
                            if (term_da != null)
                            {

                                term.translation_of = term_da.id.ToString();
                                await wcApi.Update(term, term.lang, idAttribute);
                                Utils.WriteLog($"Updated translation for term {term.id} {term.name} {term.lang} {term.translation_of} ");
                            }
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                Utils.WriteLog($"Error SyncAttributeTerms() - {ex.Message}");
                return false;

            }

            return true;
        }

        private static bool SyncAttributeTerm(int idAttribute, WC_API wcApi, List<ProductAttributeTerm> wooCommerceAttributeTerms, ProductAttributeTerm term)
        {
            term.slug = null;

            bool resyncTranslations = false;

            string s = "";
            if (term._related_producers != null)
                s = Utils.IntList2String(term._related_producers);

            ProductAttributeTerm existingAttribute = wooCommerceAttributeTerms.FirstOrDefault(p =>
                p.lang == term.lang &&
                ((term.visma_id > 0 && p.visma_id == term.visma_id) ||  // Try visma-id first!
                 (term.visma_id <= 0 && p.name == term.name)) // fall back to name
            );

            ProductAttributeTerm newAttribute = null;
            if (existingAttribute == null)
            {
                Utils.WriteLog($"Sync'ing new attribute term {term.name} {term.slug ?? ""} Crtyid:{term._country_id} RegionID:{term._region_id} Visma-Id:{term.visma_id} relations: {s}");
                newAttribute = wcApi.Add(term, term.lang, idAttribute).Result; // await wc.Attribute.Terms.Add(term, idAttribute);
                resyncTranslations = true;
            }
            else
            {
                var updMsg = "";
                term.id = existingAttribute.id.Value;
                if (idAttribute == 167)
                    updMsg = $"Sync'ing Producer existing attribute term {existingAttribute.id.Value} {term.name} {term.slug ?? ""} Crtyid:{term._country_id} RegionID:{term._region_id} VismaID:{term.visma_id} Relations:{s}";
                else
                    updMsg = $"Sync'ing existing attribute term {existingAttribute.id.Value} {term.name} {term.slug ?? ""} Crtyid:{term._country_id} RegionID:{term._region_id} VismaID:{term.visma_id}..";

                Utils.WriteLog(updMsg);
                if (term.lang == "en")
                {
                    ProductAttributeTerm term_da = wooCommerceAttributeTerms.FirstOrDefault(p => p.lang == "da" && p.visma_id == term.visma_id);
                    if (term_da != null)
                    {
                        term.translation_of = term_da.id.ToString();
                    }
                    else
                        resyncTranslations = true;
                }

                newAttribute = wcApi.Update(term, term.lang, idAttribute).Result;
            }

            return resyncTranslations;
        }

        public static async Task<bool> SyncAttributes()
        {
            Utils.WriteLog($"SYNC Attributes..");
            try
            {
                List<ProductAttribute> attributes = GetInitialAttributes();
                var wcApi = new WC_API();

                List<ProductAttribute> wooCommerceAttributes = await WooCommerceHelpers.GetWooCommerceAttributes();
                Utils.WriteLog($"Exisiting attributes {wooCommerceAttributes.Count}");

                foreach (ProductAttribute attribute in attributes)
                {
                    ProductAttribute existingAttribute = wooCommerceAttributes.FirstOrDefault(p =>
                        (p.name == attribute.name || p.slug == attribute.slug || p.slug == "pa_" + attribute.slug));
                    if (existingAttribute == null)
                    {
                        Utils.WriteLog($"Sync'ing attribute {attribute.name} {attribute.slug ?? ""}..");
                        await wcApi.Add(attribute);
                    }
                    else
                    {
                        attribute.id = existingAttribute.id.Value;
                        var res = await wcApi.Update(attribute);
                        if (res == null)
                            Utils.WriteLog($"Failed to update attribute name {attribute.name} id {attribute.id.Value}");
                    }
                }
            }
            catch (Exception ex)
            {
                Utils.WriteLog($"Error SyncAttributes()", ex);
                return false;
            }

            return true;
        }

  /*      public static bool ListUnrelatedEnglishAttributes()
        {
            List<ProductAttribute> wooCommerceAttributes = WooCommerceHelpers.GetWooCommerceAttributes().Result;
            if (wooCommerceAttributes == null)
                return false;

            ProductAttribute a = wooCommerceAttributes.FirstOrDefault(p => p.name == Constants.AttributeNameVolume);
            if (a == null)
                return false;
            List<int> unrelatedAttributes = GetUnrelatedEnglishAttributes(a.id.Value);
            if (unrelatedAttributes != null)
                Utils.WriteLog($"Found {unrelatedAttributes.Count} unrelated Volume terms");

            a = wooCommerceAttributes.FirstOrDefault(p => p.name == Constants.AttributeNameYear);
            if (a == null)
                return false;
            unrelatedAttributes = GetUnrelatedEnglishAttributes(a.id.Value);
            if (unrelatedAttributes != null)
                Utils.WriteLog($"Found {unrelatedAttributes.Count} unrelated Year terms");

            a = wooCommerceAttributes.FirstOrDefault(p => p.name == Constants.AttributeNameType);
            if (a == null)
                return false;
            unrelatedAttributes = GetUnrelatedEnglishAttributes(a.id.Value);
            if (unrelatedAttributes != null)
                Utils.WriteLog($"Found {unrelatedAttributes.Count} unrelated Type terms");

            a = wooCommerceAttributes.FirstOrDefault(p => p.name == Constants.AttributeNameCountry);
            if (a == null)
                return false;
            unrelatedAttributes = GetUnrelatedEnglishAttributes(a.id.Value);
            if (unrelatedAttributes != null)
                Utils.WriteLog($"Found {unrelatedAttributes.Count} unrelated Country terms");

            a = wooCommerceAttributes.FirstOrDefault(p => p.name == Constants.AttributeNameRegion);
            if (a == null)
                return false;
            unrelatedAttributes = GetUnrelatedEnglishAttributes(a.id.Value);
            if (unrelatedAttributes != null)
                Utils.WriteLog($"Found {unrelatedAttributes.Count} unrelated Region terms");

            a = wooCommerceAttributes.FirstOrDefault(p => p.name == Constants.AttributeNameGrape);
            if (a == null)
                return false;
            unrelatedAttributes = GetUnrelatedEnglishAttributes(a.id.Value);
            if (unrelatedAttributes != null)
                Utils.WriteLog($"Found {unrelatedAttributes.Count} unrelated Grape terms");

            a = wooCommerceAttributes.FirstOrDefault(p => p.name == Constants.AttributeNameProducer);
            if (a == null)
                return false;
            unrelatedAttributes = GetUnrelatedEnglishAttributes(a.id.Value);
            if (unrelatedAttributes != null)
                Utils.WriteLog($"Found {unrelatedAttributes.Count} unrelated Producer terms");


            return true;
        }
  */
    /*    public static List<int> GetUnrelatedEnglishAttributes(int idAttribute)
        {
            List<int> unrelatedAttributes = new List<int>();
            MyRestAPI rest = new MyRestAPI(Utils.ReadConfigString("WooCommerceUrl", ""),
            Utils.ReadConfigString("WooCommerceKey", ""),
            Utils.ReadConfigString("WooCommerceSecret", ""));
            WCObject wc = new WCObject(rest);
            List<ProductAttributeTerm> wooCommerceAttributeTerms = WooCommerceHelpers.GetWooCommerceAttributeTerms(idAttribute).Result;
            if (wooCommerceAttributeTerms == null)
                return null;

            foreach (ProductAttributeTerm  term in wooCommerceAttributeTerms)
            {
                if (term.lang == "en")
                    if (term.translations != null)
                        if (term.translations.da.HasValue == false || term.translations.da.Value == 0)
                            unrelatedAttributes.Add(term.id.Value);
            }

            return unrelatedAttributes;
        }
*/

        private static async Task<bool> DeleteUnusedProducers(List<ProductAttributeTerm> usedattributeTermsProducers)
        {
            MyRestAPI rest = new MyRestAPI(Utils.ReadConfigString("WooCommerceUrl", ""),
                        Utils.ReadConfigString("WooCommerceKey", ""),
                        Utils.ReadConfigString("WooCommerceSecret", ""));
            WCObject wc = new WCObject(rest);

            Utils.WriteLog($"DeleteUnusedProducers: Active producers in Visma: {usedattributeTermsProducers.Count}");

            if (usedattributeTermsProducers.Count == 0)
                return false;
            List<ProductAttribute> wooCommerceAttributes = WooCommerceHelpers.GetWooCommerceAttributes().Result;

            ProductAttribute a = wooCommerceAttributes.FirstOrDefault(p => p.name == Constants.AttributeNameProducer);
            if (a == null)
                return false;
            int idProducentAttribute = a.id.Value;
            List<ProductAttributeTerm> existingWooCommerceAttributeTerms = WooCommerceHelpers.GetWooCommerceAttributeTerms(idProducentAttribute).Result;
            if (existingWooCommerceAttributeTerms == null)
                return false;

            Utils.WriteLog($"DeleteUnusedProducers: Active producers in WooCommerce: {existingWooCommerceAttributeTerms.Count}");

            foreach (ProductAttributeTerm wooterm in existingWooCommerceAttributeTerms)
            {

                ProductAttributeTerm vismaProducerTerm = usedattributeTermsProducers.FirstOrDefault(p => (p.slug == wooterm.slug || p.slug == "pa_" + wooterm.slug || p.name == wooterm.name));
                if (vismaProducerTerm == null)
                {
                    string deleteString = "";
                    Utils.WriteLog($"DeleteUnusedProducers: Term {wooterm.name} is not used!");
                    try
                    {
                        deleteString = await wc.Attribute.Terms.Delete(wooterm.id.Value, idProducentAttribute, true);
                    }
                    catch (Exception ex)
                    {
                        Utils.WriteLog($"Exception  wc.Attribute.Terms.Delete() - {ex.Message}");
                        Utils.WriteLog($"{ex.StackTrace}");
                        break;
                    }
                    Utils.WriteLog($"Producer {wooterm.name} deleted {deleteString}.");
                }
            }

            return true;
        }

        public static async Task<bool> DeleteIllegalProducers()
        {
            MyRestAPI rest = new MyRestAPI(Utils.ReadConfigString("WooCommerceUrl", ""),
                       Utils.ReadConfigString("WooCommerceKey", ""),
                       Utils.ReadConfigString("WooCommerceSecret", ""));
            WCObject wc = new WCObject(rest);

            List<ProductAttribute> wooCommerceAttributes = WooCommerceHelpers.GetWooCommerceAttributes().Result;

            ProductAttribute a = wooCommerceAttributes.FirstOrDefault(p => p.name == Constants.AttributeNameProducer);
            if (a == null)
                return false;
            int idProducentAttribute = a.id.Value;
            List<ProductAttributeTerm> existingWooCommerceAttributeTerms = WooCommerceHelpers.GetWooCommerceAttributeTerms(idProducentAttribute).Result;
            if (existingWooCommerceAttributeTerms == null)
                return false;

            foreach(ProductAttributeTerm term in existingWooCommerceAttributeTerms)
            {
                if (term.visma_id <= 0)
                {
                    Utils.WriteLog($"WARNING: Illegal visma-ID found for term {term.id.Value} {term.name}");
                    string deleteString = "";
                    Utils.WriteLog($"DeleteUnusedProducers: Term {term.name} is not used!");
                    try
                    {
                        deleteString = await wc.Attribute.Terms.Delete(term.id.Value, idProducentAttribute, true);
                    }
                    catch (Exception ex)
                    {
                        Utils.WriteLog($"Exception  wc.Attribute.Terms.Delete() - {ex.Message}");
                        Utils.WriteLog($"{ex.StackTrace}");
                        break;
                    }
                    Utils.WriteLog($"Producer {term.name} deleted {deleteString}.");
                }
            }
            return true;
        }

        public static  bool DeleteAllAttributes()
        {
            MyRestAPI rest = new MyRestAPI(Utils.ReadConfigString("WooCommerceUrl", ""),
                Utils.ReadConfigString("WooCommerceKey", ""),
                Utils.ReadConfigString("WooCommerceSecret", ""));
            List<ProductAttribute> wooCommerceAttributes = WooCommerceHelpers.GetWooCommerceAttributes().Result;
            if (wooCommerceAttributes == null)
                return false;
            foreach (ProductAttribute a in wooCommerceAttributes)
            {
                bool _ =  DeleteAllTerms(a.id.Value).Result;
            }

            return true;
        }

        public static async Task<bool> DeleteAllTerms(int idAttribute)
        {
            MyRestAPI rest = new MyRestAPI(Utils.ReadConfigString("WooCommerceUrl", ""),
                        Utils.ReadConfigString("WooCommerceKey", ""),
                        Utils.ReadConfigString("WooCommerceSecret", ""));
            WCObject wc = new WCObject(rest);
            List<ProductAttributeTerm>wooCommerceAttributeTerms = WooCommerceHelpers.GetWooCommerceAttributeTerms(idAttribute).Result;
            if (wooCommerceAttributeTerms == null)
                return false;
            foreach(ProductAttributeTerm term in wooCommerceAttributeTerms)
            {
                Utils.WriteLog($"DeleteAllTerms: Deleting term {term.name}..");
                string deleteString = "";
                try
                {
                    deleteString = await wc.Attribute.Terms.Delete(term.id.Value, idAttribute, true);
                }
                catch (Exception ex)
                {
                    Utils.WriteLog($"Exception  wc.Attribute.Terms.Delete() - {ex.Message}");
                    Utils.WriteLog($"{ex.StackTrace}");
                    return false;
                }
                Utils.WriteLog($"Term {term.name} deleted {deleteString}.");
            }

            return true;
        }

    }
}
