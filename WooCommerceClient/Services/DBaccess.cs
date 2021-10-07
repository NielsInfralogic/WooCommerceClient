using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WooCommerceClient.Models;
using WooCommerceClient.Models.Visma;

namespace WooCommerceClient.Services
{
    public class DBaccess : BaseClass
    {
        readonly SqlConnection connection;
        public static string queryGetCustomerFromEmail =
          "SELECT TOP 1 A.CustNo,A.MailAd,A.Nm,A.Ad1,A.Ad2,A.Ad3,A.PNo,A.PArea,A.Phone,A.MobPh,ISNULL(D.Nm,A.Nm),ISNULL(D.Ad1,A.Ad1),ISNULL(D.Ad2,A.Ad2),ISNULL(D.Ad3,A.Ad3),ISNULL(D.PNo,A.PNo),ISNULL(D.PArea,A.PArea) " +
          "FROM Actor A WITH (NOLOCK) " +
          "LEFT OUTER JOIN Actor D  WITH (NOLOCK) ON A.ActNo=D.DelToAct " +
          "WHERE A.CustNo>0 AND A.MailAd= '#1#' ORDER BY A.CustNo DESC";

        public static string queryGetCustomerFromCustNo =
           "SELECT TOP 1 A.CustNo,A.MailAd,A.Nm,A.Ad1,A.Ad2,A.Ad3,A.PNo,A.PArea,A.Phone,A.MobPh,ISNULL(D.Nm,A.Nm),ISNULL(D.Ad1,A.Ad1),ISNULL(D.Ad2,A.Ad2),ISNULL(D.Ad3,A.Ad3),ISNULL(D.PNo,A.PNo),ISNULL(D.PArea,A.PArea) " +
           "FROM Actor A WITH (NOLOCK) " +
           "LEFT OUTER JOIN Actor D  WITH (NOLOCK) ON A.ActNo=D.DelToAct " +
           "WHERE A.CustNo = #1# ORDER BY A.CustNo DESC";

        public static string queryGetOrderByOrdNo =
           "SELECT DISTINCT Ord.Gr12, Ord.OrdDt, Ord.OrdNo, Ord.Nm, Ord.Ad1, Ord.Ad2, Ord.Ad3, Ord.PNo, Ord.PArea, " +
           "Ord.DelNm, Ord.DelAd1, Ord.DelAd2, Ord.DelAd3, Ord.DelPNo, Ord.DelPArea, Ord.DelTrm, Ord.DelDt, Ord.Inf, " +
           "Ord.Inf2, Ord.CSOrdNo, Ord.Inf4, Ord.InvoNo, Ord.LstInvDt,Ord.Inf5,Ord.FrAm, Ord.NoteNm, Ord.Inf3, Ord.CardNm " +
           "FROM Ord WITH (NOLOCK) WHERE Ord.Gr12>0";

        public static string queryGetOrderLinesByOrdNo =
            "SELECT DISTINCT OrdLn.LnNo, OrdLn.ProdNo, OrdLn.NoOrg, OrdLn.Price, OrdLn.NoInvo " +
            "FROM OrdLn WITH (NOLOCK) WHERE OrdLn.OrdNo= #1#";


        public static string queryGetDisabledProductList =
            "SELECT DISTINCT Prod.ProdNo  FROM Prod WITH (NOLOCK) " +
            "INNER JOIN FreeInf1 WITH (NOLOCK) ON FreeInf1.ProdNo=Prod.ProdNo AND FreeInf1.InfCatNo= 45 " +
            "WHERE FreeInf1.#9#=9 ";

        public static string queryGetProductList =
            "SELECT DISTINCT Prod.ProdNo, Prod.StSaleUn, Prod.Descr, ISNULL(Prod.WebPg2,'') as WebPg2, Prod.Gr10, Prod.ProdTp4, Prod.ProdTp2, " +
            "ISNULL(TXT2.Txt,'') as gruppe, " +
            "ISNULL(TXT3.Txt,'') as type, " +
            "ISNULL((SELECT TOP 1 ISNULL(ScanCd.SCd, '') FROM ScanCd WHERE ScanCd.ProdNo = Prod.ProdNo ORDER BY ChDt Desc),''), " +
            "ISNULL(TXT5.txt, ISNULL(PC1.Descr, '')) as land, " +
            "ISNULL(TXT4.txt, ISNULL(PC2.Descr, '')) as omraade, " +
            "ISNULL(TXT6.Txt,'') as kommune, " +
            "ISNULL(R6.Nm,''), " +
            "Prod.PictNo,Prod.NoteNm, " +
            "Prod.Wdtu,Prod.Gr3 as okologisk, Prod.Free1, " +

            "FreeInf1.Val1, FreeInf1.Val11, FreeInf1.Val12,  " +
             "ISNULL(TXTX.txt, ISNULL(PCX.Descr, '')) as omraade2, " +
            "(Prod.EdFMT & 262144), Prod.Gr5, " +
            "ISNULL(TXT2_45.Txt, '') as gruppe45, " +
            "ISNULL(TXT3_45.Txt, '') as gruppe46 " +
            "FROM Prod   " +
            "LEFT OUTER JOIN ScanCd ON ScanCd.ProdNo= Prod.ProdNo " +
            "LEFT OUTER JOIN R6 ON R6.RNo= Prod.R6 " +
            "INNER JOIN FreeInf1 ON FreeInf1.ProdNo=Prod.ProdNo AND FreeInf1.InfCatNo=45 " +

            "LEFT OUTER JOIN Txt AS TXT2 ON Txt2.Lang = #8# AND Txt2.TxtTp=42 AND TXT2.TxtNo= Prod.Gr " +
            "LEFT OUTER JOIN Txt AS TXT3 ON Txt3.Lang = #8# AND Txt3.TxtTp= 72 AND TXT3.TxtNo= Prod.ProdPrG3 " +

            "LEFT OUTER JOIN Txt AS TXT4 ON Txt4.Lang = #8# AND Txt4.TxtTp= 36 AND TXT4.TxtNo= R6.Gr " +
            "LEFT OUTER JOIN Txt AS TXT5 ON Txt5.Lang = #8# AND Txt5.TxtTp= 38 AND TXT5.TxtNo= R6.Gr3 " +
            "LEFT OUTER JOIN Txt AS TXT6 ON Txt6.Lang = #8# AND Txt6.TxtTp= 37 AND TXT6.TxtNo= Prod.PrcatNo3 " +

            "LEFT OUTER JOIN ProdCat AS PC1 ON  PC1.PrCatNo=R6.Gr3 " +
            "LEFT OUTER JOIN ProdCat AS PC2 ON  PC2.PrCatNo=R6.Gr " +
            "LEFT OUTER JOIN ProdCat AS PCX ON  PCX.PrCatNo=Prod.PrcatNo2  " +
            "LEFT OUTER JOIN Txt AS TXTX ON TxtX.Lang = #8# AND TxtX.TxtTp= 38 AND TXTX.TxtNo= Prod.PrcatNo2 " +
            "LEFT OUTER JOIN Txt AS TXT2_45 ON TXT2_45.Lang = 45 AND TXT2_45.TxtTp=42 AND TXT2_45.TxtNo= Prod.Gr " +
            "LEFT OUTER JOIN Txt AS TXT3_45 ON TXT3_45.Lang = 45 AND TXT3_45.TxtTp= 72 AND TXT3_45.TxtNo= Prod.ProdPrG3 " +
           "### " +
           "WHERE Prod.Descr<> '' AND Prod.Gr7<7 AND FreeInf1.#9#=1 "; //FreeInf1.Val1<>0 AND FreeInf1.Val1<>10";

        public static string queryGetProductToProcess =
            "SELECT DISTINCT Prod.ProdNo FROM Prod WITH (NOLOCK) " +
            "INNER JOIN FreeInf1 WITH (NOLOCK) ON FreeInf1.ProdNo=Prod.ProdNo AND FreeInf1.InfCatNo=45 " +
           "### " +
           "WHERE Prod.Descr<> '' AND Prod.Gr7<7 AND FreeInf1.#9#=1 "; //FreeInf1.Val1<>0  AND FreeInf1.Val1<>10";

        public static string queryGetProductBOM =
            "SELECT DISTINCT Struct.SubProd,Struct.NoPerStr,Prod.Descr,ISNULL(Prod.WebPg2,''),Prod.Wdtu,Prod.Gr10,Struct.LnNo FROM Struct INNER JOIN Prod ON Prod.ProdNo=Struct.SubProd WHERE Struct.ProdNo='#1#' ORDER BY Struct.LnNo";

        public static string queryGetStock =
             "SELECT ISNULL(SUM((Stcbal.Bal + StcBal.StcInc - StcBal.ShpRsv)),0) FROM StcBal " +
             "INNER JOIN Prod ON stcBal.ProdNo = Prod.ProdNo " +
             "WHERE StcBal.Prodno='#1#' " + //and (Stcbal.ChDt= 0 OR Stcbal.ChDt >= #2#) " +
             "AND Stcbal.StcNo IN (#3#) ";

        public static string queryGetDiscounts =
             "SELECT DISTINCT  PrDcMat.SalePr,PrDcMat.MinNo,PrDcMat.FrDt,PrDcMat.ToDt,PrDcMat.PrTp FROM PrDcMat " +
             "INNER JOIN Prod ON Prod.ProdNo=PrDcMat.ProdNo " +
             "WHERE PrDcMat.ProdNo='#1#' AND PrDcMat.CustPrg2=11 AND (PrDcMat.MinNo>0 OR (PrDcMat.MinNo=0 AND PrDcMat.PrTp=3)) AND PrDcMat.SalePr>0 AND Prod.Gr7<7 " +
             "AND (PrDcMat.FrDt<= #2# OR PrDcMat.FrDt=0) AND (PrDcMat.ToDt >= #2# OR PrDcMat.ToDt= 0) ";

        public static string queryGetPrice =
            "SELECT DISTINCT SalePr FROM PrDcMat WHERE  PrDcMat.PrTp = #3# AND PrDcMat.MinNo<=1 " +
            "AND Prodno = '#1#' " +
            "AND SalePr>0 AND  (FrDt<= #2# OR FrDt=0) AND (ToDt >= #2# OR ToDt = 0)";

        public static string queryGetPriceB2B =
         "SELECT DISTINCT SalePr FROM PrDcMat WHERE  PrDcMat.PrTp = 19 AND PrDcMat.MinNo<=1 " +
         "AND Prodno = '#1#' " +
         "AND SalePr>0 AND  (FrDt<= #2# OR FrDt=0) AND (ToDt >= #2# OR ToDt = 0)";

        /*   public static string queryGetCategories =
              "SELECT DISTINCT ProdCat.PrCatNo, ISNULL(TXT5.txt, ISNULL(ProdCat.Descr, '')),0,1 FROM ProdCat " +
              "LEFT OUTER JOIN Txt AS TXT5 ON Txt5.Lang = 45 AND Txt5.TxtTp= 40 AND TXT5.TxtNo= ProdCat.PrCatNo " +
              "WHERE Descr<> '' AND PrCatNo >=1 AND PrCatNo <= 999 " +
              "UNION " +
              "SELECT DISTINCT ProdCat.PrCatNo, ISNULL(TXT4.txt, ISNULL(ProdCat.Descr,'')), ProdCat.MainPrC,2 FROM ProdCat " +
               "LEFT OUTER JOIN Txt AS TXT4 ON Txt4.Lang = 45 AND Txt4.TxtTp= 36 AND TXT4.TxtNo= ProdCat.PrCatNo " +
               "WHERE Descr<> '' AND PrCatNo >=1000 AND PrCatNo <= 1999";
        */
        /*   public static string queryGetTags =
               "SELECT DISTINCT  Txt.Txt  from Txt WHERE Txt.Lang = 45 AND Txt.TxtTp= 72 ";
        */

        public static string queryGetProductGrapes =
            "SELECT DISTINCT txt.txt,FreeInf1.Gr3 from FreeInf1 INNER JOIN Txt on Txt.TxtNo=FreeInf1.Gr3 AND Txt.lang = #8# AND Txt.Txttp=159 " +
            "WHERE FreeInf1.infcatno=48 AND FreeInf1.ProdNo='#1#'";

        public static string queryGetProductRelations =
          "SELECT FreeInf2.ProdNo FROM FreeInf2 INNER JOIN FreeInf1 ON FreeInf2.FreeInf1=FreeInf1.PK " +
          "WHERE FreeInf1.InfCatNo=78 AND FreeInf1.ProdNo= '#1#'";

        public static string queryGetProducerRelations =
                "SELECT DISTINCT FreeInf1.Gr12 FROM R6  " +
                "INNER JOIN FreeInf1 ON FreeInf1.InfCatNo=80 AND FreeInf1.R6=#1# " +
                "WHERE R6.Nm<>'' AND FreeInf1.Gr12<> R6.RNo ";
        public DBaccess()
        {
            connection = new SqlConnection(Utils.ReadConfigString("Connectionstring", ""));
        }
        public DBaccess(int mode)
        {
            if (mode == 1)
                connection = new SqlConnection(Utils.ReadConfigString("Connectionstring", ""));
            else
                connection = new SqlConnection(Utils.ReadConfigString("Connectionstring2", ""));
        }

        public void CloseAll()
        {
            if (connection.State == ConnectionState.Open)
                connection.Close();
        }

        public bool TestConnection(out string errmsg)
        {
            errmsg = "";

            SqlCommand command = new SqlCommand("SELECT TOP 1 * FROM Prod", connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 500
            };
            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();
                reader = command.ExecuteReader();
                if (reader.Read())
                {
                    ;
                }
            }
            catch (Exception ex)
            {
                errmsg = ex.Message;
                return false;
            }
            finally
            {
                // always call Close when done reading .
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }

            return true;

        }

        /// <summary>
        /// Get database date in Visma form (int)
        /// </summary>
        /// <param name="errmsg"></param>
        /// <returns></returns>
        public int GetCurrentVismaDate(out string errmsg)
        {
            int dt = 0;
            errmsg = "";

            SqlCommand command = new SqlCommand("SELECT GETDATE()", connection)
            {
                CommandType = CommandType.Text
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                if (reader.Read())
                {
                    DateTime t = reader.GetDateTime(0);
                    dt = t.Year * 10000 + t.Month * 100 + t.Day;
                }
            }
            catch (Exception ex)
            {
                errmsg = "GetCurrentVismaDate() - " + ex.Message;

                return 0;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }

            return dt;

        }


        public bool GetAttributeTermsForGrapes(ref List<ProductAttributeTerm> attributeTerms, int langNo, out string errmsg)
        {
            attributeTerms.Clear();
            errmsg = "";
            string lang = Utils.LangNoToString(langNo);

            string sql = $"SELECT DISTINCT txt,CAST(txtno as int) FROM txt WHERE txttp = 159 AND Lang={langNo}";

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    string name = reader.GetString(0);
                    int vid = reader.GetInt32(1);
                    attributeTerms.Add(new ProductAttributeTerm()
                    {
                        name = name,
                        slug = langNo != 45 ? Utils.SanitizeSlugNameNew(name) + "-" + lang : Utils.SanitizeSlugNameNew(name),
                        visma_id = vid,
                        lang = lang,
                        //translations = new Translations()
                    });
                }

            }
            catch (Exception ex)
            {
                errmsg = ex.Message;

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }

            return true;
        }

        public bool GetAttributeTermsForYear(ref List<ProductAttributeTerm> attributeTerms, int langNo, out string errmsg)
        {
            attributeTerms.Clear();
            errmsg = "";
            string lang = Utils.LangNoToString(langNo);

            string sql = "SELECT DISTINCT Prod.ProdTp4 FROM Prod WHERE ProdTp4 > 0 ORDER BY ProdTp4";

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    int y = reader.GetInt32(0);

                    string name = y == 1 ? "N.V." : y.ToString();

                    attributeTerms.Add(
                        new ProductAttributeTerm()
                        {
                            name = name,
                            slug = null, //langNo != 45 ? Utils.SanitizeSlugNameNew(name) + "-" + lang : Utils.SanitizeSlugNameNew(name),
                            visma_id = y,
                            lang = lang,
                            //translations = new Translations()
                        });// ;
                }

            }
            catch (Exception ex)
            {
                errmsg = ex.Message;

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }

            return true;
        }

        public bool GetAttributeTermsForProducer(ref List<ProductAttributeTerm> attributeTerms, bool activeOnly, int langNo, out string errmsg)
        {
            attributeTerms.Clear();
            errmsg = "";

            string lang = Utils.LangNoToString(langNo);

            string sql;

            if (langNo == 45)
            {
                if (activeOnly)
                    sql = "SELECT DISTINCT R6.Nm,R6.NoteNm,R6.Rno  FROM R6 " +
                            "INNER JOIN Prod ON Prod.R6=R6.RNo " +
                            "INNER JOIN FreeInf1 ON FreeInf1.ProdNo=Prod.ProdNo AND FreeInf1.InfCatNo=45 " +
                            "WHERE R6.Nm<>'' AND  Prod.Gr7 < 7 AND Prod.Descr <> '' AND FreeInf1.Val1 <> 0 " +
                            "ORDER BY R6.Nm";
                else
                    sql = "SELECT DISTINCT Nm,NoteNm,Rno FROM R6 WHERE Nm<>'' AND LTRIM(RTRIM(Inf2))=''	 ORDER BY Nm";
            }
            else
            {
                if (activeOnly)
                    sql = "SELECT DISTINCT R6.Nm,ISNULL(F2.NoteNm,''),R6.Rno  FROM R6 " +
                            "INNER JOIN Prod ON Prod.R6=R6.RNo " +
                            "INNER JOIN FreeInf1 F1 ON F1.ProdNo=Prod.ProdNo AND F1.InfCatNo=45 " +   // STILL USING DANISH NAMES!
                            "LEFT OUTER JOIN FreeInf1 F2 ON F2.R6=R6.Rno AND F2.InfCatNo=81 " +
                            "WHERE R6.Nm<>'' AND  Prod.Gr7 < 7 AND Prod.Descr <> '' AND F1.Val1 <> 0 " +
                            "ORDER BY R6.Nm";
                else
                    sql = "SELECT DISTINCT R6.Nm,ISNULL(F2.NoteNm,''),Rno FROM R6 " +                 // Descriptions are translated
                            "LEFT OUTER JOIN FreeInf1 F2 ON F2.R6=R6.Rno AND F2.InfCatNo=81 " +
                            "WHERE R6.Nm <> '' AND LTRIM(RTRIM(R6.Inf2))='' " +
                            "ORDER BY R6.Nm";
            }
            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    string name = reader.GetString(0).Replace("&", "&amp;");
                    string desc = Utils.ReadMemoFile(reader.GetString(1));
                    int vid = reader.GetInt32(2);
                    attributeTerms.Add(new ProductAttributeTerm()
                    {
                        name = name,
                        description = desc,
                        slug = langNo != 45 ? Utils.SanitizeSlugNameNew(name) + "-" + lang : Utils.SanitizeSlugNameNew(name),
                        visma_id = vid,
                        lang = lang,
                        //translations = new Translations()
                    });
                }

            }
            catch (Exception ex)
            {
                errmsg = ex.Message;

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }



            return true;
        }

        public bool GetAttributeTermsForRelatedProducers(int rno, ref List<int> relatedProducers,
                                                       out string errmsg)
        {
            relatedProducers.Clear();
            errmsg = "";

            string sql = queryGetProducerRelations.Replace("#1#", rno.ToString());

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    relatedProducers.Add(reader.GetInt32(0));
                }

            }
            catch (Exception ex)
            {
                errmsg = ex.Message;

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }

            return true;
        }


        public bool GetAttributeTermsForProducerToDelete(ref List<ProductAttributeTerm> attributeTermsToDelete,
                                                               out string errmsg)
        {
            attributeTermsToDelete.Clear();
            errmsg = "";

            string sql = "SELECT DISTINCT Nm,NoteNm,Rno FROM R6 WHERE Nm<>'' AND LTRIM(RTRIM(Inf2))<>'' ORDER BY Nm";

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    string name = reader.GetString(0).Replace("&", "&amp;");
                    string desc = Utils.ReadMemoFile(reader.GetString(1));
                    int vid = reader.GetInt32(2);
                    attributeTermsToDelete.Add(new ProductAttributeTerm() { name = name, description = desc, slug = Utils.SanitizeSlugNameNew(name), visma_id = vid, lang = "da" });
                    attributeTermsToDelete.Add(new ProductAttributeTerm() { name = name, description = desc, slug = Utils.SanitizeSlugNameNew(name), visma_id = vid, lang = "en" });
                }

            }
            catch (Exception ex)
            {
                errmsg = ex.Message;

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }

            return true;
        }

        public bool GetAttributeTermsForCountry(ref List<ProductAttributeTerm> attributeTerms, int langNo, out string errmsg)
        {
            attributeTerms.Clear();
            errmsg = "";

            string lang = Utils.LangNoToString(langNo);

            string sql = $"SELECT DISTINCT txt,CAST(txtno as int) from txt where TxtTp=38 and Lang = {langNo} and txt<>'' ";

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    string name = reader.GetString(0);
                    int vid = reader.GetInt32(1);
                    attributeTerms.Add(new ProductAttributeTerm()
                    {
                        name = name,
                        slug = langNo != 45 ? Utils.SanitizeSlugNameNew(name) + "-" + lang : Utils.SanitizeSlugNameNew(name),
                        visma_id = vid,
                        lang = lang,
                        //translations = new Translations()
                    });
                }

            }
            catch (Exception ex)
            {
                errmsg = ex.Message;

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }

            return true;
        }

        public bool GetAttributeTermsForRegion(ref List<ProductAttributeTerm> attributeTerms, ref List<RegionCountryRelation> regionCountryRelationList, int langNo, out string errmsg)
        {
            regionCountryRelationList.Clear();
            attributeTerms.Clear();
            errmsg = "";

            //            string sql = "SELECT DISTINCT txt from txt where TxtTp=36 and Lang = 45 ";

            string lang = Utils.LangNoToString(langNo);

            string sql = "SELECT DISTINCT txt36.txt,ISNULL(txt38.txt, ''),CAST(txt36.txtno as int) FROM txt AS txt36 " +
                        "LEFT JOIN r6 ON r6.gr = txt36.txtno " +
                        $"LEFT JOIN txt txt38 on txt38.TxtTp = 38 AND txt38.txtno = r6.gr3 AND txt38.lang={langNo} " +
                        $"WHERE txt36.TxtTp = 36 AND txt36.Lang = {langNo} ORDER BY txt36.txt ";

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    string name = reader.GetString(0);
                    string country = reader.GetString(1);
                    int vid = reader.GetInt32(2);
                    regionCountryRelationList.Add(new RegionCountryRelation() { Region = name, Country = country });
                    attributeTerms.Add(new ProductAttributeTerm()
                    {
                        name = name,
                        slug = langNo != 45 ? Utils.SanitizeSlugNameNew(name) + "-" + lang : Utils.SanitizeSlugNameNew(name),
                        visma_id = vid,
                        lang = lang,
                        //translations = new Translations()
                    });
                }

            }
            catch (Exception ex)
            {
                errmsg = ex.Message;

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }

            return true;
        }


        public bool GetAttributeTermsForType(ref List<ProductAttributeTerm> attributeTerms, int langNo, out string errmsg)
        {
            attributeTerms.Clear();
            errmsg = "";

            string sql = $"SELECT DISTINCT txt,CAST(txtno as int) from txt where TxtTp=72 and Lang = {langNo} ";

            string lang = Utils.LangNoToString(langNo);

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    string name = reader.GetString(0);
                    int vid = reader.GetInt32(1);
                    attributeTerms.Add(new ProductAttributeTerm()
                    {
                        name = name,
                        slug = langNo != 45 ? Utils.SanitizeSlugNameNew(name) + "-" + lang : Utils.SanitizeSlugNameNew(name),
                        visma_id = vid,
                        lang = lang,
                        //translations = new Translations()
                    });

                }

            }
            catch (Exception ex)
            {
                errmsg = ex.Message;

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }

            return true;
        }

        public bool GetAttributeTermsForVolume(ref List<ProductAttributeTerm> attributeTerms, int langNo, out string errmsg)
        {
            errmsg = "";
            attributeTerms.Clear();
            string lang = Utils.LangNoToString(langNo);

            //     string sql = "SELECT DISTINCT WdtU FROM Prod WHERE  WdtU>0 ORDER BY WdtU";
            string sql = "SELECT DISTINCT WdtU FROM Prod WHERE  WdtU>=0 ORDER BY WdtU";

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    decimal ll = reader.GetDecimal(0);
                    string name = Utils.DecimalToStringFloating(ll) + " l";
                    attributeTerms.Add(new ProductAttributeTerm()
                    {
                        name = name,
                        slug = null, //langNo != 45 ? Utils.SanitizeSlugNameNew(name) + "-" + lang : Utils.SanitizeSlugNameNew(name),
                        visma_id = ll == 0.0M ? 100000000 : Decimal.ToInt32(ll * 1000.0M),
                        lang = lang,
                        //translations = new Translations()
                    }) ;
                }

            }
            catch (Exception ex)
            {
                errmsg = ex.Message;

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }

            return true;
        }

        public bool GetTagsForAlcohol(ref List<ProductTag> tags, out string errmsg)
        {
            errmsg = "";

            string sql = "SELECT DISTINCT Free1 FROM Prod WHERE Free1>=0";

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    string name = Constants.AlcoholPrefix + Utils.DecimalToStringFloating(reader.GetDecimal(0)) + "%";
                    ProductTag tag_da = new ProductTag() { name = name, slug = Utils.SanitizeSlugNameNew(name), lang = "da" };

                    name = Constants.AlcoholPrefixEn + Utils.DecimalToStringFloating(reader.GetDecimal(0)) + "%";
                    ProductTag tag_en = new ProductTag() { name = name, slug = Utils.SanitizeSlugNameNew(name) + "-en", lang = "en" };

                   // tag_da.translations = new Translations() { nameforda = tag_da.name, nameforen = tag_en.name };
                 //   tag_en.translations = new Translations() { nameforda = tag_da.name, nameforen = tag_en.name };

                    tags.Add(tag_da);
                    if (Utils.ReadConfigInt32("DoTranslation", 0) > 0)
                        tags.Add(tag_en);
                }

            }
            catch (Exception ex)
            {
                errmsg = ex.Message;

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }

            return true;
        }


        public bool GetProductsToProcess(ref List<string> products, DateTime latestSyncTime, out string errmsg)
        {
            products.Clear();
            errmsg = "";

            string valField = Utils.ReadConfigString("FreeInf1Selection", "Val1");
            int vismaDt = latestSyncTime.Year * 10000 + latestSyncTime.Month * 100 + latestSyncTime.Day;
            int vismaTm = latestSyncTime.Hour * 100 + latestSyncTime.Minute;
            if (latestSyncTime.Year < 2000)
            {
                vismaDt = 0;
                vismaTm = 0;
            }
            string sql = queryGetProductToProcess.Replace("#9#", valField);
            if (vismaDt > 20000101)
            {
                sql = sql.Replace("###", "INNER JOIN PrDcMat WITH (NOLOCK) ON PrDcMat.ProdNo=Prod.ProdNo  AND PrDcMat.PrTp>0 ");
                sql += "AND ((Prod.ChDt=0) OR (Prod.ChDt > #1#) OR (Prod.ChDt = #1# AND Prod.ChTm >= #2#) OR (PrDcMat.ChDt > #1#) OR (PrDcMat.ChDt = #1# AND PrDcMat.ChTm >= #2#)) ";
            }
            sql = sql.Replace("###", "");
            sql = sql.Replace("#1#", vismaDt.ToString()).Replace("#2#", vismaTm.ToString());
            if (!string.IsNullOrWhiteSpace(TestProductNo))
                sql += $" AND Prod.ProdNo = '{TestProductNo}' ";
            sql += " ORDER BY Prod.ProdNo";
            Utils.WriteLog("DEBUG: " + sql);

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            SqlDataReader reader = null;
            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    products.Add(reader.GetString(0));
                }

            }
            catch (Exception ex)
            {
                errmsg = ex.Message;
                Utils.WriteLog(ex.StackTrace);

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }

            return true;
        }

        public bool GetProducts(SyncType mode, ref List<Product> products, DateTime latestSyncTime,
                                                    List<ProductAttribute> wooCommerceAttributes,
                                                    List<ProductTag> wooCommerceTags,
                                                    List<ProductCategory> wooCommerceCategories,
                                                    int langNo,
                                                    out string errmsg)
        {
            products.Clear();

            List<VismaProductDetail> productDetails = new List<VismaProductDetail>();
            errmsg = "";

            string lang = Utils.LangNoToString(langNo);

            string valField = Utils.ReadConfigString("FreeInf1Selection", "Val1");

            //int descriptionMaxLength = Utils.ReadConfigInt32("DescriptionMaxLength", 4096);

            if (string.IsNullOrWhiteSpace(TestProductNo) == false)
                latestSyncTime = DateTime.MinValue;

            int vismaDt = latestSyncTime.Year * 10000 + latestSyncTime.Month * 100 + latestSyncTime.Day;
            int vismaTm = latestSyncTime.Hour * 100 + latestSyncTime.Minute;
            if (latestSyncTime.Year < 2000)
            {
                vismaDt = 0;
                vismaTm = 0;
            }

            string sql = queryGetProductList.Replace("#9#", valField).Replace("#8#", langNo.ToString());
            if (langNo != 45)
            {
                sql = sql.Replace("Prod.WebPg2", "FI_EN.WebPg");
                sql = sql.Replace("###",
                    " LEFT OUTER JOIN FreeInf1 AS FI_EN ON FI_EN.ProdNo=Prod.ProdNo AND FI_EN.InfCatNo=84 " +
                    "###");
            }

            if (vismaDt > 20000101)
            {
                if (mode == SyncType.Products || mode == SyncType.Campaigns)
                {
                    sql = sql.Replace("###", "INNER JOIN PrDcMat WITH (NOLOCK) ON PrDcMat.ProdNo=Prod.ProdNo  AND PrDcMat.PrTp>0 ");
                    sql += "AND ((Prod.ChDt=0) OR (Prod.ChDt > #1#) OR (Prod.ChDt = #1# AND Prod.ChTm >= #2#) OR (PrDcMat.ChDt > #1#) OR (PrDcMat.ChDt = #1# AND PrDcMat.ChTm >= #2#)) ";
                }
                if (mode == SyncType.Stock)
                {
                    sql = sql.Replace("###", "INNER JOIN StcBal WITH (NOLOCK) ON StcBal.ProdNo=Prod.ProdNo ");
                    sql += "AND ((Stcbal.ChDt=0) OR (Stcbal.ChDt > #1#) OR (Stcbal.ChDt = #1# AND Stcbal.ChTm >= #2#)) ";
                }
            }
            sql = sql.Replace("###", "");
            sql = sql.Replace("#1#", vismaDt.ToString()).Replace("#2#", vismaTm.ToString());
            if (!string.IsNullOrWhiteSpace(TestProductNo))
                sql += $" AND Prod.ProdNo = '{TestProductNo}' ";

            sql += " ORDER BY Prod.ProdNo";

            Utils.WriteLog("DEBUG: " + sql);

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            SqlDataReader reader = null;
            string dbg = "";
            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    int idx = 0;
                    Product item = new Product
                    {
                        id = null,
                        sku = reader.GetString(idx++), // ProdNo
                        vismaUnitsPerProdNo = reader.GetInt32(idx++)
                    };
                    string descr = Utils.SanitizeName(reader.GetString(idx++).Trim());
                    string longTitle = reader.GetString(idx++).Trim();  //Prod.WebPg2 or FI_EN.WebPg

                    if (langNo == 44)
                        longTitle = longTitle.Replace(" - ØKO", " - BIO");
                    item.name = Utils.SanitizeName(longTitle);

                    item.meta_data = new List<ProductMeta>();

                    // 2021105 Shave off dangling '-' (smagekasser..)
                    longTitle = longTitle.Trim();
                    if (longTitle.EndsWith("-"))
                        longTitle = longTitle.Substring(0, longTitle.Length - 1).Trim();

                    ProductMeta meta = new ProductMeta() { key = "_original_title", value = descr };
                    item.meta_data.Add(meta);

                    item.lang = Utils.LangNoToString(langNo);

                    dbg = item.sku;
                    if (item.vismaUnitsPerProdNo == 0)
                        item.vismaUnitsPerProdNo = 1;

                    int productType = reader.GetInt32(idx++);  // Gr10 

                    int year = reader.GetInt32(idx++);  // ProdTp4 
                    string yearString = year > 0 ? (year == 1 ? "N.V." : year.ToString()) : "";
                    int itemPerColli = reader.GetInt32(idx++);  // ProdTp2

                    string group = reader.GetString(idx++).Trim();           // Prod.Gr -> Txt
                    string prodSubType = reader.GetString(idx++).Trim();           // Prod.ProdPrG3 -> Txt
                    string barCode = reader.GetString(idx++).Trim();
                    int m = barCode.IndexOf("@");
                    if (m != -1)
                        barCode = barCode.Substring(0, m);

                    string land = reader.GetString(idx++).Trim();    // =Prod.PrCatNo -> txt
                    string omraade = reader.GetString(idx++).Trim();    // =R6.Gr -> txt
                    string kommune = reader.GetString(idx++).Trim();    // =R6.Gr2 -> txt              
                    string producent = reader.GetString(idx++).Trim().Replace("&", "&amp;");    // = R6.Nm
                    string picture = reader.GetString(idx++).Trim();    // =   Prod.PictNo
                    string memoFile = reader.GetString(idx++).Trim();


                    //string exDescription = Utils.SanitizeName(reader.GetString(idx++)).Trim(); // Prod.WebPg2
                    //if (exDescription != "")
                    //    item.name = exDescription;

                    decimal d = reader.GetDecimal(idx++);

                    string volStringL = Utils.DecimalToStringFloating(d) + " l";
                    //Convert to ml

                    string volString = "";
                    if (d < 1.0M)
                    {
                        d *= 1000;
                        volString = Utils.DecimalToStringFloating(d) + " ml";
                    } 
                    else if (d >= 1.0M)
                    { 
                        volString = Utils.DecimalToStringFloating(d) + " l";
                    }   
                    // 20210810 - add volume to title/name
                    if (d > 0 && productType != 20) // Do not add vol if 'Smagekasse'
                        item.name += " " + volString;
                    item.name = item.name.Trim();

                    // 2021105 Shave off dangling '-' (smagekasser..)
                    if (item.name.EndsWith("-"))
                        item.name = item.name.Substring(0, item.name.Length - 1).Trim();

                    int n = reader.GetInt32(idx++);                         // Prod.Gr3  
                    bool okologisk = (n == 1);

                    // Actually not used anymore...
                    //item.slug = Utils.SanitizeSlugNameNew(item.name); //20210923 don't send slug to webshop

                    d = reader.GetDecimal(idx++);                           // Prod.Free1
                    string alcString = Utils.DecimalToStringFloating(d) + "%";

                    /*   int cat0 = reader.GetInt32(idx++);   // land (NY)                       // "ISNULL(R6.Gr3,0), ISNULL(R6.Gr,0), 0, FreeInf1.Val1, FreeInf1.Val11  "
                       int cat1 = reader.GetInt32(idx++);   // område (NY)             
                       int cat2 = reader.GetInt32(idx++);   // mark (ikke brugt)
                   */

                    int val1 = Decimal.ToInt32(reader.GetDecimal(idx++)); // FreeInf1.Val1
                    item.status = val1 == 1 ? "publish" : "draft";
                    item.stock_quantity = Decimal.ToInt32(reader.GetDecimal(idx++)); // max-stock..FreeInf1.Val11

                    //item.stock_status = "InStock" , "OutOfStock"

                    int maxQty = Decimal.ToInt32(reader.GetDecimal(idx++)); // FreeInf1.Val12
                    if (maxQty > 0 || maxQty == -1)
                    {
                        ProductMeta meta2 = new ProductMeta() { key = "_wc_max_qty_product", value = (maxQty == -1) ? 0 : maxQty };
                        item.meta_data.Add(meta2);
                    }
                    string omraadeOverrule = reader.GetString(idx++).Trim();
                    if (omraadeOverrule != "")
                        omraade = omraadeOverrule;

                    bool hasBOM = reader.GetInt32(idx++) > 0;



                    bool noDiscount = reader.GetInt32(idx++) > 0;

                    string group_danish = reader.GetString(idx++).Trim();           // Prod.Gr -> Txt (lang=45)
                    string prodSubType_danish = reader.GetString(idx++).Trim();           // Prod.ProdPrG3 -> Txt (lang=45)

                    if (group == "")
                        group = group_danish;
                    if (prodSubType == "")
                        prodSubType = prodSubType_danish;

                    item.images = new List<ProductImage>();
                    if (picture != "" && Utils.ReadConfigInt32("SendImages", 0) > 0)
                        item.images.Add(new ProductImage() { src = picture, name = Path.GetFileNameWithoutExtension(picture), alt = item.name });

                    if (memoFile != "")
                        item.description = Utils.ReadMemoFile(memoFile);

                    item.categories = new List<ProductCategoryLine>();
                    /* 20200901 - categories not used..   
                   if (land != "")
                       item.categories.Add(new ProductCategoryLine() {  name = land, slug = Utils.SanitizeSlugName(land) });
                   if (omraade != "")
                       item.categories.Add(new ProductCategoryLine() {  name = omraade, slug = Utils.SanitizeSlugName(omraade) });
                   */

                    if (noDiscount)
                    {
                        ProductCategory noDiscountCategory = wooCommerceCategories.FirstOrDefault(p => p.name == (langNo == 45 ? Constants.NoDiscount : Constants.NoDiscountEn) && p.lang == Utils.LangNoToString(langNo));
                        if (noDiscountCategory != null)
                            //item.categories.Add(new ProductCategoryLine() { id = noDiscountCategory.id, name = Constants.NoDiscount, slug = noDiscountCategory.slug /*, slug = langNo == 45 ? Utils.SanitizeSlugNameNew(Constants.NoDiscount)  : Utils.SanitizeSlugNameNew(Constants.NoDiscount) + "-" + lang*/});
                            item.categories.Add(new ProductCategoryLine() { id = noDiscountCategory.id, name = noDiscountCategory.name, slug = noDiscountCategory.slug /*, slug = langNo == 45 ? Utils.SanitizeSlugNameNew(Constants.NoDiscount)  : Utils.SanitizeSlugNameNew(Constants.NoDiscount) + "-" + lang*/});
                        else // add anyway...(without slug?)
                            item.categories.Add(new ProductCategoryLine() { name = (langNo == 45 ? Constants.NoDiscount : Constants.NoDiscountEn)/*, slug = (langNo == 45 ? Utils.SanitizeSlugNameNew(Constants.NoDiscount)  : Utils.SanitizeSlugNameNew(Constants.NoDiscount)) + "-" + lang*/ });
                    }

                    // Added 20210204 - moved from attribute
                    if (okologisk)
                    {
                        ProductCategory OekologiskCategory = langNo == 45
                            ? wooCommerceCategories.FirstOrDefault(p => p.name == Constants.CategoryOekologisk && p.lang == Utils.LangNoToString(langNo))
                            : wooCommerceCategories.FirstOrDefault(p => (p.name == Constants.CategoryOekologiskEn || p.name == Constants.Old_CategoryOekologiskEn)
                            && p.lang == Utils.LangNoToString(langNo));
                        if (OekologiskCategory != null)
                            // item.categories.Add(new ProductCategoryLine() { id = OekologiskCategory.id, name = (langNo == 45 ? Constants.CategoryOekologisk : Constants.CategoryOekologiskEn), slug = OekologiskCategory.slug /*, slug = langNo == 45 ? Utils.SanitizeSlugNameNew(Constants.CategoryOekologiskSlug)  : Utils.SanitizeSlugNameNew(Constants.CategoryOekologiskSlug) + "-" + lang*/ });
                            item.categories.Add(new ProductCategoryLine() { id = OekologiskCategory.id, name = OekologiskCategory.name, slug = OekologiskCategory.slug /*, slug = langNo == 45 ? Utils.SanitizeSlugNameNew(Constants.CategoryOekologiskSlug)  : Utils.SanitizeSlugNameNew(Constants.CategoryOekologiskSlug) + "-" + lang*/ });
                        else // add anyway...
                            item.categories.Add(new ProductCategoryLine() { name = (langNo == 45 ? Constants.CategoryOekologisk : Constants.CategoryOekologiskEn)/*, slug = langNo == 45 ? Utils.SanitizeSlugNameNew(Constants.CategoryOekologiskSlug)  : Utils.SanitizeSlugNameNew(Constants.CategoryOekologiskSlug) + "-" + lang*/ });
                    }

                    if (hasBOM)
                    {
                        ProductCategory bomCategory = wooCommerceCategories.FirstOrDefault(p => (langNo == 45 ? p.name == Constants.ProductWithBom : p.name == Constants.ProductWithBomEn) && p.lang == Utils.LangNoToString(langNo));
                        if (bomCategory != null)
                            //item.categories.Add(new ProductCategoryLine() { id = bomCategory.id, name = (langNo == 45 ?  Constants.ProductWithBom :  Constants.ProductWithBomEn)/*, slug = langNo == 45 ? Utils.SanitizeSlugNameNew(Constants.ProductWithBom)  : Utils.SanitizeSlugNameNew(Constants.ProductWithBom) + "-" + lang*/ });
                            item.categories.Add(new ProductCategoryLine() { id = bomCategory.id, name = bomCategory.name, slug = bomCategory.slug/*, slug = langNo == 45 ? Utils.SanitizeSlugNameNew(Constants.ProductWithBom)  : Utils.SanitizeSlugNameNew(Constants.ProductWithBom) + "-" + lang*/ });
                    }

                    products.Add(item);

                    productDetails.Add(new VismaProductDetail()
                    {
                        ProdNo = item.sku,
                        Volume = volString,
                        VolumeL = volStringL,
                        Alcohol = alcString,
                        Country = land,
                        Region = omraade,
                        Producer = producent,
                        ProductType = prodSubType,
                        Eco = okologisk ? "Ja" : "", // or "Nej" ?? 
                        Year = yearString,
                        County = kommune,
                        HasBOM = hasBOM,
                        NoDiscount = noDiscount

                    });

                }
            }
            catch (Exception ex)
            {
                errmsg = $"GetProducts({dbg}) - {ex.Message}";
                Utils.WriteLog(ex.StackTrace);

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }

            Utils.WriteLog($"Fetching details for {products.Count} products...");
            int counter = 0;
            foreach (Product product in products)
            {
                //Utils.WriteLog(product.vismaProdNo);
                //if (mode == SyncType.Products || mode == SyncType.Campaigns || mode == SyncType.Stock)
                //{

                if ((++counter % 100) == 0)
                    Utils.WriteLog($"{counter} Products fetched..");
                decimal price = 0.0M;
                if (GetPrice(product.sku, ref price, false, out errmsg) == false)
                {
                    Utils.WriteLog($"ERROR: GetPrice() - {errmsg}");
                    return false;
                }


                product.regular_price = price;// * product.vismaUnitsPerProdNo;

                // EURO price
                price = 0.0M;
                if (GetPrice(product.sku, ref price, true, out errmsg) == false)
                {
                    Utils.WriteLog($"ERROR: GetPrice() - {errmsg}");
                    return false;
                }
                product.vismaPriceEUR = price;
                if (price > 0.0M)
                {
                    if (product.meta_data == null)
                        product.meta_data = new List<ProductMeta>();
                    product.meta_data.Add(new ProductMeta()
                    {
                        key = "_regular_currency_prices",
                        value = new ProductMetaEur()
                        {
                            EUR = Utils.DecimalToString(price)
                        }
                    });
                }

                string longDescr = "";
                if (GetLongDescription(product.sku, ref longDescr, langNo, out errmsg) == false)
                {
                    Utils.WriteLog($"ERROR: GetLongDescription() - {errmsg}");
                    return false;
                }

                product.description = longDescr;


                //}

                //if (mode == SyncType.Stock || mode == SyncType.Products)
                // {
                int stockCount = 0;
                if (GetStock(product.sku, 0, ref stockCount, out errmsg) == false)
                    return false;
                int extraStockCount = 0;
                if (GetWebReservedStock(product.sku, ref extraStockCount, out errmsg))
                    stockCount += extraStockCount;

                if (product.stock_quantity == 0 || stockCount < product.stock_quantity.Value)
                    product.stock_quantity = stockCount;



                //}

                try
                {

                    VismaProductDetail detail = productDetails.FirstOrDefault(p => p.ProdNo == product.sku);

                    if (detail.HasBOM)
                    {
                        List<VismaBomItem> bomItems = new List<VismaBomItem>();
                        if (GetProductBOM(product.sku, langNo, ref bomItems, out errmsg) == false)
                            return false;
                        //                        string prodList = "";
                        int minStock = 9999;

                        List<ProductBomItem> bomList = new List<ProductBomItem>();
                        foreach (VismaBomItem bomItem in bomItems)
                        {
                            detail.VismaBomItems.Add(bomItem); // not yet used..

                            //                              if (prodList != "")
                            //                                prodList += ",";
                            //                          prodList += bomItem.ProdNo;

                            int stock = 0;
                            GetStock(bomItem.ProdNo, 0, ref stock, out errmsg);
                            if (stock < minStock)
                                minStock = stock;

                            bomList.Add(new ProductBomItem() { sku = bomItem.ProdNo, quantity = bomItem.Qty, title = bomItem.Descr });
                        }

                        // ProductMeta meta = new ProductMeta() { key = "_in_box", value = prodList };
                        if (product.meta_data == null)
                            product.meta_data = new List<ProductMeta>();
                        product.meta_data.Add(new ProductMeta() { key = "_in_box", value = bomList });

                        product.stock_quantity = minStock;

                    }

                    List<string> categoryTags = new List<string>();

                    product.attributes = new List<ProductAttributeLine>();
                    product.tags = new List<ProductTagLine>();

                    List<string> grapes = new List<string>();
                    if (GetProductGrapes(product.sku, ref grapes, langNo, out errmsg) == false)
                    {
                        Utils.WriteLog($"ERROR: GetProductGrapes() - {errmsg}");
                        return false;
                    }
                    if (grapes.Count == 0 && langNo == 44)
                    {
                        if (GetProductGrapes(product.sku, ref grapes, 45, out errmsg) == false)
                        {
                            Utils.WriteLog($"ERROR: GetProductGrapes() - {errmsg}");
                            return false;
                        }
                    }
                    if (wooCommerceAttributes != null)
                    {
                        if (grapes.Count > 0 && detail.HasBOM == false)
                        {

                            ProductAttribute grapeAttribute = wooCommerceAttributes.Find(p => p.name == Constants.AttributeNameGrape);
                            product.attributes = new List<ProductAttributeLine>();
                            ProductAttributeLine a = new ProductAttributeLine() { id = grapeAttribute.id, name = grapeAttribute.name, visible = true, options = grapes };
                            product.attributes.Add(a);

                            //  Utils.WriteLog($"Grapes set for products {product.sku}.");
                        }

                        if (detail.Country != "" && detail.HasBOM == false)
                        {
                            ProductAttribute countryAttribute = wooCommerceAttributes.FirstOrDefault(p => p.name == Constants.AttributeNameCountry);
                            ProductAttributeLine a = new ProductAttributeLine() { id = countryAttribute.id, name = countryAttribute.name, visible = true, options = new List<string> { detail.Country } };


                            product.attributes.Add(a);
                            //  Utils.WriteLog($"Country {detail.Country} set for products {product.sku}.");
                        }

                        if (detail.Region != "" && detail.HasBOM == false)
                        {
                            ProductAttribute regionAttribute = wooCommerceAttributes.FirstOrDefault(p => p.name == Constants.AttributeNameRegion);
                            ProductAttributeLine a = new ProductAttributeLine() { id = regionAttribute.id, name = regionAttribute.name, visible = true, options = new List<string> { detail.Region } };
                            product.attributes.Add(a);
                            //    Utils.WriteLog($"Region set for products {product.sku}.");
                        }

                        /*       if (detail.County != "" && detail.HasBOM == false)
                               {
                                   ProductAttribute countyAttribute = wooCommerceAttributes.FirstOrDefault(p => p.name == Constants.AttributeNameCounty);
                                   ProductAttributeLine a = new ProductAttributeLine() { id = countyAttribute.id, name = countyAttribute.name, visible = true, options = new List<string> { detail.County } };
                                   product.attributes.Add(a);
                                   //   Utils.WriteLog($"County set for products {product.sku}.");
                               }*/

                        if (detail.Producer != "" && detail.HasBOM == false)
                        {
                            ProductAttribute producerAttribute = wooCommerceAttributes.FirstOrDefault(p => p.name == Constants.AttributeNameProducer);
                            ProductAttributeLine a = new ProductAttributeLine() { id = producerAttribute.id, name = producerAttribute.name, visible = true, options = new List<string> { detail.Producer.Replace("&amp;", "&") } };          // WHAAAAT???
                            product.attributes.Add(a);

                        }

                        if (detail.ProductType != "" && detail.HasBOM == false)
                        {
                            ProductAttribute countyAttribute = wooCommerceAttributes.FirstOrDefault(p => p.name == Constants.AttributeNameType);
                            ProductAttributeLine a = new ProductAttributeLine() { id = countyAttribute.id, name = countyAttribute.name, visible = true, options = new List<string> { detail.ProductType } };
                            product.attributes.Add(a);
                            //   Utils.WriteLog($"ProductType set for products {product.sku}.");
                        }
                         
                        // Disabled 20210204 - moved to category
                        /* if (detail.Eco != "" && detail.HasBOM == false)
                         {
                             ProductAttribute ecoAttribute = wooCommerceAttributes.FirstOrDefault(p => p.name == Constants.AttributeNameEcologic);
                             ProductAttributeLine a = new ProductAttributeLine() { id = ecoAttribute.id, name = ecoAttribute.name, visible = true, options = new List<string> { detail.Eco } };
                             product.attributes.Add(a);
                             //   Utils.WriteLog($"Eco set for products {product.sku}.");
                         }*/

                        if (detail.Year != "" && detail.HasBOM == false)
                        {
                            ProductAttribute yearAttribute = wooCommerceAttributes.FirstOrDefault(p => p.name == Constants.AttributeNameYear);
                            ProductAttributeLine a = new ProductAttributeLine() { id = yearAttribute.id, name = yearAttribute.name, visible = true, options = new List<string> { detail.Year } };
                            product.attributes.Add(a);
                            //    Utils.WriteLog($"Year set for products {product.sku}.");
                        }

                        if (detail.VolumeL != "" && detail.HasBOM == false)
                        {
                            ProductAttribute volumeAttribute = wooCommerceAttributes.FirstOrDefault(p => p.name == Constants.AttributeNameVolume);
                            ProductAttributeLine a = new ProductAttributeLine() { id = volumeAttribute.id, name = volumeAttribute.name, visible = true, options = new List<string> { detail.VolumeL } };
                          //  Utils.WriteLog($"Volume {detail.VolumeL} set for products {product.sku} {volumeAttribute.id}");
                            product.attributes.Add(a);
                            //    Utils.WriteLog($"Volume set for products {product.sku}.");
                        }
                    }
                    if (wooCommerceTags != null)
                    {

                        /*
                        if (detail.Year != "")
                        {

                            ProductTag yearTag = wooCommerceTags.FirstOrDefault(p => p.name == Constants.YearPrefix + detail.Year);
                            ProductTagLine t;
                            if (yearTag != null)
                            {
                                t = new ProductTagLine() { id = yearTag.id, name = yearTag.name, slug = yearTag.slug };                               
                            }
                            else // do create anyway...
                            {
                                t = new ProductTagLine() { name = Constants.YearPrefix + detail.Year, slug = Utils.SanitizeSlugName(Constants.YearPrefix + detail.Year) };
                            }
                            product.tags.Add(t);
                            Utils.WriteLog($"Year set for products {product.sku}.");
                        }
                        */



                        /*if (detail.Volume != "")
                        {
                            ProductTag volumeTag = wooCommerceTags.FirstOrDefault(p => p.name == Constants.VolumePrefix + detail.Volume);
                            ProductTagLine t;
                            if (volumeTag != null)
                            {
                                t = new ProductTagLine() { id = volumeTag.id, name = volumeTag.name, slug = volumeTag.slug };
                            }
                            else // do create anyway...
                            {
                                t = new ProductTagLine() { name = Constants.VolumePrefix + detail.Volume, slug = Utils.SanitizeSlugName(Constants.VolumePrefix + detail.Volume) };
                            }
                            product.tags.Add(t);
                            Utils.WriteLog($"Volume set for products {product.sku}.");
                        }*/

                        if (detail.Alcohol != "" && detail.HasBOM == false)
                        {
                            string prefix = langNo != 45 ? Constants.AlcoholPrefixEn : Constants.AlcoholPrefix;
                            string old_prefix = langNo != 45 ? Constants.Old_AlcoholPrefixEn : Constants.Old_AlcoholPrefix;
                            ProductTag alcoholTag = wooCommerceTags.FirstOrDefault(p =>
                                ((p.name == prefix + detail.Alcohol) || (p.name == old_prefix + detail.Alcohol))
                                && p.lang == lang);
                            ProductTagLine t;
                            if (alcoholTag != null)
                            {
                                if (alcoholTag.name.StartsWith(old_prefix))
                                    alcoholTag.name.Replace(old_prefix, prefix);  //update to new prefix

                                t = new ProductTagLine() { id = alcoholTag.id, name = alcoholTag.name/*, slug = alcoholTag.slug*/ };
                            }
                            else // do create anyway...
                            {
                                t = new ProductTagLine() { name = prefix + detail.Alcohol/*, slug = langNo != 45 ? Utils.SanitizeSlugNameNew(Constants.AlcoholPrefix + detail.Alcohol) + "-" + lang : Utils.SanitizeSlugNameNew(Constants.AlcoholPrefix + detail.Alcohol)*/ };
                            }
                            product.tags.Add(t);
                            //  Utils.WriteLog($"Alcohol set for products {product.sku}.");
                        }

                        /* if (detail.NoDiscount)
                         {
                             ProductTag nodiscountTag = wooCommerceTags.FirstOrDefault(p => p.name == Constants.NoDiscount);
                             ProductTagLine t;
                             if (nodiscountTag != null)
                             {
                                 t = new ProductTagLine() { id = nodiscountTag.id, name = nodiscountTag.name, slug = nodiscountTag.slug };
                             }
                             else // do create anyway...
                             {
                                 t = new ProductTagLine() { name = Constants.NoDiscount, slug = Utils.SanitizeSlugName(Constants.NoDiscount) };
                             }
                             product.tags.Add(t);
                         }*/

                        //  }

                        List<string> relatedProducts = new List<string>();
                        if (GetProductRelations(product.sku, ref relatedProducts, out errmsg) == false)
                        {
                            Utils.WriteLog($"ERROR: GetProductRelations() - {errmsg}");
                            return false;
                        }

                        product.vismaRelatedProduct = relatedProducts;

                        //product.description = "";
                        //string longDescr = "";
                        //if (GetLongDescription(product.vismaProdNo, ref longDescr, out errmsg) == true)
                        //    product.description = Utils.SanitizeDescription(longDescr, true);
                        //product.description += "\n";

                        foreach (ProductImage img in product.images)
                        {
                            //Utils.WriteLog($"Image {img.src}..");

                            img.src = img.src.ToLower().Replace("v:", Utils.ReadConfigString("v-drive", "v:"));
                            img.src = img.src.ToLower().Replace("f:", Utils.ReadConfigString("f-drive", "f:"));
                            img.src = img.src.ToLower().Replace("g:", Utils.ReadConfigString("g-drive", "g:"));

                            if (Utils.ReadConfigString("ForceImagePath", "") != "")
                            {
                                if (File.Exists(Utils.ReadConfigString("ForceImagePath", "") + Path.GetFileName(img.src)))
                                    img.src = Utils.ReadConfigString("ForceImagePath", "") + Path.GetFileName(img.src);
                            }

                            if (File.Exists(img.src) == false)
                            {
                                product.images.Remove(img);
                                break;
                            }
                            else
                            {
                                //Utils.WriteLog("Image: " + product.image_url);
                                string newName = "";
                                if (Utils.ReadConfigString("CopyImageFolder", "") != "")
                                {
                                    // newName = System.Web.HttpUtility.UrlEncode(Path.GetFileName(product.image_url)).Replace("+","-");
                                    newName = Utils.SanitizeUrl(Path.GetFileName(img.src));
                                    try
                                    {
                                        string newPath = Utils.ReadConfigString("CopyImageFolder", "") + @"\" + newName;
                                        if (Utils.ReadConfigInt32("SkipImageCopy", 0) == 0)
                                        {
                                            Utils.WriteLog("Copy to : " + Utils.ReadConfigString("CopyImageFolder", "") + @"\" + Path.GetFileName(img.src));
                                            File.Copy(img.src, newPath, true);

                                            /*   if (Utils.ReadConfigInt32("LimitImageSize", 0) > 0)
                                               {
                                                   Utils.ResampleToSize(newPath, Utils.ReadConfigInt32("LimitImageSize", 0));
                                               }*/
                                            Utils.WriteLog("Image: " + img.src);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Utils.WriteLog("Exception copy: " + ex.Message);
                                    }
                                }
                                img.src = newName != "" ? Utils.ReadConfigString("ImageURL", "") + newName.ToUpper()
                                                                  : Utils.ReadConfigString("ImageURL", "") + Path.GetFileName(img.src).ToUpper();
                            }
                            Utils.WriteLog($"Image {img.src}.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Utils.WriteLog($"Error - exception - {ex.Message}");
                    errmsg = ex.Message;
                    return false;
                }

            }

            return true;
        }


        public bool GetProductBOM(string prodNo, int langNo, ref List<VismaBomItem> vismaBomItems, out string errmsg)
        {
            errmsg = "";
            vismaBomItems.Clear();

            string sql = queryGetProductBOM.Replace("#1#", prodNo);

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    string p = reader.GetString(0).Trim();
                    int qty = Decimal.ToInt32(reader.GetDecimal(1));
                    if (qty == 0)
                        qty = 1;
                    string descr = reader.GetString(2).Trim();
                    string descr1 = reader.GetString(3).Trim();
                    decimal vol = reader.GetDecimal(4);
                    int productType = reader.GetInt32(5);
                    if (descr1 != "")
                        descr = descr1;


                    if (langNo == 44)
                        descr = descr.Replace(" - ØKO", " - BIO");
                    string volString = "";
                    if (vol < 1.0M)
                    {
                        vol *= 1000;
                        volString = Utils.DecimalToStringFloating(vol) + " ml";
                    }
                    else if (vol >= 1.0M)
                    {
                        volString = Utils.DecimalToStringFloating(vol) + " l";
                    }

                    
                    // 20210810 - add volume to title/name
                    if (vol > 0 && productType != 20) // Do not add vol if 'Smagekasse'
                        descr += " " + volString;
                    vismaBomItems.Add(new VismaBomItem() { ProdNo = p, Qty = qty, Descr = descr });
                }
            }
            catch (Exception ex)
            {
                errmsg = "GetProducerBOM() - " + ex.Message + " " + sql;

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }

            return true;

        }

        public bool GetProductScores(string prodNo, ref string scoresText, out string errmsg)
        {
            errmsg = "";
            scoresText = "";

            string sql = $"SELECT DISTINCT Txt.txt, FreeInf1.Val1, FreeInf1.PK FROM Txt WITH (NOLOCK) INNER JOIN FreeInf1 WITH (NOLOCK) ON FreeInf1.Gr2 = Txt.TxtNo WHERE Txt.Lang = 45 AND Txt.TxtTp = 158 AND FreeInf1.InfcatNo = 47 AND FreeInf1.ProdNo = '{prodNo}' ORDER BY FreeInf1.PK ";

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    string s = reader.GetString(0);
                    int score = Decimal.ToInt32(reader.GetDecimal(1));

                    if (s.IndexOf("#Stjerner") != -1)
                        scoresText += string.Format("{0} Stjerner {1}\n", score, s.Replace("#Stjerner", ""));
                    else if (s.IndexOf("#Point") != -1)
                        scoresText += string.Format("{0} Point {1}\n", score, s.Replace("#Point", ""));
                    else
                        scoresText += string.Format("{0}\n", score, s.Replace("#", score.ToString() + " "));
                }
            }
            catch (Exception ex)
            {
                errmsg = "GetProductScores() - " + ex.Message + " " + sql;

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }

            return true;

        }

        public bool UpdateDeletedProduct(string prodNo, bool forced, out string errmsg)
        {
            errmsg = "";


            string sql = $"UPDATE FreeInf1 SET Val1=10 WHERE InfCatNo=45 AND ProdNo='{prodNo}' ";

            if (forced == false)
                sql += " AND Val1=9";
            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text
            };

            Utils.WriteLog(sql);

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                errmsg = "UpdateDeletedProduct() - " + ex.Message + " " + sql;

                return false;
            }
            finally
            {
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }
            return true;


        }

        public bool GetProductsToDelete(ref List<string> prodNoList, DateTime latestSyncTime, out string errmsg)
        {
            errmsg = "";
            prodNoList.Clear();
            int vismaDt = latestSyncTime.Year * 10000 + latestSyncTime.Month * 100 + latestSyncTime.Day;
            int vismaTm = latestSyncTime.Hour * 100 + latestSyncTime.Minute;
            if (latestSyncTime.Year < 2000)
            {
                vismaDt = 0;
                vismaTm = 0;
            }

            string valField = Utils.ReadConfigString("FreeInf1Selection", "Val1");

            string sql = queryGetDisabledProductList.Replace("#9#", valField);

            if (latestSyncTime != DateTime.MinValue)
                sql += " AND ((Prod.ChDt=0) OR (Prod.ChDt > #1#) OR (Prod.ChDt = #1# AND Prod.ChTm >= #2#)) ";
            sql = sql.Replace("#1#", vismaDt.ToString()).Replace("#2#", vismaTm.ToString());

            // Utils.WriteLog("DEBUG: " + sql);
            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    prodNoList.Add(reader.GetString(0));
                }
            }
            catch (Exception ex)
            {
                errmsg = "GetProductsToDelete() - " + ex.Message + " " + sql;

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }

            return true;

        }

        public bool GetProductsZeroStock(ref List<string> prodNoList, out string errmsg)
        {
            prodNoList.Clear();
            errmsg = "";

            string valField = Utils.ReadConfigString("FreeInf1Selection", "Val1");
            string stockList = Utils.ReadConfigString("StockList", "1,2");

            string sql = " SELECT DISTINCT Prod.ProdNo,ISNULL(SUM((Stcbal.Bal + StcBal.StcInc - StcBal.ShpRsv)),0)  FROM Prod WITH (NOLOCK) " +
            "INNER JOIN FreeInf1 WITH(NOLOCK) ON FreeInf1.ProdNo = Prod.ProdNo AND FreeInf1.InfCatNo = 45  " +
            "INNER JOIN StcBal ON StcBal.ProdNo = Prod.ProdNo " +
            "WHERE FreeInf1.#9# = 1 AND Prod.Gr10 <> 20 AND StcBal.StcNo IN (#1#) " +
            "GROUP BY Prod.ProdNo";

            sql = sql.Replace("#1#", stockList).Replace("#9#", valField);

            Utils.WriteLog("DEBUG: " + sql);
            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    string prodNo = reader.GetString(0);
                    decimal stock = reader.GetDecimal(1);
                    if (stock == 0.0M)
                        prodNoList.Add(prodNo);
                }
            }
            catch (Exception ex)
            {
                errmsg = "GetProductsZeroStock() - " + ex.Message + " " + sql;

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }


            return true;
        }
        public bool GetPrice(string prodNo, ref decimal price, bool euPrice, out string errmsg)
        {
            int prTp = euPrice ? Utils.ReadConfigInt32("PriceTypeEUR", 20) : Utils.ReadConfigInt32("PriceType", 19);
            errmsg = "";
            price = 0.0M;
            int vismaNow = DateTime.Now.Year * 10000 + DateTime.Now.Month * 100 + DateTime.Now.Day;
            string sql = queryGetPrice.Replace("#1#", prodNo.Replace("'", "''")).Replace("#2#", vismaNow.ToString()).Replace("#3#", prTp.ToString());

            if (euPrice)
                sql += " AND Cur=999 ";

            // Utils.WriteLog($"DEBUG: {sql}");
            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                if (reader.Read())
                {
                    price = reader.GetDecimal(0);
                    //                    price *= 1.25M;
                }
            }
            catch (Exception ex)
            {
                errmsg = "GetPrice() - " + ex.Message + " " + sql;

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }

            return true;
        }

        // Not used..
        public bool GetPriceB2B(string prodNo, int prodGr, ref decimal price, out string errmsg)
        {
            int prTp = Utils.ReadConfigInt32("PriceTypeB2B", 19);
            errmsg = "";
            price = 0.0M;
            int vismaNow = DateTime.Now.Year * 10000 + DateTime.Now.Month * 100 + DateTime.Now.Day;
            string sql = queryGetPriceB2B.Replace("#1#", prodNo.Replace("'", "''")).Replace("#2#", vismaNow.ToString()).Replace("#3#", prTp.ToString());

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                if (reader.Read())
                {
                    price = reader.GetDecimal(0);
                    //                        price *= 1.25M;
                }
            }
            catch (Exception ex)
            {
                errmsg = "GetPrice() - " + ex.Message + " " + sql;

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }

            return true;
        }

        public bool GetStock(string prodNo, int specificStcNo, ref int stock, out string errmsg)
        {
            errmsg = "";
            stock = 0; ;

            string sql = queryGetStock.Replace("#1#", prodNo.Replace("'", "''"));

            string stockList = Utils.ReadConfigString("StockList", "1,2");
            if (specificStcNo > 0)
                stockList = specificStcNo.ToString();
            sql = sql.Replace("#3#", stockList);
            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                if (reader.Read())
                {
                    stock = Decimal.ToInt32(reader.GetDecimal(0));
                }
            }
            catch (Exception ex)
            {
                errmsg = "GetStock() - " + ex.Message;
                Utils.WriteLog("ERROR: " + errmsg);
                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }

            return true;
        }

        public bool GetWebReservedStock(string prodNo, ref int stock, out string errmsg)
        {
            errmsg = "";
            stock = 0;

            int ordNoReservations = Utils.ReadConfigInt32("OrdNoReservations", 35918);

            string sql = $"SELECT NoInvoAb FROM OrdLn WHERE OrdNo={ordNoReservations} AND  ProdNo='{prodNo.Replace("'", "''")}' AND NoInvoAb>0 ";

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                if (reader.Read())
                {
                    stock = Decimal.ToInt32(reader.GetDecimal(0));
                }
            }
            catch (Exception ex)
            {
                errmsg = "GetWebReservedStock() - " + ex.Message;
                Utils.WriteLog("ERROR: " + errmsg);
                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }

            return true;
        }

        public bool GetDiscounts(string prodNo, ref List<VismaDiscount> discounts, out string errmsg)
        {

            int dt = GetCurrentVismaDate(out errmsg);
            string sql = queryGetDiscounts.Replace("#1#", prodNo.Replace("'", "''")).Replace("#2#", dt.ToString());

            //  Utils.WriteLog(sql);

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text
            };
            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    int idx = 0;
                    decimal discount = reader.GetDecimal(idx++);
                    string thisProdNo = prodNo;
                    int minNo = decimal.ToInt32(reader.GetDecimal(idx++));
                    if (minNo == 0)
                        minNo = 1;
                    int fromDate = reader.GetInt32(idx++);
                    int toDate = reader.GetInt32(idx++);

                    if (discount > 1000000.0M)
                        continue;

                    if (Utils.ReadConfigInt32("Vinoble", 0) > 0 || Utils.ReadConfigInt32("Novin", 0) > 0)
                        discount *= 1.25M;
                    // Vinoble B2B - no vat

                    bool found = false;
                    foreach (VismaDiscount vd in discounts)
                    {
                        if (vd.Discount == discount && vd.MinNo == minNo)
                        {
                            found = true;
                            break;
                        }
                    }
                    if (found == false)
                    {
                        discounts.Add(new VismaDiscount()
                        {
                            Discount = discount,
                            MinNo = minNo,
                        });

                    }
                }
            }
            catch (Exception ex)
            {
                errmsg = "GetDiscounts() - " + ex.Message + " - " + sql;

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }

            return true;

        }

        public bool GetLongDescription(string prodNo, ref string descr, int langNo, out string errmsg)
        {
            errmsg = "";
            descr = "";

            string sql = string.Format("SELECT TOP 1 FreeInf1.NoteNm FROM FreeInf1 WITH (NOLOCK) WHERE InfCatNo=54 AND ProdNo = '{0}'", prodNo.Replace("'", "''"));
            if (langNo == 44)
                sql = $"SELECT TOP 1 FreeInf1.NoteNm FROM FreeInf1 WITH (NOLOCK) WHERE InfCatNo=82 AND ProdNo = '{prodNo.Replace("'", "''")}'";
            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                if (reader.Read())
                {
                    string noteNm = reader.GetString(0).Trim();              // FreeInf1.NoteNm
                    descr = Utils.ReadMemoFile(noteNm);
                }
            }
            catch (Exception ex)
            {
                errmsg = "GetLongDescription() - " + ex.Message;

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection != null)
                    connection.Close();
                command.Dispose();
            }

            return true;
        }


        public bool GetAttributeChangeDates(ref AttributeChangeDates attributeChangeDates, out string errmsg)
        {
            errmsg = "";
            string sql = "SELECT  'ProducerDa',MAX(CAST(ChDt as bigint) * 10000 + CAST(Chtm as bigint)) FROM R6 WHERE Nm<>'' AND LTRIM(RTRIM(Inf2))='' " +
                            "UNION " +
                            "SELECT 'ProducerEn',MAX(CAST(ChDt as bigint) * 10000 + CAST(Chtm as bigint))  FROM FreeInf1 WHERE InfCatNo = 81 " +
                            "UNION " +
                            "SELECT  'Grapes',MAX(CAST(ChDt as bigint) * 10000 + CAST(Chtm as bigint))  FROM txt WHERE txttp = 159 " +
                            "UNION " +
                            "SELECT  'Country', MAX(CAST(ChDt as bigint) * 10000 + CAST(Chtm as bigint))  FROM txt WHERE txttp = 38 " +
                            "UNION " +
                            "SELECT  'Region', MAX(CAST(ChDt as bigint) * 10000 + CAST(Chtm as bigint))  FROM txt WHERE txttp = 36 " +
                            "UNION " +
                            "SELECT  'Type', MAX(CAST(ChDt as bigint) * 10000 + CAST(Chtm as bigint))  FROM txt WHERE txttp = 72 " +
                            "UNION " +
                            "SELECT  'Volume', MAX(CAST(ChDt as bigint) * 10000 + CAST(Chtm as bigint))  FROM Prod  ";
            SqlCommand command;
            SqlDataReader reader = null;

            command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                int row = 0;
                while (reader.Read())
                {
                    string s = reader.GetString(0).Trim();
                    long chdttm = reader.GetInt64(1);
                    long chdt = chdttm / 10000;
                    long chtm = chdttm - 10000 * chdt;
                    DateTime tm = Utils.VismaDate2DateTime((int)chdt, (int)chtm);
                    switch (row)                    
                    {
                        case 0:
                            attributeChangeDates.ChDtProducerDa = tm;
                            break;
                        case 1:
                            attributeChangeDates.ChDtProducerEn = tm;
                            break;
                        case 2:
                            attributeChangeDates.ChDtGrapes = tm;
                            break;
                        case 3:
                            attributeChangeDates.ChDtCountry = tm;
                            break;
                        case 4:
                            attributeChangeDates.ChDtRegion = tm;
                            break;
                        case 5:
                            attributeChangeDates.ChDtType = tm;
                            break;
                        case 6:
                            attributeChangeDates.ChDtVolume = tm;
                            attributeChangeDates.ChDtYear = tm;
                            break;
                    }
                    row++;
                    
                }
            }
            catch (Exception ex)
            {
                errmsg = "GetAttributeChangeDates() - " + ex.Message;

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection != null)
                    connection.Close();
                command.Dispose();
            }

            return true;

        }

        public bool GetProductImage(string prodNo, ref string url, out string errmsg)
        {
            url = "";
            errmsg = "";

            string sql = string.Format("SELECT TOP 1 FreeInf1.PictFNm,FreeInf1.PictFNm FROM FreeInf1 WITH (NOLOCK) WHERE FreeInf1.Val1=1 AND FreeInf1.InfcatNo = 52 AND FreeInf1.ProdNo = '{0}'", prodNo.Replace("'", "''"));
            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                if (reader.Read())
                {
                    url = reader.GetString(0);
                    if (url == "")
                        url = reader.GetString(1);
                }
            }
            catch (Exception ex)
            {
                errmsg = "GetProductImage() - " + ex.Message;

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }

            return true;
        }

        public bool GetProductRelations(string prodNo, ref List<string> prodNoRelations, out string errmsg)
        {
            prodNoRelations.Clear();
            errmsg = "";
            string sql = queryGetProductRelations.Replace("#1#", prodNo);

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    prodNoRelations.Add(reader.GetString(0).Trim());
                }
            }
            catch (Exception ex)
            {
                errmsg = "GetLongDescription() - " + ex.Message;

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection != null)
                    connection.Close();
                command.Dispose();
            }

            return true;
        }


        // Find Relations from producer to country and region..
        public bool GetProducerInfo(ref List<VismaProducer> producers, int langNo, out string errmsg)
        {
            producers.Clear();
            errmsg = "";

            string sql = "SELECT R6.Nm,ISNULL(Txt4.Txt, ''),ISNULL(Txt1.Txt, ''),R6.Rno FROM R6 " +
                        $"LEFT OUTER JOIN Txt As Txt4 ON Txt4.Lang = {langNo} AND Txt4.TxtTp = 36 AND TXT4.TxtNo = R6.Gr " +
                        $"LEFT OUTER JOIN Txt As Txt1 ON Txt1.Lang = {langNo} AND Txt1.TxtTp = 38 AND TXT1.TxtNo = r6.Gr3 " +
                        "WHERE R6.Nm <> '' ";


            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    producers.Add(new VismaProducer()
                    {
                        LangNo = langNo,
                        Producer = reader.GetString(0),
                        ProducerRegion = reader.GetString(1),
                        ProducerCountry = reader.GetString(2),
                        VismaID = reader.GetInt32(3)
                    });

                }
            }
            catch (Exception ex)
            {
                errmsg = "GetLongDescription() - " + ex.Message;

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection != null)
                    connection.Close();
                command.Dispose();
            }

            return true;
        }

        public bool GetProductGrapes(string prodNo, ref List<string> grapes, int langNo, out string errmsg)
        {
            grapes.Clear();
            errmsg = "";
            string sql = queryGetProductGrapes.Replace("#1#", prodNo).Replace("#8#", langNo.ToString());

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    //grapes.Add(new VismaAttribute(){ Name= reader.GetString(0).Trim(), Id = reader.GetInt32(1), Type = Constants.AttributeNameGrape );
                    grapes.Add(reader.GetString(0).Trim());
                }
            }
            catch (Exception ex)
            {
                errmsg = "GetLongDescription() - " + ex.Message;

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection != null)
                    connection.Close();
                command.Dispose();
            }

            return true;
        }

        public bool UpdateOrderLineProccessingMethod(int ordNo, int lnNo, bool setBit, out string errmsg)
        {
            errmsg = "";
            string sql;

            if (setBit)
                sql = string.Format("UPDATE OrdLn SET ExcPrint = ExcPrint | 0x00004000 WHERE OrdNo={0} AND LnNo={1}", ordNo, lnNo);
            else
                sql = string.Format("UPDATE OrdLn SET ExcPrint = ExcPrint & 0xFFFFBFFF WHERE OrdNo={0} AND LnNo={1}", ordNo, lnNo);

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            try
            {

                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                command.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                errmsg = "UpdateOrderStatus() -" + ex.Message;

                return false;
            }
            finally
            {
                if (connection.State == ConnectionState.Open)
                    connection.Close();
            }
            return true;
        }

        public int ExistingVismaWebOrder(string webOrderNumber, ref DateTime changeTime, out string errmsg)
        {
            int ordNo = 0;
            errmsg = "";
            changeTime = DateTime.Now;

            string sql = "SELECT TOP 1 OrdNo,ChDt,ChTm FROM Ord WITH (NOLOCK) WHERE CSOrdNo='" + webOrderNumber + "'";

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 500
            };

            SqlDataReader reader = null;

            try
            {

                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                if (reader.Read())
                {
                    ordNo = reader.GetInt32(0);

                    int dt = reader.GetInt32(1);
                    int tm = (int)reader.GetInt16(2);

                    changeTime = Utils.VismaDate2DateTime(dt, tm);
                }
            }
            catch (Exception ex)
            {
                errmsg = "ExistingVismaWebOrder() -" + ex.Message;
                Utils.WriteLog("ERROR ExistingVismaWebOrder() - " + ex.Message);

                return -1;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
            }
            return ordNo;
        }

        public int HasProdID(string prodNo, out string errmsg)
        {
            errmsg = "";
            int hasProduct = 0;

            string sql = "SELECT TOP 1 ProdNo FROM Prod WHERE ProdNo='" + prodNo + "'";
            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                if (reader.Read())
                {
                    hasProduct = 1;
                }
            }
            catch (Exception ex)
            {
                errmsg = "GetProductUnit() - " + ex.Message;
                Utils.WriteLog("HasProdID", ex);
                return -1;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
            }

            return hasProduct;
        }

        public bool GetCustomer(string email, int custNoIn, ref Models.Visma.VismaActor customer, out string errmsg)
        {
            errmsg = "";
            customer.CustomerNo = 0;
            string sql;
            if (custNoIn > 0)
                sql = queryGetCustomerFromCustNo.Replace("#1#", custNoIn.ToString());
            else
                sql = queryGetCustomerFromEmail.Replace("#1#", email);

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                if (reader.Read())
                {
                    int fld = 0;

                    customer.CustomerNo = reader.GetInt32(fld++);
                    customer.EmailAddress = reader.GetString(fld++);
                    customer.Name = reader.GetString(fld++);
                    customer.AddressLine1 = reader.GetString(fld++);
                    customer.AddressLine2 = reader.GetString(fld++);
                    customer.AddressLine3 = reader.GetString(fld++);
                    customer.PostCode = reader.GetString(fld++);
                    customer.PostalArea = reader.GetString(fld++);
                    customer.Phone = reader.GetString(fld++);
                    customer.Mobile = reader.GetString(fld++);

                    customer.DeliveryName = reader.GetString(fld++);
                    customer.DeliveryAddressLine1 = reader.GetString(fld++);
                    customer.DeliveryAddressLine2 = reader.GetString(fld++);
                    customer.DeliveryAddressLine3 = reader.GetString(fld++);
                    customer.DeliveryPostCode = reader.GetString(fld++);
                    customer.DeliveryPostalArea = reader.GetString(fld++);
                }
            }
            catch (Exception ex)
            {
                errmsg = ex.Message;
                Utils.WriteLog("GetCustomer", ex);
                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
                command.Dispose();
            }

            return true;

        }

        public bool GetNextFreeCustomerNo(ref int custNo, out string errmsg)
        {
            custNo = -1;
            errmsg = "";
            string sql = string.Format("SELECT ISNULL(MAX(CustNo)+1,1) FROM Actor WITH (NOLOCK)");

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            SqlDataReader reader = null;

            try
            {

                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                if (reader.Read())
                {
                    custNo = reader.GetInt32(0);

                }
            }
            catch (Exception ex)
            {
                errmsg = "GetNextFreeCustomerNo() -" + ex.Message;

                return false;
            }
            finally
            {
                if (connection.State == ConnectionState.Open)
                    connection.Close();
            }
            return true;
        }

        public int CheckCustomer(string email, out string errmsg)
        {
            errmsg = "";
            int custNo = 0;
            string sql = "SELECT CustNo FROM Actor WHERE MailAd='" + email + "'";
            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                if (reader.Read())
                {
                    custNo = reader.GetInt32(0);
                }
            }
            catch (Exception ex)
            {
                errmsg = "GetCustomer() - " + ex.Message;

                return -1;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
            }

            return custNo;
        }

        public bool UpdateCustomer(int CustNo, string phone, string mobile, out string errmsg)
        {
            errmsg = "";
            string sql = string.Format("UPDATE Actor SET Phone='{0}', MobPh='{1}' WHERE CustNo={2}", phone, mobile, CustNo);

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            try
            {

                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                command.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                errmsg = "UpdateCustomer() -" + ex.Message;

                return false;
            }
            finally
            {
                if (connection.State == ConnectionState.Open)
                    connection.Close();
            }
            return true;

        }

        public bool UpdateCustomerCity(int custNo, string city, int ctry, out string errmsg)
        {
            errmsg = "";
            string sql = $"UPDATE Actor SET PArea='{city}', Ctry={ctry} WHERE CustNo={custNo}";

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            try
            {

                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                command.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                errmsg = "UpdateCustomer() -" + ex.Message;

                return false;
            }
            finally
            {
                if (connection.State == ConnectionState.Open)
                    connection.Close();
            }
            return true;

        }


        public bool UpdateCustNo(int ActNo, int NewCustNo, int DelToAct, out string errmsg)
        {
            errmsg = "";
            string sql;

            if (DelToAct > 0)
                sql = string.Format("UPDATE Actor SET CustNo=0, DelToAct={0} WHERE ActNo={1}", DelToAct, ActNo);
            else
                sql = string.Format("UPDATE Actor SET CustNo={0}, DelToAct=0 WHERE ActNo={1}", NewCustNo, ActNo);



            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            try
            {

                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                command.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                errmsg = "UpdateCustomer() -" + ex.Message;

                return false;
            }
            finally
            {
                if (connection.State == ConnectionState.Open)
                    connection.Close();
            }
            return true;

        }


        public bool UpdateOrderCardNm(int ordNo, string cardNm, string cardAc, out string errmsg)
        {
            errmsg = "";
            DateTime now = DateTime.Now;
            int today = now.Year * 10000 + now.Month * 100 + now.Day;
            int time = now.Hour * 100 + now.Minute;
            string sql = string.Format("UPDATE Ord SET CardNm='{0}', CardAc='{1}' WHERE OrdNo={2}", cardNm, cardAc, ordNo);


            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            try
            {

                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                command.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                errmsg = "UpdateOrderStatus() -" + ex.Message;

                return false;
            }
            finally
            {
                if (connection.State == ConnectionState.Open)
                    connection.Close();
            }
            return true;
        }

        public bool UpdateOrderStatus(int ordNo, int status, out string errmsg)
        {
            errmsg = "";
            DateTime now = DateTime.Now;
            int today = now.Year * 10000 + now.Month * 100 + now.Day;
            int time = now.Hour * 100 + now.Minute;
            string sql = $"UPDATE Ord SET Gr={status},Chdt={today},ChTm={time}  WHERE OrdNo={ordNo}";

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            try
            {

                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                command.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                errmsg = "UpdateOrderStatus() -" + ex.Message;

                return false;
            }
            finally
            {
                if (connection.State == ConnectionState.Open)
                    connection.Close();
            }
            return true;
        }

        public int GetCurrentVismaDate(int offset, out string errmsg)
        {
            int dt = 0;
            errmsg = "";

            string sql = "SELECT GETDATE()";
            //  if (offset > 0)
            //      sql = string.Format("SELECT DATEADD(day,{0}, GETDATE())", offset);
            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                if (reader.Read())
                {
                    DateTime t = reader.GetDateTime(0);
                    dt = t.Year * 10000 + t.Month * 100 + t.Day;

                    /* if (offset > 0)
                     {
                         DateTime t1 = t.AddDays(1 * offset);

                         if (t1.DayOfWeek == DayOfWeek.Saturday)
                             t1.AddDays(offset + 2);
                         if (t1.DayOfWeek == DayOfWeek.Sunday)
                             t1.AddDays(offset + 1);
                         dt = t1.Year * 10000 + t1.Month * 100 + t1.Day;
                     } */
                    if (offset > 0)
                    {
                        DateTime t1 = t.AddDays(-1 * offset);
                        if (t.DayOfWeek == DayOfWeek.Saturday)
                            t1 = t.AddDays(-1 * offset - 1);
                        if (t.DayOfWeek == DayOfWeek.Sunday)
                            t1 = t.AddDays(-1 * offset - 2);

                        if (offset == 1)
                        {
                            if (t1.DayOfWeek == DayOfWeek.Sunday)
                                t1 = t1.AddDays(-2); // prev friday
                            if (t1.DayOfWeek == DayOfWeek.Saturday) // coming from sunday!!
                                t1 = t1.AddDays(-2);// prev thursday
                        }

                        if (offset == 2)
                        {
                            if (t1.DayOfWeek == DayOfWeek.Sunday)
                                t1 = t1.AddDays(-2); // prev friday
                            if (t1.DayOfWeek == DayOfWeek.Saturday)
                                t1 = t1.AddDays(-2);// prev thursday
                        }

                        if (t1.DayOfWeek == DayOfWeek.Saturday)
                            t1.AddDays(-1 - offset);
                        if (t1.DayOfWeek == DayOfWeek.Sunday)
                            t1.AddDays(-2);
                        dt = t1.Year * 10000 + t1.Month * 100 + t1.Day;
                    }
                }
            }
            catch (Exception ex)
            {
                errmsg = ex.Message;

                return -1;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
            }

            return dt;
        }

        public bool GetAllWebOrdersToClose(ref List<VismaOrderStatus> vismaOrderStatuses, out string errmsg)
        {
            errmsg = "";
            vismaOrderStatuses.Clear();



            string sql = "SELECT Ord.OrdNo,OrdDoc.CsOrdNo,OrdDoc.OrdDocNo FROM OrdDoc " +
                         "INNER JOIN Ord ON Ord.OrdNo = OrdDoc.OrdNo " +
                         "WHERE OrdDoc.CsOrdNo = Ord.CsOrdNo AND LTRIM(RTRIM(OrdDoc.CsOrdNo)) <> '' AND " +
                         "OrdDoc.Gr10 = 0 AND OrdDoc.Doctp = 1";

            //Utils.WriteLog("Main query: " + sql);

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                while (reader.Read())
                {
                    vismaOrderStatuses.Add(new VismaOrderStatus()
                    {
                        OrderNumber = reader.GetInt32(0),
                        CustomerOrSupplierOrderNo = reader.GetString(1),
                        OrderDocNumber = reader.GetInt32(2)
                    });
                }
            }
            catch (Exception ex)
            {
                errmsg = ex.Message;

                return false;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
            }

            return true;
        }

        public bool UpdateOrderStatusFlag(int ordNo, int ordDocNo, out string errmsg)
        {
            errmsg = "";

            string sql = $"UPDATE OrdDoc SET Gr10=1 WHERE OrdNo={ordNo} AND OrdDocNo={ordDocNo}";

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text,
                CommandTimeout = 600
            };

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                errmsg = ex.Message;
                return false;
            }
            finally
            {
                if (connection.State == ConnectionState.Open)
                    connection.Close();
            }
            return true;
        }

        public int LookupCountryCode(string country, out string errmsg)
        {
            errmsg = "";

            if (country.Trim() == "")
                return Utils.ReadConfigInt32("DefaultCountryCode", 45);


            int ctryNo = 0;

            string sql = $"SELECT TOP 1 CtryNo FROM Ctry WHERE ISO='{country}' OR Nm='{country}'";

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                if (reader.Read())
                {
                    ctryNo = reader.GetInt32(0);
                }
            }
            catch (Exception ex)
            {
                errmsg = "LookupCountryCode() -" + ex.Message;

                return -1;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
            }
            return ctryNo;
        }


        public int LookupCurrencyCode(string currency, out string errmsg)
        {
            errmsg = "";



            int curNo = 0;

            string sql = $"SELECT TOP 1 CurNo FROM Cur WHERE ISO='{currency}' OR Nm='{currency}'";

            SqlCommand command = new SqlCommand(sql, connection)
            {
                CommandType = CommandType.Text
            };

            SqlDataReader reader = null;

            try
            {
                if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                    connection.Open();

                reader = command.ExecuteReader();

                if (reader.Read())
                {
                    curNo = reader.GetInt32(0);
                }
            }
            catch (Exception ex)
            {
                errmsg = "LookupCurrencyCode() -" + ex.Message;

                return -1;
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if (connection.State == ConnectionState.Open)
                    connection.Close();
            }
            return curNo;
        }


        /*   public bool GetAttributeTermsForCounty(ref List<ProductAttributeTerm> attributeTerms, out string errmsg)
           {
               attributeTerms.Clear();
               errmsg = "";

               string sql = "SELECT DISTINCT txt,CAST(txtno as int) from txt where TxtTp=37 and Lang = 45 ";

               SqlCommand command = new SqlCommand(sql, connection)
               {
                   CommandType = CommandType.Text,
                   CommandTimeout = 600
               };

               SqlDataReader reader = null;

               try
               {
                   if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                       connection.Open();

                   reader = command.ExecuteReader();

                   while (reader.Read())
                   {
                       string name = reader.GetString(0);
                       int vid = reader.GetInt32(1);
                       attributeTerms.Add(new ProductAttributeTerm() { name = name, slug = Utils.SanitizeSlugName(name), visma_id = vid });
                   }

               }
               catch (Exception ex)
               {
                   errmsg = ex.Message;

                   return false;
               }
               finally
               {
                   if (reader != null)
                       reader.Close();
                   if (connection.State == ConnectionState.Open)
                       connection.Close();
                   command.Dispose();
               }

               return true;
           }*/

        /*      public bool GetSecondaryCategories(string prodNo, ref List<string> descr, out string errmsg)
              {
                  errmsg = "";
                  descr.Clear();

                  string sql =
                          "SELECT DISTINCT ProdCat.Descr " +
                             "FROM FreeInf1 WITH (NOLOCK) " +
                             "INNER JOIN ProdCat WITH (NOLOCK) ON ProdCat.PrCatNo = FreeInf1.PrCatNo " +
                             $"WHERE FreeInf1.InfCatNo = 90 AND FreeInf1.PrCatNo > 0 AND FreeInf1.ProdNo={prodNo} AND ProdCat.Descr<>'' ";

                  SqlCommand command;
                  SqlDataReader reader = null;

                  command = new SqlCommand(sql, connection)
                  {
                      CommandType = CommandType.Text,
                      CommandTimeout = 600
                  };

                  try
                  {
                      if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                          connection.Open();

                      reader = command.ExecuteReader();

                      while (reader.Read())
                      {
                          descr.Add(reader.GetString(0).Trim());
                      }
                  }
                  catch (Exception ex)
                  {
                      errmsg = "GetLongDescription() - " + ex.Message;

                      return false;
                  }
                  finally
                  {
                      if (reader != null)
                          reader.Close();
                      if (connection != null)
                          connection.Close();
                      command.Dispose();
                  }

                  return true;
              }
        */

        //public bool GetCategoryTag(string prodNo, ref List<string> categoryTags, out string errmsg)
        //{
        //    errmsg = "";
        //    categoryTags.Clear();
        //    string p = prodNo.Replace("'", "''");

        //    string sql = $"SELECT DISTINCT ProdCat.Descr FROM ProdCat INNER JOIN FreeInf1 ON FreeInf1.PrCatNo=ProdCat.PrCatNo WHERE FreeInf1.InfCatNo=74 AND FreeInf1.ProdNo='{p}'";
        //    SqlCommand command = new SqlCommand(sql, connection)
        //    {
        //        CommandType = CommandType.Text
        //    };

        //    SqlDataReader reader = null;

        //    try
        //    {
        //        if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
        //            connection.Open();

        //        reader = command.ExecuteReader();

        //        while (reader.Read())
        //        {
        //            string s = reader.GetString(0);
        //            if (categoryTags.Contains(s) == false)
        //            {
        //                categoryTags.Add(s);
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        errmsg = "GetPrice() - " + ex.Message + " " + sql;

        //        return false;
        //    }
        //    finally
        //    {
        //        if (reader != null)
        //            reader.Close();
        //        if (connection.State == ConnectionState.Open)
        //            connection.Close();
        //        command.Dispose();
        //    }

        //    if (categoryTags.Count == 0)
        //        categoryTags.Add("Øvrige");

        //    return true;
        //}


        /*       public bool GetProductDetails(string prodNo, ref VismaProductDetails item, out string errmsg)
               {
                   errmsg = "";

                   string sql = queryGetProductDetails.Replace("#1#", prodNo.Replace("'", "''"));

                   SqlCommand command = new SqlCommand(sql, connection)
                   {
                       CommandType = CommandType.Text
                   };

                   SqlDataReader reader = null;

                   try
                   {
                       if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                           connection.Open();

                       reader = command.ExecuteReader();

                       if (reader.Read())
                       {
                           int idx = 0;
                           item.Type = reader.GetString(idx++).Trim();
                           item.Year = reader.GetInt32(idx++).ToString();// ProdTp4
                           if (item.Year == "0")
                               item.Year = "";
                           if (item.Year == "1")
                               item.Year = "NV";
                           int organic = reader.GetInt32(idx++);           // Gr3
                           int bio = reader.GetInt32(idx++);               // Gr5
                           int nature = reader.GetInt32(idx++);            // Gr11
                           item.Organic = organic == 1 ? "Ja" : "Nej";
                           item.Biodynamic = bio == 1 ? "Ja" : "Nej";
                           item.Naturewine = nature == 1 ? "Ja" : "Nej";

                           item.Weight = reader.GetDecimal(idx++);                     // TareU
                           item.Content = reader.GetDecimal(idx++) * 100.0M;                        // WdtU
                           item.Alcohol = reader.GetDecimal(idx++);  //Prod.Free1,
                           if (item.Alcohol > 990.0M)
                               item.Alcohol = 0.0M;
                           item.UnitsPerPackage = reader.GetInt32(idx++);              // ProdTp2
                           item.Country = reader.GetString(idx++).Trim();  // Cat1 tekst = reader.GetString(idx++).Trim();  // Cat2 tekst

                           item.District = reader.GetString(idx++).Trim();  // Cat2 tekst = reader.GetString(idx++).Trim();  // Cat2 tekst
                           item.County = reader.GetString(idx++).Trim();  // Cat3 tekst
                           item.Producer = reader.GetString(idx++).Trim();  // Cat4 tekst
                           item.Classification = reader.GetString(idx++).Trim();  // Cat5 tekst
                           item.Mark = reader.GetString(idx++).Trim();  // Cat6 tekst

                           item.SubType = reader.GetString(idx++).Trim(); // SubType
                       }
                   }
                   catch (Exception ex)
                   {
                       errmsg = "GetProductDetails() - " + ex.Message + " " + sql;

                       return false;
                   }
                   finally
                   {
                       if (reader != null)
                           reader.Close();
                       if (connection.State == ConnectionState.Open)
                           connection.Close();
                       command.Dispose();
                   }
                   string scores = "";
                   if (GetProductScores(prodNo, ref scores, out errmsg))
                       item.Scores = scores;
                   return true;
               }
               */

        /*   public bool GetCategories(ref List<ProductCategory> categories, out string errmsg)
           {
               categories.Clear();
               errmsg = "";

               string sql = queryGetCategories;

               SqlCommand command = new SqlCommand(sql, connection)
               {
                   CommandType = CommandType.Text,
                   CommandTimeout = 600
               };

               SqlDataReader reader = null;

               try
               {
                   if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                       connection.Open();

                   reader = command.ExecuteReader();

                   while (reader.Read())
                   {
                       int catNo = reader.GetInt32(0);
                       string name = reader.GetString(1).Trim();
                       int parent_catNo = reader.GetInt32(2);



                       int level = reader.GetInt32(3);
                       ProductCategory c = categories.FirstOrDefault(p => p.vismaCatNo == catNo);
                       if (c == null)
                       {

                           if (categories.Exists(p => p.name == name && p.vismaLevel != parent_catNo))
                               name += ".";

                           categories.Add(new ProductCategory()
                           {
                               name = name,
                               slug = Utils.SanitizeSlugName(name),
                               description = name,

                               vismaCatNo = catNo,
                               vismaParentCatNo = parent_catNo,
                               vismaLevel = level
                           });
                       }
                   }

               }
               catch (Exception ex)
               {
                   errmsg = ex.Message;

                   return false;
               }
               finally
               {
                   if (reader != null)
                       reader.Close();
                   if (connection.State == ConnectionState.Open)
                       connection.Close();
                   command.Dispose();
               }


               foreach(ProductCategory c in categories)
               {
                   if (c.vismaParentCatNo > 0)
                   {
                       // Find parent
                       ProductCategory c_parent = categories.FirstOrDefault(p => p.vismaCatNo == c.vismaParentCatNo);
                       if (c_parent != null)
                           c.parent = c_parent.id;
                   }
               }
               return true;
           }
        */
        /*     public bool GetTags(ref List<ProductTag> tags, out string errmsg)
             {
                 tags.Clear();
                 errmsg = "";

                 string sql = queryGetTags;

                 SqlCommand command = new SqlCommand(sql, connection)
                 {
                     CommandType = CommandType.Text,
                     CommandTimeout = 600
                 };

                 SqlDataReader reader = null;

                 try
                 {
                     if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                         connection.Open();

                     reader = command.ExecuteReader();

                     while (reader.Read())
                     {
                         string name = reader.GetString(0);
                         tags.Add(new ProductTag() { name = name, slug = Utils.SanitizeSlugName(name) });
                     }

                 }
                 catch (Exception ex)
                 {
                     errmsg = ex.Message;

                     return false;
                 }
                 finally
                 {
                     if (reader != null)
                         reader.Close();
                     if (connection.State == ConnectionState.Open)
                         connection.Close();
                     command.Dispose();
                 }

                 tags.Add(new ProductTag() { name = "Økologisk", slug = Utils.SanitizeSlugName("Økologisk") });

                 return true;
             }
        */
        /*      public bool GetTagsForYear(ref List<ProductTag> tags, int langNo, out string errmsg)
              {
                  errmsg = "";

                  string sql = "SELECT DISTINCT ProdTp4 FROM Prod WHERE ProdTp4>0";



                  SqlCommand command = new SqlCommand(sql, connection)
                  {
                      CommandType = CommandType.Text,
                      CommandTimeout = 600
                  };

                  SqlDataReader reader = null;

                  try
                  {
                      if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                          connection.Open();

                      reader = command.ExecuteReader();

                      while (reader.Read())
                      {
                          int y = reader.GetInt32(0);
                          string name = Constants.YearPrefix + ( y == 1 ? "N.V." : y.ToString());
                          tags.Add(new ProductTag() { name = name, slug = Utils.SanitizeSlugNameNew(name), lang = Utils.LangNoToString(langNo), translations = new Translations() });
                      }

                  }
                  catch (Exception ex)
                  {
                      errmsg = ex.Message;

                      return false;
                  }
                  finally
                  {
                      if (reader != null)
                          reader.Close();
                      if (connection.State == ConnectionState.Open)
                          connection.Close();
                      command.Dispose();
                  }

                  return true;
              }
        */
        /*      public bool GetTagsForVolume(ref List<ProductTag> tags, int langNo, out string errmsg)
              {
                  errmsg = "";

                  string sql = "SELECT DISTINCT WdtU FROM Prod WHERE  WdtU>0 ORDER BY WdtU";

                  SqlCommand command = new SqlCommand(sql, connection)
                  {
                      CommandType = CommandType.Text,
                      CommandTimeout = 600
                  };

                  SqlDataReader reader = null;

                  try
                  {
                      if (connection.State == ConnectionState.Closed || connection.State == ConnectionState.Broken)
                          connection.Open();

                      reader = command.ExecuteReader();

                      while (reader.Read())
                      {
                          string name = Constants.VolumePrefix + Utils.DecimalToStringFloating(reader.GetDecimal(0)) + " l";
                          tags.Add(new ProductTag() { name = name, slug = Utils.SanitizeSlugName(name), lang = Utils.LangNoToString(langNo), translations = new Translations() });
                      }

                  }
                  catch (Exception ex)
                  {
                      errmsg = ex.Message;

                      return false;
                  }
                  finally
                  {
                      if (reader != null)
                          reader.Close();
                      if (connection.State == ConnectionState.Open)
                          connection.Close();
                      command.Dispose();
                  }

                  return true;
              }
        */

    }
}
