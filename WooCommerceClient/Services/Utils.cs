using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Drawing;
using System.Xml.Serialization;
using System.Text;
using Newtonsoft.Json;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using WooCommerceClient.Models.Visma;

namespace WooCommerceClient
{
    public static class Utils
    {
        public static int ReadConfigInt32(string setting, int defaultValue)
        {
            int ret = defaultValue;
            if (ConfigurationManager.AppSettings[setting] != null)
            {
                try
                {
                    return Convert.ToInt32((string)ConfigurationManager.AppSettings[setting]);
                }
                catch { }
            }

            return ret;
        }

        public static string ReadConfigString(string setting, string defaultValue)
        {
            string ret = defaultValue;
            if (ConfigurationManager.AppSettings[setting] != null)
            {
                try
                {
                    return (string)ConfigurationManager.AppSettings[setting];
                }
                catch { }
            }

            return ret;
        }

        public static string SanitizeSlugNameNew(string s)
        {
            return SanitizeSlugName(s).Replace("_", "-");
        }

        private static string SanitizeSlugName(string s)
        {
            if (s == "")
                return s;

            s = s.ToLower().Replace("-", "_");
            s = s.Replace(" ", "_");
            s = s.Replace(".", "_");
            s = s.Replace("æ", "ae");
            s = s.Replace("ø", "oe");
            s = s.Replace("/", "_");
            s = s.Replace("ñ", "n");

            s = s.Replace("õ", "o");
            s = s.Replace("ô", "o");
            s = s.Replace("ó", "o");
            s = s.Replace("ò", "o");
            s = s.Replace("ö", "o");

            s = s.Replace("î", "i");
            s = s.Replace("í", "i");
            s = s.Replace("ì", "i");
            s = s.Replace("ï", "i");


            s = s.Replace("ã", "a");
            s = s.Replace("á", "a");
            s = s.Replace("à", "a");
            s = s.Replace("ä", "a");

            s = s.Replace("ý", "y");
            s = s.Replace("ÿ", "y");

            s = s.Replace("é", "e");
            s = s.Replace("è", "e");
            s = s.Replace("ë", "e");
            s = s.Replace("ê", "e");

            s = s.Replace("ü", "u");
            s = s.Replace("ú", "u");
            s = s.Replace("ù", "u");
            s = s.Replace("û", "u");

            s = s.Replace("&", "_");

            s = s.Replace("%", "");
            s = s.Replace(":", "");

            // if (s.Length > 28)
            //   s = s.Substring(0, 28);
            return s;
        }

        public static void WriteLog(string prefix, Exception ex)
        {
            var e = ex.GetBaseException();
            WriteLog($"{prefix} got {e.GetType().Name}:\n{e.Message}\n{e.StackTrace}");
        }

        public static void WriteLog(string logoutput)
        {
            Console.WriteLine(logoutput);
            string LogPath = AppDomain.CurrentDomain.BaseDirectory + @"\WooCommerceClient.log";

           
            if (LogPath.ToLower() != "" && LogPath.ToLower() != "stdout")
            {
                try
                {
                    StreamWriter w = File.AppendText(LogPath);
                    w.WriteLine(System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + ": " + logoutput);
                    w.Flush();
                    w.Close();

                    FileInfo fi = new FileInfo(LogPath);
                    if (fi.Length > 20 * 1024 * 1024)
                    {
                        File.Move(LogPath, AppDomain.CurrentDomain.BaseDirectory + @"\WooCommerceClient2.log");
                    }
                }
                catch (Exception)
                {
                }
            }
        }

        public static int UnixTimeStampNow()
        {
            var timeSpan = (DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0));
            return (int)timeSpan.TotalSeconds;
        }

        public static int UnixTimeStamp(DateTime dt)
        {
            //return (int)Math.Truncate((dt.ToUniversalTime().Subtract(new DateTime(1970, 1, 1))).TotalSeconds);

            return (int)((TimeZoneInfo.ConvertTimeToUtc(dt) -
                   new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds);
        }

        public static DateTime UnixTimeStampToDateTime(int unixDateTime)
        {
            var epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            var timeSpan = TimeSpan.FromSeconds(unixDateTime);
            return epoch.Add(timeSpan).ToLocalTime();
        }

        public static string GenerateTimeStamp()
        {
            DateTime t = DateTime.Now;
            return string.Format("{0:00}{1:00}{2:00}{3:00}{4:00}{5:00}", t.Year - 2000, t.Month, t.Day, t.Hour, t.Minute, t.Second);
        }
        public static string GenerateTimeStamp(DateTime t)
        {
            return string.Format("{0:0000}-{1:00}-{2:00} {3:00}:{4:00}:{5:00}", t.Year, t.Month, t.Day, t.Hour, t.Minute, t.Second);
        }

        public static string GenerateTimeStampT(DateTime t)
        {
            return string.Format("{0:0000}-{1:00}-{2:00}T{3:00}:{4:00}:{5:00}", t.Year, t.Month, t.Day, t.Hour, t.Minute, t.Second);
        }

        public static string GenerateDateStamp(DateTime t)
        {
            return string.Format("{0:0000}-{1:00}-{2:00}", t.Year, t.Month, t.Day);
        }

        public static string DateTimeToVismaDateString(DateTime t)
        {
            return string.Format("{0:0000}{1:00}{2:00}", t.Year, t.Month, t.Day);
        }

        public static string DateTimeToISO8601(DateTime t)
        {
            return t.ToString("yyyy-MM-ddTHH:mm:sszzz");    // local
        }

        public static int StringToInt(string s)
        {
            Int32.TryParse(s, out int n);

            return n;
        }



        public static DateTime VismaDate2DateTime(int dateint)
        {
            ///           20160101
            if (dateint < 10000000)
                return DateTime.MinValue;
            int year = dateint / 10000;
            int month = (dateint - year * 10000) / 100;
            int day = dateint - year * 10000 - month * 100;
            if (year <= 1900 || month < 1 || month > 12 || day < 1 || day > 31)
                return DateTime.MinValue;
            try
            {
                return new DateTime(year, month, day);
            }
            catch
            {
                return DateTime.MinValue;
            }
        }

        public static DateTime VismaDate2DateTime(int dateint, int timeint)
        {
            ///           20160101
            if (dateint < 10000000)
                return DateTime.MinValue;
            int year = dateint / 10000;
            int month = (dateint - year * 10000) / 100;
            int day = dateint - year * 10000 - month * 100;
            int hour = timeint / 100;
            int min = timeint - (hour * 100);
            int sec = 0;
            if (year <= 1900 || month < 1 || month > 12 || day < 1 || day > 31)
                return DateTime.MinValue;
            try
            {
                return new DateTime(year, month, day, hour, min, sec);
            }
            catch
            {
                return DateTime.MinValue;
            }
        }

        public static DateTime VismaDate2DateTimeFromString(string datestring)
        {
            Int32.TryParse(datestring, out int dateint);
            return VismaDate2DateTime(dateint);
        }
        

        public static string IntList2String(List<int> arr)
        {
            string s = "";

            foreach (int n in arr)
            {
                if (s != "")
                    s += ",";
                s += n.ToString();
            }
            return s;
        }

        public static int DateTimeToVismaDate(DateTime dt)
        {
            return dt.Year * 10000 + dt.Month * 100 + dt.Day;
        }

        public static string QuoteIt(string s)
        {
            return "'" + s + "'";
        }

        public static string DecimalToString(Decimal dec)
        {
            string s = string.Format("{0:0.00}", dec);
            return s.Replace(',', '.');
        }

        public static int DecimalToWallMobInt(Decimal dec)
        {
            Decimal d10 = dec * 100.0m;
            int i;
            try
            {
                i = Decimal.ToInt32(d10);
            }
            catch
            {
                return 0;
            }
            return i;
        }

        public static decimal WallMobIntToDecimal(int n)
        {
            decimal dec = (decimal)n;

            return dec / 100.0m;
        }
        public static double StringToDouble(string s)
        {

            Double.TryParse(s.Replace(".", ","), out double f);

            return f;
        }

        public static string DecimalToStringFloating(Decimal dec)
        {
            if (decimal.Truncate(dec) == dec)
                return Decimal.ToInt32(dec).ToString();
            Decimal d1 = dec * 10.0M;
            if (decimal.Truncate(d1) == d1)
                return string.Format("{0:0.0}", dec).Replace(',', '.');
            d1 = dec * 100.0M;
            if (decimal.Truncate(d1) == d1)
                return string.Format("{0:0.00}", dec).Replace(',', '.');

            return string.Format("{0:0.000}", dec).Replace(',', '.');

        }

        public static decimal StringToDecimal(string s)
        {
            /*

            s = s.Replace(",", ".");
            double f = 0;
            Double.TryParse(s, out f);
            try
            {
                return Convert.ToDecimal(f);
            }
            catch
            {
                return 0;
            }

           */

            CultureInfo culture = CultureInfo.GetCultureInfo("da-DK");

            NumberFormatInfo nfi = culture.NumberFormat;

            NumberStyles style;
            style = NumberStyles.AllowDecimalPoint;
            if (Decimal.TryParse(s, style, culture, out decimal number))
                return number;
            else

                return 0;

        }

        public static string LangNoToString(int langNo)
        {
           switch (langNo)
            {
                case 45: return "da";
                case 44: return "en";

                default:
                    return "da";
            }
        }

        public static int DecimalToIntSafe(decimal d)
        {
            try
            {
                return Decimal.ToInt32(d);
            }
            catch
            {
                return 0;
            }
        }

        public static DateTime StringToDateTime(string s)
        {
            //2016.10.17 13:42:51
            //0123456789012345678
            if (s.Length < 19)
                return DateTime.MinValue;
            try
            {
                return new DateTime(Int32.Parse(s.Substring(0, 4)), Int32.Parse(s.Substring(5, 2)), Int32.Parse(s.Substring(8, 2)), Int32.Parse(s.Substring(5, 2)), Int32.Parse(s.Substring(11, 2)), Int32.Parse(s.Substring(14, 2)), Int32.Parse(s.Substring(17, 2)));
            }
            catch
            {

            }

            return DateTime.MinValue;
        }

        public static DateTime StringToDate(string s)
        {
            //2016.10.17
            if (s.Length < 10)
                return DateTime.MinValue;
            try
            {
                return new DateTime(Int32.Parse(s.Substring(0, 4)), Int32.Parse(s.Substring(5, 2)), Int32.Parse(s.Substring(8, 2)));
            }
            catch
            {

            }

            return DateTime.MinValue;
        }

        public static string Base64Encode(string plainText)
        {
            if (plainText == "")
                return "";
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
            return System.Convert.ToBase64String(plainTextBytes);
        }

        public static byte[] Base64EncodeToByteArray(string plainText)
        {
            if (plainText == "")
                return null;
            return System.Text.Encoding.UTF8.GetBytes(plainText);
        }

        public static string Base64Decode(string base64EncodedData)
        {
            if (base64EncodedData == "")
                return "";
            var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }


        public static void CleanFolder(string folder)
        {
            try
            {

                DirectoryInfo dirInfo = new DirectoryInfo(folder);
                if (dirInfo.Exists == false)
                    return;

                foreach (FileInfo fileinfo in dirInfo.GetFiles())
                {
                    FileAttributes attributes = File.GetAttributes(fileinfo.FullName);
                    if ((attributes & FileAttributes.Directory) == FileAttributes.Directory)
                        continue;
                    if ((attributes & FileAttributes.Hidden) == FileAttributes.Hidden)
                        continue;
                    if ((attributes & FileAttributes.System) == FileAttributes.System)
                        continue;
                    if (fileinfo.Name[0] == '.')
                        continue;
                    if (fileinfo.Length == 0)
                        continue;

                    File.Delete(fileinfo.FullName);

                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"The process failed: {e.Message}");
            }
            finally { }

        }

        // server=.;uid=sa;pwd=infra;database=F9999
        //public static string FormConnectionString()
        // {
        // return string.Format("server={0};uid={1};pwd={2};database={3}", Properties.Settings.Default.DBserver, Properties.Settings.Default.DBUser, Encryptor.Decrypt(Properties.Settings.Default.DBpassword), Properties.Settings.Default.Database);
        // }

        public static DateTime FirstSecondOfDay(DateTime value)
        {
            return new DateTime(value.Year, value.Month, value.Day, 0, 0, 0);
        }

        public static DateTime LastSecondOfDay(DateTime value)
        {
            return new DateTime(value.Year, value.Month, value.Day, 23, 59, 59);
        }

        public static string StringListToString(List<string> list)
        {
            string s = "";
            foreach (string sl in list)
            {
                if (s != string.Empty)
                    s += ",";
                s += sl;
            }

            return s;
        }

        public static string SanitizeName(string name)
        {
            // to bo defined...
            return name.Replace(@"\", "");
        }

        public static string SanitizeUrl(string name)
        {
            // to bo defined...
            name = name.Replace("\x60", "");
            name = name.Replace("\x2c", "");
            name = name.Replace("\x27", "");
            name = name.Replace(@"'", "");
            name = name.Replace("\"", "");
            name = name.Replace(@"´", "");
            name = name.Replace(@"`", "");

            name = name.Replace(@"+", "-");
            name = name.Replace(@" ", "-");
            name = name.Replace(@"_", "-");
            name = name.Replace(@",", "-");
            name = name.Replace(@"%", "pct");
            name = name.Replace(@"æ", "ae");
            name = name.Replace(@"ø", "oe");
            name = name.Replace(@"å", "aa");
            name = name.Replace(@"Æ", "AE");
            name = name.Replace(@"Ø", "OE");
            name = name.Replace(@"Å", "AA");

            name = name.Replace(@"ü", "uu");
            name = name.Replace(@"ú", "uu");
            name = name.Replace(@"û", "uu");
            name = name.Replace(@"ä", "aa");
            name = name.Replace(@"á", "aa");
            name = name.Replace(@"â", "aa");
            name = name.Replace(@"ö", "oo");
            name = name.Replace(@"ó", "oo");
            name = name.Replace(@"ô", "oo");
            name = name.Replace(@"ÿ", "yy");

            name = name.Replace(@"ï", "ii");
            name = name.Replace(@"í", "ii");
            name = name.Replace(@"î", "ii");
            name = name.Replace(@"ë", "ee");
            name = name.Replace(@"ê", "ee");
            name = name.Replace(@"é", "ee");
            name = name.Replace(@"Ü", "UU");
            name = name.Replace(@"Ú", "UU");
            name = name.Replace(@"Û", "UU");
            name = name.Replace(@"Ä", "AA");
            name = name.Replace(@"Á", "AA");
            name = name.Replace(@"Â", "AA");
            name = name.Replace(@"Ö", "OO");
            name = name.Replace(@"Ó", "OO");
            name = name.Replace(@"Ô", "OO");
            name = name.Replace(@"Ï", "II");
            name = name.Replace(@"Í", "II");
            name = name.Replace(@"Î", "II");
            name = name.Replace(@"Ë", "EE");
            name = name.Replace(@"Ê", "EE");
            name = name.Replace(@"É", "EE");
            return name;
        }

        public static string CleanForJSON(string s)
        {
            if (s == null || s.Length == 0)
            {
                return "";
            }

            char c = '\0';

            int len = s.Length;
            StringBuilder sb = new StringBuilder(len + 4);
            String t;

            for (int i = 0; i < len; i += 1)
            {
                c = s[i];
                switch (c)
                {

                    case '\\':
                        //sb.Append("\\");
                        //.Append(c);
                        break;
                    case '"':
                        //  case '/':
                        sb.Append("\\");
                        sb.Append(c);
                        break;
                    case '\b':
                        sb.Append("\\b");
                        break;
                    case '\t':
                        sb.Append("\\t");
                        break;
                    case '\n':
                        sb.Append("\\n");
                        break;
                    case '\f':
                        sb.Append("\\f");
                        break;
                    case '\r':
                        sb.Append("\\r");
                        break;
                    default:
                        if (c < ' ')
                        {
                            t = "000" + String.Format("X", c);
                            sb.Append("\\u" + t.Substring(t.Length - 4));
                        }
                        else
                        {
                            sb.Append(c);
                        }
                        break;
                }
            }
            return sb.ToString();
        }

        public static string SanitizeDescription(string descr, bool killHtmlTags)
        {
            //WriteLog("Description:" + descr);
            descr = CleanForJSON(descr.Replace("\r", "").Replace("\\'", "'").Replace("\"", "'").Replace("”", "'").Replace("“", "'").Replace("‘", "'").Replace("’", "'"));

            //    descr = descr.Replace("`", "'").Replace("´", "'").Replace("\"", "'").Replace("&nbsp;", " ").Replace("?????", "*****").Replace("????", "****").Replace("???", "***").Replace("??", "**");
            descr = descr.Replace("&nbsp;", " ").Replace("?????", "*****").Replace("????", "****").Replace("???", "***").Replace("??", "**");

            descr = descr.Replace("<br>", "\\n");
            descr = descr.Replace("<br/>", "\\n");
            descr = descr.Replace("<p>", "");
            descr = descr.Replace("</p>", "\\n");


            descr = descr.Replace("\\n", "\n");
            descr = descr.Replace("\\t", "\t");

            if (killHtmlTags)
            {
                do
                {
                    int n = descr.IndexOf("<");
                    if (n == -1)
                        break;
                    int m = descr.IndexOf(">", n);
                    if (m == -1)
                        break;
                    if (m + 1 < descr.Length)
                        descr = descr.Substring(0, n) + descr.Substring(m + 1);
                    else
                        descr = descr.Substring(0, n);
                } while (descr.IndexOf("<") != -1);
            }

            //  WriteLog("Description2:" + descr);

            return descr;
        }

        public static string ReadMemoFile(string memofilename)
        {
            string txt = "";
            if (memofilename == "")
                return "";

            memofilename = memofilename.ToLower();
            memofilename = memofilename.Replace(@"v:", Utils.ReadConfigString("v-drive", ""));
            memofilename = memofilename.Replace(@"f:", Utils.ReadConfigString("f-drive", ""));

            memofilename = memofilename.Replace(@"g:", Utils.ReadConfigString("g-drive", @"\\192.168.100.31\group"));

            try
            {
                if (File.Exists(memofilename))
                {
                    //txt = (new StreamReader(memofilename, Encoding.Default)).ReadToEnd();
                    txt = File.ReadAllText(memofilename, Encoding.Default);
                }
            }
            catch //(Exception exception)
            {
            }
            return txt;
        }


        public static Sync ReadSyncTime(string syncTimeFileName)
        {
            Sync sync = new Sync() { LastestSync = DateTime.MinValue };

            try
            { 
                if (File.Exists(syncTimeFileName))
                {
                    using (StreamReader reader = new StreamReader(syncTimeFileName))
                    {
                        XmlSerializer xs = new XmlSerializer(typeof(Sync));
                        try
                        {
                            sync = (Sync)xs.Deserialize(reader);
                        }
                        catch (Exception ex)
                        {
                            Utils.WriteLog($"Error reading {syncTimeFileName} - {ex.Message}");
                        }
                    }
                }
                else
                    Utils.WriteLog($"Sync time file {syncTimeFileName} not found");
            }
            catch (Exception ex)
            {
                Utils.WriteLog($"Sync time file {syncTimeFileName} not found - {ex.Message}");
            }
            return sync;
        }

        public static void WriteSyncTime(Sync sync, string syncTimeFileName)
        {
            DateTime now = DateTime.Now;
            sync.LastestSync = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0);
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(Sync));

            try
            {

                using (FileStream fs = new FileStream(syncTimeFileName, FileMode.Create))
                {
                    XmlSerializer xSer = new XmlSerializer(typeof(Sync));
                    try
                    {
                        xSer.Serialize(fs, sync);
                    }
                    catch
                    {
                        Utils.WriteLog($"Error writing {syncTimeFileName}");
                    }
                }
            }
            catch { }

        }

        public static string ProdNoClean(string prodNo)
        {
            return prodNo.Replace(".", "~").Replace("/", "^").Replace("$", "!").Replace("#", "=").Replace("[", "{").Replace("]", "}");
        }

        public static string ProdNoUnClean(string prodNo)
        {
            return prodNo.Replace("~", ".").Replace("^", "/").Replace("!", "$").Replace("=", "#").Replace("{", "[").Replace("}", "]");
        }
        public static bool ResampleToSize(string filePath, int maxSize)
        {
            try
            {
                Image original = Image.FromFile(filePath);

                if (original.Width > maxSize || original.Height > maxSize)
                {
                    ImageFormat ordFormst = original.RawFormat;

                    Image resized = ResizeImage(original, new Size(maxSize, maxSize), true);

                    try
                    {
                        File.Move(filePath, filePath + GenerateDateStamp(DateTime.Now));

                    }
                    catch (Exception ex)
                    {
                        WriteLog($"ERROR: File.Move() - {ex.Message}");
                        return false;
                    }

                    try
                    {
                        File.Delete(filePath);
                    }
                    catch { }

                    resized.Save(filePath, ordFormst);
                }
            }
            catch (Exception ex)
            {
                WriteLog($"ERROR: ResampleToSize() - {ex.Message}");
                return false;
            }

            return true;

        }

        public static Image ResizeImage(Image image, Size size, bool preserveAspectRatio = true)
        {
            int newWidth;
            int newHeight;
            if (preserveAspectRatio)
            {
                int originalWidth = image.Width;
                int originalHeight = image.Height;
                float percentWidth = (float)size.Width / (float)originalWidth;
                float percentHeight = (float)size.Height / (float)originalHeight;
                float percent = percentHeight < percentWidth ? percentHeight : percentWidth;
                newWidth = (int)(originalWidth * percent);
                newHeight = (int)(originalHeight * percent);
            }
            else
            {
                newWidth = size.Width;
                newHeight = size.Height;
            }
            Image newImage = new Bitmap(newWidth, newHeight);
            using (Graphics graphicsHandle = Graphics.FromImage(newImage))
            {
                graphicsHandle.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphicsHandle.DrawImage(image, 0, 0, newWidth, newHeight);
            }
            return newImage;
        }


        public static string EncodeJSON(string json)
        {
            return JsonConvert.ToString(json).Replace("\"", "");
            // return JsonConvert.ToString(json).Substring(1, json.Length - 2);
        }


        public static bool WriteToMemoFile(string memoPath, string memotext)
        {
            memoPath = memoPath.Replace(@"v:\", @"\\File10\k_4018$\Visma\");
            memoPath = memoPath.Replace(@"V:\", @"\\File10\k_4018$\Visma\");

            try
            {
                if (!File.Exists(memoPath))
                    Directory.CreateDirectory(Path.GetDirectoryName(memoPath));
                else
                    File.Delete(memoPath);

                using (FileStream fileStream = new FileStream(memoPath, FileMode.OpenOrCreate, FileAccess.Write))
                {
                    using (StreamWriter streamWriter = new StreamWriter(fileStream, Encoding.Default))
                    {
                        streamWriter.Write(memotext);
                        streamWriter.Flush();
                        streamWriter.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog($"Error writing memo file {memoPath } - {ex.Message}");
                return false;
            }

            return true;
        }

    }
}
