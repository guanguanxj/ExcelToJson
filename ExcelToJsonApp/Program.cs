using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using Microsoft.International.Converters.PinYinConverter;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Collections.ObjectModel;

namespace ExcelToJsonApp
{
    public class Program
    {
        static void Main(string[] args)
        {
            //C:读取excel 
            //E:read the excel file
            string filename = "ViewExcel.xlsx";
            DataSet ds = ImportExcelToDataSet(filename);
            DataTable dt = ds.Tables[0];
            var rowscount = dt.Rows.Count;
            ViewValue view;
            List<ViewValue> viewList = new List<ViewValue>();
            for (var i = 0; i < rowscount; i++)
            {
                view = new ViewValue();
                view.Text = dt.Rows[i][0].ToString();
                view.Id = Convert.ToInt32(dt.Rows[i][1]);
                view.Type = dt.Rows[i][2].ToString();
                viewList.Add(view);
            }

            //C:汉字转拼音
            //E:Convert Chinese characters to Chinese phonetic alphabet which is named "PinYin".
            var firstCharList = viewList.AsEnumerable()
               .OrderBy(x => ConvertToPinYin(x.Text, false))
               .Select(x => GetFirstChar(ConvertToPinYin(x.Text, false))).Distinct();

            List<ViewJsonResult> results = new List<ViewJsonResult>();

            foreach (var firstChar in firstCharList)
            {
                ViewJsonResult result = new ViewJsonResult();
                result.Index = firstChar;
                result.Value = viewList.Where(x => SpecialChinese(x.Text, firstChar))
                    .OrderBy(x => ConvertToPinYin(x.Text, false)).ToList();
                results.Add(result);
            }
            //输出文件
            //output the result
            string fileContent = ObjectToJson(results);
            WriteFile("D://result.txt", fileContent);

        }
        /*some special Chinese characteres convert to correct phonetic alphabet.
         Such as "广", which first phonetic alphabet is "G",but the Microsoft.International.Converters.PinYinConverter return  "A".
         * */
        public static bool SpecialChinese(string text, string firstChar)
        {
            bool flag = false;

            flag = (!(firstChar == "A" && text.StartsWith("广")) && ConvertToPinYin(text, false).StartsWith(firstChar))
                  && (!(firstChar == "B" && text.StartsWith("屏")) && ConvertToPinYin(text, false).StartsWith(firstChar))
                  && (!(firstChar == "D" && text.StartsWith("石")) && ConvertToPinYin(text, false).StartsWith(firstChar))
                  && (!(firstChar == "G" && text.StartsWith("邢")) && ConvertToPinYin(text, false).StartsWith(firstChar))
                  && (!(firstChar == "G" && text.StartsWith("合")) && ConvertToPinYin(text, false).StartsWith(firstChar))
                  && (!(firstChar == "G" && text.StartsWith("句")) && ConvertToPinYin(text, false).StartsWith(firstChar))
                  && (!(firstChar == "H" && text.StartsWith("许")) && ConvertToPinYin(text, false).StartsWith(firstChar))
                  && (!(firstChar == "J" && text.StartsWith("齐")) && ConvertToPinYin(text, false).StartsWith(firstChar))
                  && (!(firstChar == "M" && text.StartsWith("万")) && ConvertToPinYin(text, false).StartsWith(firstChar))
                  && (!(firstChar == "M" && text.StartsWith("无")) && ConvertToPinYin(text, false).StartsWith(firstChar))
                  && (!(firstChar == "S" && text.StartsWith("汤")) && ConvertToPinYin(text, false).StartsWith(firstChar))
                  || (firstChar == "G" && text.StartsWith("广"))
                  || (firstChar == "T" && text.StartsWith("汤"))
                  || (firstChar == "P" && text.StartsWith("屏"))
                  || (firstChar == "S" && text.StartsWith("石"))
                  || (firstChar == "X" && text.StartsWith("邢"))
                  || (firstChar == "H" && text.StartsWith("合"))
                  || (firstChar == "J" && text.StartsWith("句"))
                  || (firstChar == "X" && text.StartsWith("许"))
                  || (firstChar == "Q" && text.StartsWith("齐"))
                  || (firstChar == "W" && text.StartsWith("万"))
                  || (firstChar == "W" && text.StartsWith("无"));
            return flag;
        }
        /// <summary> 
        /// C:判断是否是中文字 
        /// E:is Chinese char or not
        /// </summary> 
        /// <param name="ch">要检查的字节:ch ,parm :ch</param> 
        /// <returns>True:是 yes;false:否 no</returns> 
        public static bool IsValidChar(char ch)
        {
            return ChineseChar.IsValidChar(ch);
        }
        //Convert Chinese characteres to PinYin
        public static string ConvertToPinYin(string chineseStr, bool includeTone)
        {
            if (chineseStr == null)
                throw new ArgumentNullException("chineseStr");
            char[] charArray = chineseStr.ToCharArray();
            ChineseChar chineseChar = null;
            StringBuilder sb = new StringBuilder();
            foreach (char c in charArray)
            {
                if (IsValidChar(c))
                {
                    chineseChar = new ChineseChar(c);
                    ReadOnlyCollection<string> pyColl = chineseChar.Pinyins;
                    foreach (string py in pyColl)
                    {
                        if (!string.IsNullOrEmpty(py))
                        {
                            sb.Append(py);
                            break;
                        }
                    }
                }
                else
                {
                    sb.Append(c);
                }
            }
            if (!includeTone)
            {
                StringBuilder sb2 = new StringBuilder();
                foreach (char c in sb.ToString())
                {
                    if (!char.IsNumber(c))
                        sb2.Append(c);
                }
                return sb2.ToString();
            }
            return sb.ToString();
        }


        public static string GetFirstChar(string text)
        {
            return !string.IsNullOrEmpty(text) ? text.Substring(0, 1) : "";
        }

        public static string ObjectToJson(object obj)
        {
            DataContractJsonSerializer serializer = new DataContractJsonSerializer(obj.GetType());
            MemoryStream stream = new MemoryStream();
            serializer.WriteObject(stream, obj);
            byte[] dataBytes = new byte[stream.Length];
            stream.Position = 0;
            stream.Read(dataBytes, 0, (int)stream.Length);
            return Encoding.UTF8.GetString(dataBytes);
        }

        private static void WriteFile(string path, string content)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
            FileStream fs = new FileStream(path, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            sw.Write(content);
            sw.Flush();
            sw.Close();
            fs.Close();
        }

        private static DataSet ImportExcelToDataSet(string FilePath)
        {
            DataSet ds = new DataSet();
            try
            {
                string strConn;
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FilePath + ";Extended Properties=Excel 12.0;";
                OleDbConnection conn = new OleDbConnection(strConn);
                conn.Open();
                string sql = "SELECT Originaltext, TextId, Type FROM [result$]";
                OleDbDataAdapter myCommand = new OleDbDataAdapter(sql, strConn);

                myCommand.Fill(ds);
                conn.Close();
            }
            catch (Exception ex)
            {
                throw new Exception("Excel Invaild," + ex.Message);
            }
            return ds;
        }

    }
}
