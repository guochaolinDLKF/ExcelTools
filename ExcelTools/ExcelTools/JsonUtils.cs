using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelTools
{
   public static class JsonUtils
    {
        public static int ToInt(this string str)
        {
            int index = -1;
            if (int.TryParse(str, out index))
            {
                return index;
            }
            return index;
        }
        public static Int64 ToInt64(this string str)
        {
            Int64 index = -1;
            if (Int64.TryParse(str, out index))
            {
                return index;
            }
            return index;
        }
       
        public static float ToFloat(this string str)
        {
            float res = -1f;
            if (float.TryParse(str, out res))
            {
                return res;
            }
            return res;
        }
        public static bool ToBool(this string str,out int res)
        {
            if (string.IsNullOrEmpty(str))
            {
                res = 1;
                return false;
            }
            bool temp = false;
            bool.TryParse(str, out temp);
            if (!temp)
            {
                if (str.Trim().Equals("1"))
                {
                    temp = true;
                }
            }
            res = 0;
            return temp;
        }
        public static double ToDouble(this string str)
        {
            double res = -1.0f;
            if (double.TryParse(str, out res))
            {
                return res;
            }
            return res;
        }
        static string ConvertJsonString(string str)
        {
            //格式化json字符串
            JsonSerializer serializer = new JsonSerializer();
            TextReader tr = new StringReader(str);
            JsonTextReader jtr = new JsonTextReader(tr);
            object obj = serializer.Deserialize(jtr);
            if (obj != null)
            {
                StringWriter textWriter = new StringWriter();
                JsonTextWriter jsonWriter = new JsonTextWriter(textWriter)
                {
                    Formatting = Formatting.Indented,
                    Indentation = 4,
                    IndentChar = ' '
                };
                serializer.Serialize(jsonWriter, obj);
                return textWriter.ToString();
            }
            else
            {
                return str;
            }
        }
    }
}
