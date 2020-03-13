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
   public class JsonUtils
    {
        /// <summary>
        /// DataTable转Json
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static string DataTableToJson(DataTable dtb)
        {
            string JsonString = string.Empty;
            if (dtb.Rows.Count > 0)
            {
                ArrayList dic = new ArrayList();
                foreach (DataRow row in dtb.Rows)
                {
                    Dictionary<string, object> drow = new Dictionary<string, object>();
                    foreach (DataColumn col in dtb.Columns)
                    {
                        drow.Add(col.ColumnName, row[col.ColumnName]);
                    }
                    dic.Add(drow);
                }

                JsonString = JsonConvert.SerializeObject(dic);
                return ConvertJsonString(JsonString);
            }

            return JsonString;
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
