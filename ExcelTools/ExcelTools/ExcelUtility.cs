using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

namespace ExcelTools
{
    public class ExcelUtility
    { /// <summary>
      /// 获取excel的DataTable
      /// </summary>
      /// <param name="filePath"></param>
      /// <param name="sheetName"></param>
      /// <returns></returns>
        static DataTable GetExcelContent(String filePath, string sheetName)
        {
            if (sheetName == "_xlnm#_FilterDatabase")
                return null;
            DataSet dateSet = new DataSet();
            OleDbConnection connection = null;
            string strExtension = System.IO.Path.GetExtension(filePath);
            switch (strExtension)
            {
                case ".xls":
                    connection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";" + "Extended Properties=\"Excel 8.0;HDR=yes;IMEX=1;\"");
                    break;
                case ".xlsx":
                    connection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";" + "Extended Properties=\"Excel 12.0;HDR=yes;IMEX=1;\"");
                    //此连接可以操作.xls与.xlsx文件 (支持Excel2003 和 Excel2007 的连接字符串) 
                    //"HDR=yes;"是说Excel文件的第一行是列名而不是数，"HDR=No;"正好与前面的相反。"IMEX=1 "如果列中的数据类型不一致，使用"IMEX=1"可必免数据类型冲突。 
                    break;
                default:
                    connection = null;
                    break;
            }
            String commandString = string.Format("SELECT * FROM [{0}$]", sheetName);
            connection.Open();
            using (OleDbCommand command = new OleDbCommand(commandString, connection))
            {
                OleDbCommand objCmd = new OleDbCommand(commandString, connection);
                OleDbDataAdapter myData = new OleDbDataAdapter(commandString, connection);
                myData.Fill(dateSet, sheetName);
                DataTable table = dateSet.Tables[sheetName];
                for (int i = 0; i < table.Rows[0].ItemArray.Length; i++)
                {
                    var cloumnName = table.Rows[0].ItemArray[i].ToString();
                    if (!string.IsNullOrEmpty(cloumnName))
                        table.Columns[i].ColumnName = cloumnName;
                }
                table.Rows.RemoveAt(0);
                connection.Close();
                return table;
            }

        }

        /// <summary>
        /// 获取excel文件里面的所有的工作表名称
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        static List<string> GetExcelSheetNames(string filePath)
        {
            OleDbConnection connection = null;
            System.Data.DataTable dt = null;
            try
            {//获取文件扩展名
                string strExtension = System.IO.Path.GetExtension(filePath);
                switch (strExtension)
                {
                    case ".xls":
                        connection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";" + "Extended Properties=\"Excel 8.0;HDR=yes;IMEX=1;\"");
                        break;
                    case ".xlsx":
                        connection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";" + "Extended Properties=\"Excel 12.0;HDR=yes;IMEX=1;\"");
                        //此连接可以操作.xls与.xlsx文件 (支持Excel2003 和 Excel2007 的连接字符串) 
                        //"HDR=yes;"是说Excel文件的第一行是列名而不是数，"HDR=No;"正好与前面的相反。"IMEX=1 "如果列中的数据类型不一致，使用"IMEX=1"可必免数据类型冲突。 
                        break;
                    default:
                        connection = null;
                        break;
                }
                connection.Open();
                dt = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dt == null)
                {
                    return new List<string>();
                }

                String[] excelSheets = new String[dt.Rows.Count];
                int i = 0;
                foreach (DataRow row in dt.Rows)
                {
                    excelSheets[i] = row["TABLE_NAME"].ToString().Split('$')[0];
                    i++;
                }
                return excelSheets.Distinct().ToList();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return new List<string>();
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                    connection.Dispose();
                }
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
        }

        /// <summary>
        /// excel转换为json
        /// </summary>
        /// <param name="filePath"></param>
        public static JObject ExcelToJson(string filePath)
        {
            List<string> tableNames = GetExcelSheetNames(filePath);
            JObject json = new JObject();
            tableNames.ForEach(tableName =>
            {
                var table = new JArray() as dynamic;
                DataTable dataTable = GetExcelContent(filePath, tableName);
                foreach (DataRow dataRow in dataTable.Rows)
                {
                    dynamic row = new JObject();
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        row.Add(column.ColumnName, dataRow[column.ColumnName].ToString());
                    }
                    table.Add(row);
                }
                json.Add(tableName, table);
            });
            return json;
        }
       public static string ConvertJsonString(string str)
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
