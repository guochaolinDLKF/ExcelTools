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
        public static DataTable GetExcelContent(String filePath, string sheetName)
        {
            if (sheetName == "_xlnm#_FilterDatabase")
                return null;
            DataSet dateSet = new DataSet();
            string strExtension = System.IO.Path.GetExtension(filePath);
            string strConn = "";
            switch (strExtension)
            {
                case ".xls":
                    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";" +
                              "Extended Properties=\"Excel 8.0;HDR=yes;IMEX=1;\"";
                    break;
                case ".xlsx":
                    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";" +
                              "Extended Properties=\"Excel 12.0;HDR=yes;IMEX=1;\"";
                    //此连接可以操作.xls与.xlsx文件 (支持Excel2003 和 Excel2007 的连接字符串) 
                    //"HDR=yes;"是说Excel文件的第一行是列名而不是数，"HDR=No;"正好与前面的相反。"IMEX=1 "如果列中的数据类型不一致，使用"IMEX=1"可必免数据类型冲突。 
                    break;
                default:
                    break;
            }
            String commandString = string.Format("SELECT * FROM [{0}$]", sheetName);


            using (OleDbConnection connection = new OleDbConnection(strConn))
            {
                connection.Open();
                using (OleDbDataAdapter myData = new OleDbDataAdapter(commandString, connection))
                {
                    myData.Fill(dateSet, sheetName);
                    DataTable table = dateSet.Tables[sheetName];
                    for (int i = 0; i < table.Rows[0].ItemArray.Length; i++)
                    {
                        var cloumnName = table.Rows[0].ItemArray[i].ToString();
                        if (!string.IsNullOrEmpty(cloumnName))
                            table.Columns[i].ColumnName = cloumnName;
                    }
                    return table;
                }
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

        public static JArray ExcelSingleSheetToJson(string filePath, Action<string> call)
        {
            List<string> tableNames = GetExcelSheetNames(filePath);
            //JObject json = new JObject();
            JArray table = new JArray();
            //JObject table = new JObject();
            //for (int h = 0; h < tableNames.Count; h++)
            //{

            DataTable dataTable = GetExcelContent(filePath, tableNames[0]);
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                JObject row = new JObject();
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    if (dataTable.Columns[j].ColumnName ==
                        dataTable.Rows[i][dataTable.Columns[j].ColumnName].ToString())
                    {
                        continue;
                    }
                    if (i == 1) continue;
                    int index = -1;
                    float flo = 0.0f;
                    double doub = 0.0f;
                    bool boolStr = false;
                    string typeStr = dataTable.Rows[1][dataTable.Columns[j].ColumnName].ToString().ToLower();
                    if (typeStr == "") continue;
                    if(dataTable.Rows[i][dataTable.Columns[j].ColumnName].ToString()=="")continue;
                    switch (typeStr)
                    {
                        case "int":
                            index = dataTable.Rows[i][dataTable.Columns[j].ColumnName].ToString().ToInt();
                            if (index != -1)
                                row.Add(dataTable.Columns[j].ColumnName, index);
                            else
                            {
                                call(string.Format("表格：{0}中int类型转换失败，请检查第{1}行第{2}列数据", tableNames[0], i+2, j));
                                return new JArray();
                            }
                            break;
                        case "string":
                            row.Add(dataTable.Columns[j].ColumnName, dataTable.Rows[i][dataTable.Columns[j].ColumnName].ToString());
                            break;
                        case "float":
                            flo = dataTable.Rows[i][dataTable.Columns[j].ColumnName].ToString().ToLower().ToFloat();
                            if (flo != -1.0f)
                                row.Add(dataTable.Columns[j].ColumnName, flo);
                            else
                            {
                                call(string.Format("表格：{0}中float类型转换失败，请检查第{1}行第{2}列数据", tableNames[0], i + 2, j));
                                return new JArray();
                            }
                            break;
                        case "double":
                            doub = dataTable.Rows[i][dataTable.Columns[j].ColumnName].ToString().ToDouble();
                            if (doub != -1.0f)
                                row.Add(dataTable.Columns[j].ColumnName, doub);
                            else
                            {
                                call(string.Format("表格：{0}中double类型转换失败，请检查第{1}行第{2}列数据", tableNames[0], i + 2, j));
                                return new JArray();
                            }
                            break;
                        default:
                            call(string.Format("表格：{0}遇到无法转换的类型：{1}", tableNames[0], dataTable.Rows[1][dataTable.Columns[j].ColumnName].ToString()));
                            return new JArray();
                    }

                }
                if (!row.HasValues)
                {
                    continue;
                }
                table.Add(row);
            }
            //json.Add(tableNames[h], table);
            // }
            return table;
        }
        /// <summary>
        /// excel转换为json
        /// </summary>
        /// <param name="filePath"></param>
        public static JObject ExcelMutileSheetToJson(string filePath, Action<string> call)
        {
            List<string> tableNames = GetExcelSheetNames(filePath);
            JObject json = new JObject();
            for (int h = 0; h < tableNames.Count; h++)
            {
                JArray table = new JArray();
                DataTable dataTable = GetExcelContent(filePath, tableNames[h]);
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    JObject row = new JObject();
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        if (dataTable.Columns[j].ColumnName ==
                            dataTable.Rows[i][dataTable.Columns[j].ColumnName].ToString())
                        {
                            continue;
                        }
                        if (i == 1) continue;
                        int index = -1;
                        double doub = 0.0f;
                        string typeStr = dataTable.Rows[1][dataTable.Columns[j].ColumnName].ToString().ToLower();
                        if (typeStr == "") continue;
                        switch (typeStr)
                        {
                            case "int":
                                index = dataTable.Rows[i][dataTable.Columns[j].ColumnName].ToString().ToInt();
                                if (index != -1)
                                    row.Add(dataTable.Columns[j].ColumnName, index);
                                break;
                            case "string":
                                row.Add(dataTable.Columns[j].ColumnName, dataTable.Rows[i][dataTable.Columns[j].ColumnName].ToString());
                                break;
                            case "double":
                                doub = dataTable.Rows[i][dataTable.Columns[j].ColumnName].ToString().ToDouble();
                                if (doub != 0.0f)
                                    row.Add(dataTable.Columns[j].ColumnName, doub);
                                break;
                            default:
                                call(string.Format("表格：{0}遇到无法转换的类型：{1}", tableNames[h], dataTable.Rows[1][dataTable.Columns[j].ColumnName].ToString()));
                                return new JObject();
                        }

                    }
                    if (!row.HasValues)
                    {
                        continue;
                    }
                    table.Add(row);
                }
                json.Add(tableNames[h], table);
            }
            return json;
        }

        public static JObject ExcelToJsonArray(string filePath)
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
