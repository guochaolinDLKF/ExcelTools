using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelTools;
using Newtonsoft.Json.Linq;

namespace ExcelTest
{
    class Program
    {
        static void Main(string[] args)
        {
           JObject obj= ExcelUtility.ExcelToJson("fish_s_db.xls");

           Console.WriteLine(obj.ToString());
           Console.ReadKey();
        }
    }
}
