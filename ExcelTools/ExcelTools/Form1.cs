using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelTools
{
    public enum OutSheetNum
    {
        Single,
        Mutile
    }
    public partial class ExcelTools : Form
    {

        public ExcelTools()
        {

            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ExcelPath.Text = "ExcelPath/";
            OutPath.Text = "OutPath/";
        }
        /// <summary>
        /// 选择excel路径
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fb = new FolderBrowserDialog();
            fb.RootFolder = Environment.SpecialFolder.MyComputer;
            fb.Description = "选择Excel存放目录";
            if (fb.ShowDialog() == DialogResult.OK)
            {
                String DirPath = fb.SelectedPath;

                this.ExcelPath.Text = DirPath + "\\";//G:\新建文件夹

                SearchExcelFileAdd(DirPath);

            }
        } 
        void SearchJsonFileAdd(string dirPath)
        {
            JsonFileList.Items.Clear();
            DirectoryInfo theFolder = new DirectoryInfo(dirPath);
            FileInfo[] fileInfo = theFolder.GetFiles();
            foreach (FileInfo NextFile in fileInfo) //遍历文件
            {
                if (NextFile.Extension.Equals(".json"))
                {
                    JsonFileList.Items.Add(OutPath.Text + NextFile.Name);
                }
            }
        }
        void SearchExcelFileAdd(string dirPath)
        {
            ExcelFileList.Items.Clear();
            DirectoryInfo theFolder = new DirectoryInfo(dirPath);
            FileInfo[] fileInfo = theFolder.GetFiles();
            foreach (FileInfo NextFile in fileInfo) //遍历文件
            {
                if (NextFile.Extension.Equals(".xls") ||
                    NextFile.Extension.Equals(".xlsx"))
                {
                    string path = dirPath.Replace("\\", "/");
                    ExcelFileList.Items.Add(path +"/"+ NextFile.Name);
                }
            }
        }
        /// <summary>
        /// 选择输出的路径
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog f = new FolderBrowserDialog();
            f.RootFolder = Environment.SpecialFolder.MyComputer;
            f.Description = "选择Json文件存放目录";
            if (f.ShowDialog() == DialogResult.OK)
            {
                String DirPath = f.SelectedPath;
                this.OutPath.Text = DirPath + "\\";//G:\新建文件夹
                
            }
        }
        /// <summary>
        /// 一键转换C#
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            Debug.Items.Clear();
            if (OutPath.Text == "")
            {
                MessageBox.Show("请选择输出目录");
            }
            if (ExcelPath.Text == "")
            {
                MessageBox.Show("请选择Excel目录");
            }
            MessageBox.Show("功能未开发");
        }
        /// <summary>
        /// 一键转换Json
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            Debug.Items.Clear();
            if (OutPath.Text == "")
            {
                MessageBox.Show("请选择输出目录");
            }
            if (ExcelPath.Text == "")
            {
                MessageBox.Show("请选择Excel目录");
            }
            //转为json
            ExcelToJSON();
        }

        void ExcelToJSON()
        {
            //如果存在导出目录，则创建
            if (!Directory.Exists(OutPath.Text))
            {
                Directory.CreateDirectory(OutPath.Text);
            }
            if (ExcelFileList.Items.Count <= 0)
            {
                SearchExcelFileAdd("ExcelPath");
            }
            for (int i = 0; i < ExcelFileList.Items.Count; i++)
            {
                JArray arr = new JArray();
                arr = ExcelUtility.ExcelSingleSheetToJson(ExcelFileList.Items[i].ToString().Replace("\\", "/"), ToError);
                //如果没有值，则不做操作
                if (!arr.HasValues) return;

                string fileName = Path.GetFileName( ExcelFileList.Items[i].ToString());

                fileName = fileName.Split('.')[0];
                fileName = OutPath.Text + fileName + ".json";
                if (File.Exists(fileName))
                {
                    File.Delete(fileName);
                }
                using (FileStream fs = new FileStream(fileName, FileMode.CreateNew, FileAccess.Write))
                {
                    string js = ExcelUtility.ConvertJsonString(arr.ToString());

                    if (js.Equals("")) return;
                    byte[] data = Encoding.UTF8.GetBytes(js);
                    fs.Write(data, 0, data.Length);
                }
            }
            SearchJsonFileAdd(OutPath.Text);
            MessageBox.Show("转换完成");
        }

        void ToError(string error)
        {
            Debug.Items.Add(error);
        }
        /// <summary>
        /// 选择Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }


        private void textBox1_TextChanged_2(object sender, EventArgs e)
        {

        }




        private void label1_Click(object sender, EventArgs e)
        {

        }



        private void checkBox1_CheckedChanged_1(object sender, EventArgs e)
        {

        }

        private void JsonFileList_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Debug_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void CSharpFileList_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


    }
}
