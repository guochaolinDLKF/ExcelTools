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
            JSON.Checked = true;
        }
        /// <summary>
        /// 选择excel路径
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog f = new FolderBrowserDialog();
            f.Description = "选择Excel存放目录";
            f.RootFolder = Environment.SpecialFolder.Desktop;
            if (f.ShowDialog() == DialogResult.OK)
            {
                String DirPath = f.SelectedPath;

                this.ExcelPath.Text = DirPath + "\\";//G:\新建文件夹

                SearchFileAdd(DirPath);

            }
        }

        void SearchFileAdd(string dirPath)
        {
            ExcelFileList.Items.Clear();
            DirectoryInfo theFolder = new DirectoryInfo(dirPath);
            FileInfo[] fileInfo = theFolder.GetFiles();
            foreach (FileInfo NextFile in fileInfo) //遍历文件
            {
                if (NextFile.Extension.Equals(".xls") ||
                    NextFile.Extension.Equals(".xlsx"))
                {
                    ExcelFileList.Items.Add("ExcelPath/" + NextFile.Name);
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
            if (f.ShowDialog() == DialogResult.OK)
            {
                String DirPath = f.SelectedPath;
                this.OutPath.Text = DirPath + "\\";//G:\新建文件夹

            }
        }
        /// <summary>
        /// 一键转换
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            if (!JSON.Checked && !XML.Checked)
            {
                MessageBox.Show("请至少选择一项转换");
            }

            if (OutPath.Text == "")
            {
                MessageBox.Show("请选择输出目录");
            }
            if (ExcelPath.Text == "")
            {
                MessageBox.Show("请选择Excel目录");
            }
            //转为json
            if (JSON.Checked)
            {
                ExcelToJSON();
            }
            else if (XML.Checked)//转为XML
            {

            }
        }

        void ExcelToJSON()
        {
            //如果不存在导出目录，则创建
            if (!Directory.Exists(OutPath.Text))
            {
                Directory.CreateDirectory(OutPath.Text);
            }

            if (ExcelFileList.Items.Count <= 0)
            {
                SearchFileAdd("ExcelPath/");
            }
            for (int i = 0; i < ExcelFileList.Items.Count; i++)
            {
                JObject obj = ExcelUtility.ExcelToJson(ExcelFileList.Items[i].ToString());
                string fileName = Path.GetFileName(ExcelFileList.Items[i].ToString());

                fileName = fileName.Split('.')[0];
                fileName = OutPath.Text + fileName + ".json";
                if (File.Exists(fileName))
                {
                    File.Delete(fileName);
                }
                using (FileStream fs = new FileStream(fileName, FileMode.CreateNew, FileAccess.Write))
                {
                    byte[] data = Encoding.UTF8.GetBytes(ExcelUtility.ConvertJsonString(obj.ToString()));
                    fs.Write(data, 0, data.Length);
                }
            }
            MessageBox.Show("转换完成");
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

        private void JSON_CheckedChanged(object sender, EventArgs e)
        {
            if (JSON.Checked)
            {
                XML.Checked = false;
            }

        }

        private void checkBox2_CheckedChanged_1(object sender, EventArgs e)
        {
            if (XML.Checked)
            {
                JSON.Checked = false;
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
