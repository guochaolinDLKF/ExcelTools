namespace ExcelTools
{
    partial class ExcelTools
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.ExcelPath = new System.Windows.Forms.TextBox();
            this.ExcelFileList = new System.Windows.Forms.ListBox();
            this.OutPath = new System.Windows.Forms.TextBox();
            this.JSON = new System.Windows.Forms.CheckBox();
            this.XML = new System.Windows.Forms.CheckBox();
            this.tip = new System.Windows.Forms.Label();
            this.Debug = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(105, 75);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(149, 47);
            this.button1.TabIndex = 0;
            this.button1.Text = "选择Excel文件路径";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 26.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(314, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(182, 35);
            this.label1.TabIndex = 1;
            this.label1.Text = "Excel工具";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(105, 290);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(148, 54);
            this.button2.TabIndex = 3;
            this.button2.Text = "选择要输出的路径";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(462, 407);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(121, 38);
            this.button3.TabIndex = 6;
            this.button3.Text = "一键转换";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // ExcelPath
            // 
            this.ExcelPath.Location = new System.Drawing.Point(303, 89);
            this.ExcelPath.Name = "ExcelPath";
            this.ExcelPath.Size = new System.Drawing.Size(475, 21);
            this.ExcelPath.TabIndex = 7;
            this.ExcelPath.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // ExcelFileList
            // 
            this.ExcelFileList.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ExcelFileList.FormattingEnabled = true;
            this.ExcelFileList.ItemHeight = 16;
            this.ExcelFileList.Location = new System.Drawing.Point(303, 130);
            this.ExcelFileList.Name = "ExcelFileList";
            this.ExcelFileList.Size = new System.Drawing.Size(568, 148);
            this.ExcelFileList.TabIndex = 8;
            // 
            // OutPath
            // 
            this.OutPath.Location = new System.Drawing.Point(303, 308);
            this.OutPath.Name = "OutPath";
            this.OutPath.Size = new System.Drawing.Size(475, 21);
            this.OutPath.TabIndex = 12;
            this.OutPath.TextChanged += new System.EventHandler(this.textBox1_TextChanged_2);
            // 
            // JSON
            // 
            this.JSON.AutoSize = true;
            this.JSON.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.JSON.Location = new System.Drawing.Point(105, 204);
            this.JSON.Name = "JSON";
            this.JSON.Size = new System.Drawing.Size(54, 18);
            this.JSON.TabIndex = 15;
            this.JSON.Text = "JSON";
            this.JSON.UseVisualStyleBackColor = true;
            this.JSON.CheckedChanged += new System.EventHandler(this.JSON_CheckedChanged);
            // 
            // XML
            // 
            this.XML.AutoSize = true;
            this.XML.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.XML.Location = new System.Drawing.Point(211, 204);
            this.XML.Name = "XML";
            this.XML.Size = new System.Drawing.Size(47, 18);
            this.XML.TabIndex = 16;
            this.XML.Text = "XML";
            this.XML.UseVisualStyleBackColor = true;
            this.XML.CheckedChanged += new System.EventHandler(this.checkBox2_CheckedChanged_1);
            // 
            // tip
            // 
            this.tip.AutoSize = true;
            this.tip.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tip.Location = new System.Drawing.Point(93, 173);
            this.tip.Name = "tip";
            this.tip.Size = new System.Drawing.Size(161, 16);
            this.tip.TabIndex = 17;
            this.tip.Text = "请选择要转换的格式";
            // 
            // Debug
            // 
            this.Debug.FormattingEnabled = true;
            this.Debug.ItemHeight = 12;
            this.Debug.Location = new System.Drawing.Point(26, 378);
            this.Debug.Name = "Debug";
            this.Debug.Size = new System.Drawing.Size(348, 160);
            this.Debug.TabIndex = 18;
            // 
            // ExcelTools
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(906, 550);
            this.Controls.Add(this.Debug);
            this.Controls.Add(this.tip);
            this.Controls.Add(this.XML);
            this.Controls.Add(this.JSON);
            this.Controls.Add(this.OutPath);
            this.Controls.Add(this.ExcelFileList);
            this.Controls.Add(this.ExcelPath);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Name = "ExcelTools";
            this.Text = "ExcelTools";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.TextBox ExcelPath;
        private System.Windows.Forms.ListBox ExcelFileList;
        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Windows.Forms.TextBox OutPath;
        private System.Windows.Forms.CheckBox JSON;
        private System.Windows.Forms.CheckBox XML;
        private System.Windows.Forms.Label tip;
        private System.Windows.Forms.ListBox Debug;
    }
}

