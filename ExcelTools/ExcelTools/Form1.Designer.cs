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
            this.tip = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.JsonFileList = new System.Windows.Forms.ListBox();
            this.label3 = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.CSharpFileList = new System.Windows.Forms.ListBox();
            this.label5 = new System.Windows.Forms.Label();
            this.Debug = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(148, 84);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(149, 29);
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
            this.label1.Size = new System.Drawing.Size(453, 35);
            this.label1.TabIndex = 1;
            this.label1.Text = "批量转换C#、Json格式工具";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(149, 131);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(148, 29);
            this.button2.TabIndex = 3;
            this.button2.Text = "选择要输出的路径";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(376, 172);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(121, 41);
            this.button3.TabIndex = 6;
            this.button3.Text = "一键转换Json";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // ExcelPath
            // 
            this.ExcelPath.Location = new System.Drawing.Point(303, 89);
            this.ExcelPath.Name = "ExcelPath";
            this.ExcelPath.Size = new System.Drawing.Size(802, 21);
            this.ExcelPath.TabIndex = 7;
            this.ExcelPath.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // ExcelFileList
            // 
            this.ExcelFileList.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ExcelFileList.FormattingEnabled = true;
            this.ExcelFileList.ItemHeight = 16;
            this.ExcelFileList.Location = new System.Drawing.Point(42, 260);
            this.ExcelFileList.Name = "ExcelFileList";
            this.ExcelFileList.Size = new System.Drawing.Size(1063, 84);
            this.ExcelFileList.TabIndex = 8;
            // 
            // OutPath
            // 
            this.OutPath.Location = new System.Drawing.Point(303, 136);
            this.OutPath.Name = "OutPath";
            this.OutPath.Size = new System.Drawing.Size(802, 21);
            this.OutPath.TabIndex = 12;
            this.OutPath.TextChanged += new System.EventHandler(this.textBox1_TextChanged_2);
            // 
            // tip
            // 
            this.tip.AutoSize = true;
            this.tip.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tip.Location = new System.Drawing.Point(146, 182);
            this.tip.Name = "tip";
            this.tip.Size = new System.Drawing.Size(161, 16);
            this.tip.TabIndex = 17;
            this.tip.Text = "请选择要转换的格式";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(40, 233);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(83, 12);
            this.label2.TabIndex = 22;
            this.label2.Text = "Excel文件列表";
            // 
            // JsonFileList
            // 
            this.JsonFileList.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.JsonFileList.FormattingEnabled = true;
            this.JsonFileList.ItemHeight = 16;
            this.JsonFileList.Location = new System.Drawing.Point(42, 388);
            this.JsonFileList.Name = "JsonFileList";
            this.JsonFileList.Size = new System.Drawing.Size(1063, 100);
            this.JsonFileList.TabIndex = 23;
            this.JsonFileList.SelectedIndexChanged += new System.EventHandler(this.JsonFileList_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(40, 361);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(77, 12);
            this.label3.TabIndex = 24;
            this.label3.Text = "Json文件列表";
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(523, 172);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(121, 41);
            this.button4.TabIndex = 25;
            this.button4.Text = "一键转换C#";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(40, 503);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 12);
            this.label4.TabIndex = 27;
            this.label4.Text = "C#文件列表";
            // 
            // CSharpFileList
            // 
            this.CSharpFileList.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CSharpFileList.FormattingEnabled = true;
            this.CSharpFileList.ItemHeight = 16;
            this.CSharpFileList.Location = new System.Drawing.Point(42, 518);
            this.CSharpFileList.Name = "CSharpFileList";
            this.CSharpFileList.Size = new System.Drawing.Size(1063, 100);
            this.CSharpFileList.TabIndex = 26;
            this.CSharpFileList.SelectedIndexChanged += new System.EventHandler(this.CSharpFileList_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(40, 630);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(53, 12);
            this.label5.TabIndex = 29;
            this.label5.Text = "错误日志";
            // 
            // Debug
            // 
            this.Debug.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Debug.FormattingEnabled = true;
            this.Debug.ItemHeight = 16;
            this.Debug.Location = new System.Drawing.Point(42, 645);
            this.Debug.Name = "Debug";
            this.Debug.Size = new System.Drawing.Size(1063, 100);
            this.Debug.TabIndex = 28;
            this.Debug.SelectedIndexChanged += new System.EventHandler(this.Debug_SelectedIndexChanged);
            // 
            // ExcelTools
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1208, 804);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.Debug);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.CSharpFileList);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.JsonFileList);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.tip);
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
        private System.Windows.Forms.Label tip;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ListBox JsonFileList;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ListBox CSharpFileList;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ListBox Debug;
    }
}

