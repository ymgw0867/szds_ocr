namespace SZDS_TIMECARD.prePrint
{
    partial class prePrint
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(prePrint));
            this.label1 = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.linkLblOn = new System.Windows.Forms.LinkLabel();
            this.linkLblOff = new System.Windows.Forms.LinkLabel();
            this.chkWhite = new System.Windows.Forms.CheckBox();
            this.linkPrn = new System.Windows.Forms.LinkLabel();
            this.linkRtn = new System.Windows.Forms.LinkLabel();
            this.label3 = new System.Windows.Forms.Label();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.label4 = new System.Windows.Forms.Label();
            this.cmbPrnName = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.Location = new System.Drawing.Point(13, 75);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(85, 23);
            this.label1.TabIndex = 0;
            this.label1.Text = "印刷日付：";
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.CustomFormat = "yyyy/MM/dd(ddd)";
            this.dateTimePicker1.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.dateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePicker1.Location = new System.Drawing.Point(104, 72);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(160, 30);
            this.dateTimePicker1.TabIndex = 2;
            this.dateTimePicker1.ValueChanged += new System.EventHandler(this.dateTimePicker1_ValueChanged);
            // 
            // dateTimePicker2
            // 
            this.dateTimePicker2.CustomFormat = "yyyy/MM/dd(ddd)";
            this.dateTimePicker2.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.dateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePicker2.Location = new System.Drawing.Point(299, 72);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.Size = new System.Drawing.Size(160, 30);
            this.dateTimePicker2.TabIndex = 3;
            this.dateTimePicker2.ValueChanged += new System.EventHandler(this.dateTimePicker2_ValueChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(268, 75);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(24, 19);
            this.label2.TabIndex = 3;
            this.label2.Text = "～";
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(17, 107);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 21;
            this.dataGridView1.Size = new System.Drawing.Size(442, 288);
            this.dataGridView1.TabIndex = 4;
            // 
            // linkLblOn
            // 
            this.linkLblOn.AutoSize = true;
            this.linkLblOn.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.linkLblOn.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.linkLblOn.LinkColor = System.Drawing.Color.Blue;
            this.linkLblOn.Location = new System.Drawing.Point(267, 399);
            this.linkLblOn.Name = "linkLblOn";
            this.linkLblOn.Size = new System.Drawing.Size(80, 18);
            this.linkLblOn.TabIndex = 6;
            this.linkLblOn.TabStop = true;
            this.linkLblOn.Text = "全てチェック";
            this.linkLblOn.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLblOn_LinkClicked);
            // 
            // linkLblOff
            // 
            this.linkLblOff.AutoSize = true;
            this.linkLblOff.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.linkLblOff.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.linkLblOff.LinkColor = System.Drawing.Color.Blue;
            this.linkLblOff.Location = new System.Drawing.Point(355, 399);
            this.linkLblOff.Name = "linkLblOff";
            this.linkLblOff.Size = new System.Drawing.Size(104, 18);
            this.linkLblOff.TabIndex = 7;
            this.linkLblOff.TabStop = true;
            this.linkLblOff.Text = "全てチェックオフ";
            this.linkLblOff.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLblOff_LinkClicked);
            // 
            // chkWhite
            // 
            this.chkWhite.AutoSize = true;
            this.chkWhite.Font = new System.Drawing.Font("メイリオ", 9.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.chkWhite.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.chkWhite.Location = new System.Drawing.Point(18, 397);
            this.chkWhite.Name = "chkWhite";
            this.chkWhite.Size = new System.Drawing.Size(80, 24);
            this.chkWhite.TabIndex = 5;
            this.chkWhite.Text = "白紙印刷";
            this.chkWhite.UseVisualStyleBackColor = true;
            this.chkWhite.CheckedChanged += new System.EventHandler(this.chkWhite_CheckedChanged);
            // 
            // linkPrn
            // 
            this.linkPrn.Font = new System.Drawing.Font("メイリオ", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.linkPrn.Image = ((System.Drawing.Image)(resources.GetObject("linkPrn.Image")));
            this.linkPrn.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.linkPrn.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.linkPrn.LinkColor = System.Drawing.Color.Blue;
            this.linkPrn.Location = new System.Drawing.Point(305, 430);
            this.linkPrn.Name = "linkPrn";
            this.linkPrn.Size = new System.Drawing.Size(87, 19);
            this.linkPrn.TabIndex = 8;
            this.linkPrn.TabStop = true;
            this.linkPrn.Text = "印刷実行";
            this.linkPrn.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.linkPrn.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkPrn_LinkClicked);
            // 
            // linkRtn
            // 
            this.linkRtn.Font = new System.Drawing.Font("メイリオ", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.linkRtn.Image = ((System.Drawing.Image)(resources.GetObject("linkRtn.Image")));
            this.linkRtn.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.linkRtn.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.linkRtn.LinkColor = System.Drawing.Color.Blue;
            this.linkRtn.Location = new System.Drawing.Point(400, 429);
            this.linkRtn.Name = "linkRtn";
            this.linkRtn.Size = new System.Drawing.Size(59, 19);
            this.linkRtn.TabIndex = 9;
            this.linkRtn.TabStop = true;
            this.linkRtn.Text = "終了";
            this.linkRtn.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.linkRtn.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel4_LinkClicked);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("メイリオ", 8.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label3.Location = new System.Drawing.Point(14, 432);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(42, 18);
            this.label3.TabIndex = 14;
            this.label3.Text = "label3";
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripProgressBar1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 462);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(476, 22);
            this.statusStrip1.TabIndex = 15;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            this.toolStripProgressBar1.Size = new System.Drawing.Size(450, 16);
            this.toolStripProgressBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label4.Location = new System.Drawing.Point(13, 18);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(85, 23);
            this.label4.TabIndex = 16;
            this.label4.Text = "帳票名　：";
            // 
            // cmbPrnName
            // 
            this.cmbPrnName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbPrnName.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cmbPrnName.FormattingEnabled = true;
            this.cmbPrnName.Items.AddRange(new object[] {
            "勤怠データＩ／Ｐ票",
            "応援移動票"});
            this.cmbPrnName.Location = new System.Drawing.Point(104, 11);
            this.cmbPrnName.Name = "cmbPrnName";
            this.cmbPrnName.Size = new System.Drawing.Size(355, 31);
            this.cmbPrnName.TabIndex = 0;
            this.cmbPrnName.SelectedIndexChanged += new System.EventHandler(this.cmbPrnName_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label5.Location = new System.Drawing.Point(13, 48);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(85, 23);
            this.label5.TabIndex = 18;
            this.label5.Text = "勤務体系：";
            // 
            // label6
            // 
            this.label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label6.Font = new System.Drawing.Font("Meiryo UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label6.Location = new System.Drawing.Point(104, 46);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(320, 22);
            this.label6.TabIndex = 19;
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Meiryo UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button1.Location = new System.Drawing.Point(425, 45);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(34, 25);
            this.button1.TabIndex = 1;
            this.button1.Text = "...";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // prePrint
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(476, 484);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.cmbPrnName);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.linkRtn);
            this.Controls.Add(this.linkPrn);
            this.Controls.Add(this.chkWhite);
            this.Controls.Add(this.linkLblOff);
            this.Controls.Add(this.linkLblOn);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dateTimePicker2);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Meiryo UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "prePrint";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "勤怠データＩ／Ｐ票発行";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.prePrint_FormClosing);
            this.Load += new System.EventHandler(this.prePrint_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.DateTimePicker dateTimePicker2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.LinkLabel linkLblOn;
        private System.Windows.Forms.LinkLabel linkLblOff;
        private System.Windows.Forms.CheckBox chkWhite;
        private System.Windows.Forms.LinkLabel linkPrn;
        private System.Windows.Forms.LinkLabel linkRtn;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cmbPrnName;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}