﻿namespace SZDS_TIMECARD.config
{
    partial class frmCalendar
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmCalendar));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.monthCalendar1 = new System.Windows.Forms.MonthCalendar();
            this.txtDate = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.lnkLblUpdate = new System.Windows.Forms.LinkLabel();
            this.lnkLblDelete = new System.Windows.Forms.LinkLabel();
            this.lnkLblClr = new System.Windows.Forms.LinkLabel();
            this.linkLabel4 = new System.Windows.Forms.LinkLabel();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            this.dataGridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(13, 50);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 21;
            this.dataGridView1.Size = new System.Drawing.Size(414, 322);
            this.dataGridView1.StandardTab = true;
            this.dataGridView1.TabIndex = 1;
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
            this.dataGridView1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);
            // 
            // monthCalendar1
            // 
            this.monthCalendar1.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.monthCalendar1.Location = new System.Drawing.Point(474, 50);
            this.monthCalendar1.Margin = new System.Windows.Forms.Padding(12, 15, 12, 15);
            this.monthCalendar1.Name = "monthCalendar1";
            this.monthCalendar1.TabIndex = 2;
            this.monthCalendar1.DateSelected += new System.Windows.Forms.DateRangeEventHandler(this.monthCalendar1_DateSelected);
            // 
            // txtDate
            // 
            this.txtDate.BackColor = System.Drawing.SystemColors.Window;
            this.txtDate.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtDate.Location = new System.Drawing.Point(498, 232);
            this.txtDate.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtDate.Name = "txtDate";
            this.txtDate.ReadOnly = true;
            this.txtDate.Size = new System.Drawing.Size(175, 30);
            this.txtDate.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(450, 237);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(48, 20);
            this.label1.TabIndex = 7;
            this.label1.Text = "日付：";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(451, 275);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(48, 20);
            this.label2.TabIndex = 8;
            this.label2.Text = "名称：";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(498, 270);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(175, 28);
            this.comboBox1.TabIndex = 4;
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Location = new System.Drawing.Point(74, 14);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(73, 28);
            this.comboBox2.TabIndex = 0;
            this.comboBox2.SelectedIndexChanged += new System.EventHandler(this.comboBox2_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 17);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(61, 20);
            this.label3.TabIndex = 15;
            this.label3.Text = "対象年：";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(497, 323);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(176, 49);
            this.button1.TabIndex = 5;
            this.button1.Text = "土・日一括登録";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // lnkLblUpdate
            // 
            this.lnkLblUpdate.Image = ((System.Drawing.Image)(resources.GetObject("lnkLblUpdate.Image")));
            this.lnkLblUpdate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lnkLblUpdate.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.lnkLblUpdate.Location = new System.Drawing.Point(340, 397);
            this.lnkLblUpdate.Name = "lnkLblUpdate";
            this.lnkLblUpdate.Size = new System.Drawing.Size(67, 37);
            this.lnkLblUpdate.TabIndex = 6;
            this.lnkLblUpdate.TabStop = true;
            this.lnkLblUpdate.Text = "登録";
            this.lnkLblUpdate.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lnkLblUpdate.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // lnkLblDelete
            // 
            this.lnkLblDelete.Image = ((System.Drawing.Image)(resources.GetObject("lnkLblDelete.Image")));
            this.lnkLblDelete.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lnkLblDelete.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.lnkLblDelete.Location = new System.Drawing.Point(428, 397);
            this.lnkLblDelete.Name = "lnkLblDelete";
            this.lnkLblDelete.Size = new System.Drawing.Size(71, 37);
            this.lnkLblDelete.TabIndex = 7;
            this.lnkLblDelete.TabStop = true;
            this.lnkLblDelete.Text = "削除";
            this.lnkLblDelete.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lnkLblDelete.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel2_LinkClicked);
            // 
            // lnkLblClr
            // 
            this.lnkLblClr.Image = ((System.Drawing.Image)(resources.GetObject("lnkLblClr.Image")));
            this.lnkLblClr.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lnkLblClr.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.lnkLblClr.Location = new System.Drawing.Point(520, 397);
            this.lnkLblClr.Name = "lnkLblClr";
            this.lnkLblClr.Size = new System.Drawing.Size(71, 37);
            this.lnkLblClr.TabIndex = 8;
            this.lnkLblClr.TabStop = true;
            this.lnkLblClr.Text = "取消";
            this.lnkLblClr.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lnkLblClr.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel3_LinkClicked);
            // 
            // linkLabel4
            // 
            this.linkLabel4.Image = ((System.Drawing.Image)(resources.GetObject("linkLabel4.Image")));
            this.linkLabel4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.linkLabel4.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.linkLabel4.Location = new System.Drawing.Point(607, 397);
            this.linkLabel4.Name = "linkLabel4";
            this.linkLabel4.Size = new System.Drawing.Size(71, 37);
            this.linkLabel4.TabIndex = 9;
            this.linkLabel4.TabStop = true;
            this.linkLabel4.Text = "終了";
            this.linkLabel4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.linkLabel4.Click += new System.EventHandler(this.linkLabel4_Click);
            // 
            // frmCalendar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(690, 443);
            this.Controls.Add(this.linkLabel4);
            this.Controls.Add(this.lnkLblClr);
            this.Controls.Add(this.lnkLblDelete);
            this.Controls.Add(this.lnkLblUpdate);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.comboBox2);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtDate);
            this.Controls.Add(this.monthCalendar1);
            this.Controls.Add(this.dataGridView1);
            this.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.MaximizeBox = false;
            this.Name = "frmCalendar";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "休日登録";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmCalendar_FormClosing);
            this.Load += new System.EventHandler(this.frmCalendar_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.MonthCalendar monthCalendar1;
        private System.Windows.Forms.TextBox txtDate;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.LinkLabel lnkLblUpdate;
        private System.Windows.Forms.LinkLabel lnkLblDelete;
        private System.Windows.Forms.LinkLabel lnkLblClr;
        private System.Windows.Forms.LinkLabel linkLabel4;
    }
}