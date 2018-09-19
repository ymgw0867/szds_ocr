namespace SZDS_TIMECARD.OCR
{
    partial class frmOuenCorrect
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmOuenCorrect));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.leadImg = new Leadtools.WinForms.RasterImageViewer();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.hScrollBar1 = new System.Windows.Forms.HScrollBar();
            this.btnEnd = new System.Windows.Forms.Button();
            this.btnNext = new System.Windows.Forms.Button();
            this.btnBefore = new System.Windows.Forms.Button();
            this.btnFirst = new System.Windows.Forms.Button();
            this.btnPlus = new System.Windows.Forms.Button();
            this.btnMinus = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblErrMsg = new System.Windows.Forms.Label();
            this.lblNoImage = new System.Windows.Forms.Label();
            this.lnkIP = new System.Windows.Forms.LinkLabel();
            this.gcMultiRow3 = new GrapeCity.Win.MultiRow.GcMultiRow();
            this.template52 = new SZDS_TIMECARD.OCR.Template5();
            this.gcMultiRow2 = new GrapeCity.Win.MultiRow.GcMultiRow();
            this.template42 = new SZDS_TIMECARD.OCR.Template4();
            this.gcMultiRow1 = new GrapeCity.Win.MultiRow.GcMultiRow();
            this.template32 = new SZDS_TIMECARD.OCR.Template3();
            this.template62 = new SZDS_TIMECARD.OCR.Template6();
            this.button1 = new System.Windows.Forms.Button();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.lnkDel = new System.Windows.Forms.LinkLabel();
            this.lnkRtn = new System.Windows.Forms.LinkLabel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lnkErrCheck = new System.Windows.Forms.LinkLabel();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("メイリオ", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.Location = new System.Drawing.Point(563, 73);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(259, 23);
            this.label1.TabIndex = 3;
            this.label1.Text = "【日中（シフト勤務内）応援記入欄】";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("メイリオ", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label2.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label2.Location = new System.Drawing.Point(563, 317);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(259, 23);
            this.label2.TabIndex = 4;
            this.label2.Text = "【残業（シフト勤務外）応援記入欄】";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // leadImg
            // 
            this.leadImg.Location = new System.Drawing.Point(3, 3);
            this.leadImg.Name = "leadImg";
            this.leadImg.Size = new System.Drawing.Size(564, 580);
            this.leadImg.TabIndex = 123;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(3, 2);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(564, 581);
            this.pictureBox1.TabIndex = 122;
            this.pictureBox1.TabStop = false;
            // 
            // hScrollBar1
            // 
            this.hScrollBar1.Location = new System.Drawing.Point(229, 586);
            this.hScrollBar1.Name = "hScrollBar1";
            this.hScrollBar1.Size = new System.Drawing.Size(298, 34);
            this.hScrollBar1.TabIndex = 11;
            this.hScrollBar1.Scroll += new System.Windows.Forms.ScrollEventHandler(this.hScrollBar1_Scroll);
            // 
            // btnEnd
            // 
            this.btnEnd.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnEnd.Image = ((System.Drawing.Image)(resources.GetObject("btnEnd.Image")));
            this.btnEnd.Location = new System.Drawing.Point(189, 586);
            this.btnEnd.Name = "btnEnd";
            this.btnEnd.Size = new System.Drawing.Size(37, 34);
            this.btnEnd.TabIndex = 10;
            this.btnEnd.TabStop = false;
            this.btnEnd.UseVisualStyleBackColor = true;
            this.btnEnd.Click += new System.EventHandler(this.btnEnd_Click);
            // 
            // btnNext
            // 
            this.btnNext.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnNext.Image = ((System.Drawing.Image)(resources.GetObject("btnNext.Image")));
            this.btnNext.Location = new System.Drawing.Point(152, 586);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(37, 34);
            this.btnNext.TabIndex = 9;
            this.btnNext.TabStop = false;
            this.btnNext.UseVisualStyleBackColor = true;
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // btnBefore
            // 
            this.btnBefore.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnBefore.Image = ((System.Drawing.Image)(resources.GetObject("btnBefore.Image")));
            this.btnBefore.Location = new System.Drawing.Point(115, 586);
            this.btnBefore.Name = "btnBefore";
            this.btnBefore.Size = new System.Drawing.Size(37, 34);
            this.btnBefore.TabIndex = 8;
            this.btnBefore.TabStop = false;
            this.btnBefore.UseVisualStyleBackColor = true;
            this.btnBefore.Click += new System.EventHandler(this.btnBefore_Click);
            // 
            // btnFirst
            // 
            this.btnFirst.Font = new System.Drawing.Font("Meiryo UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnFirst.Image = ((System.Drawing.Image)(resources.GetObject("btnFirst.Image")));
            this.btnFirst.Location = new System.Drawing.Point(78, 586);
            this.btnFirst.Name = "btnFirst";
            this.btnFirst.Size = new System.Drawing.Size(37, 34);
            this.btnFirst.TabIndex = 7;
            this.btnFirst.TabStop = false;
            this.btnFirst.UseVisualStyleBackColor = true;
            this.btnFirst.Click += new System.EventHandler(this.btnFirst_Click);
            // 
            // btnPlus
            // 
            this.btnPlus.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnPlus.Image = ((System.Drawing.Image)(resources.GetObject("btnPlus.Image")));
            this.btnPlus.Location = new System.Drawing.Point(4, 586);
            this.btnPlus.Name = "btnPlus";
            this.btnPlus.Size = new System.Drawing.Size(37, 34);
            this.btnPlus.TabIndex = 5;
            this.btnPlus.TabStop = false;
            this.btnPlus.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnPlus.UseVisualStyleBackColor = true;
            this.btnPlus.Click += new System.EventHandler(this.btnPlus_Click);
            // 
            // btnMinus
            // 
            this.btnMinus.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnMinus.Image = ((System.Drawing.Image)(resources.GetObject("btnMinus.Image")));
            this.btnMinus.Location = new System.Drawing.Point(41, 586);
            this.btnMinus.Name = "btnMinus";
            this.btnMinus.Size = new System.Drawing.Size(37, 34);
            this.btnMinus.TabIndex = 6;
            this.btnMinus.TabStop = false;
            this.btnMinus.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnMinus.UseVisualStyleBackColor = true;
            this.btnMinus.Click += new System.EventHandler(this.btnMinus_Click);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.lblErrMsg);
            this.panel1.Location = new System.Drawing.Point(573, 527);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(598, 56);
            this.panel1.TabIndex = 163;
            // 
            // lblErrMsg
            // 
            this.lblErrMsg.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblErrMsg.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblErrMsg.ForeColor = System.Drawing.Color.Red;
            this.lblErrMsg.Location = new System.Drawing.Point(0, 0);
            this.lblErrMsg.Name = "lblErrMsg";
            this.lblErrMsg.Size = new System.Drawing.Size(594, 52);
            this.lblErrMsg.TabIndex = 0;
            this.lblErrMsg.Text = "label33";
            this.lblErrMsg.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblNoImage
            // 
            this.lblNoImage.Font = new System.Drawing.Font("メイリオ", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblNoImage.ForeColor = System.Drawing.SystemColors.AppWorkspace;
            this.lblNoImage.Location = new System.Drawing.Point(125, 300);
            this.lblNoImage.Name = "lblNoImage";
            this.lblNoImage.Size = new System.Drawing.Size(322, 42);
            this.lblNoImage.TabIndex = 172;
            this.lblNoImage.Text = "画像はありません";
            this.lblNoImage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lnkIP
            // 
            this.lnkIP.AutoSize = true;
            this.lnkIP.Font = new System.Drawing.Font("Meiryo UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lnkIP.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.lnkIP.LinkColor = System.Drawing.Color.Blue;
            this.lnkIP.Location = new System.Drawing.Point(1, 621);
            this.lnkIP.Name = "lnkIP";
            this.lnkIP.Size = new System.Drawing.Size(185, 17);
            this.lnkIP.TabIndex = 13;
            this.lnkIP.TabStop = true;
            this.lnkIP.Text = "勤怠データＩ／Ｐ票データ画面へ";
            this.lnkIP.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lnkIP.Visible = false;
            this.lnkIP.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // gcMultiRow3
            // 
            this.gcMultiRow3.AllowUserToAddRows = false;
            this.gcMultiRow3.AllowUserToDeleteRows = false;
            this.gcMultiRow3.AllowUserToResize = false;
            this.gcMultiRow3.AllowUserToZoom = false;
            this.gcMultiRow3.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.gcMultiRow3.Location = new System.Drawing.Point(573, 338);
            this.gcMultiRow3.Name = "gcMultiRow3";
            this.gcMultiRow3.ScrollBarMode = GrapeCity.Win.MultiRow.ScrollBarMode.Automatic;
            this.gcMultiRow3.Size = new System.Drawing.Size(598, 183);
            this.gcMultiRow3.TabIndex = 2;
            this.gcMultiRow3.Template = this.template52;
            this.gcMultiRow3.Text = "gcMultiRow3";
            this.gcMultiRow3.CellValueChanged += new System.EventHandler<GrapeCity.Win.MultiRow.CellEventArgs>(this.gcMultiRow3_CellValueChanged);
            this.gcMultiRow3.CellEnter += new System.EventHandler<GrapeCity.Win.MultiRow.CellEventArgs>(this.gcMultiRow3_CellEnter);
            this.gcMultiRow3.CellLeave += new System.EventHandler<GrapeCity.Win.MultiRow.CellEventArgs>(this.gcMultiRow3_CellLeave);
            this.gcMultiRow3.EditingControlShowing += new System.EventHandler<GrapeCity.Win.MultiRow.EditingControlShowingEventArgs>(this.gcMultiRow3_EditingControlShowing);
            this.gcMultiRow3.CurrentCellDirtyStateChanged += new System.EventHandler(this.gcMultiRow3_CurrentCellDirtyStateChanged);
            // 
            // gcMultiRow2
            // 
            this.gcMultiRow2.AllowUserToAddRows = false;
            this.gcMultiRow2.AllowUserToDeleteRows = false;
            this.gcMultiRow2.AllowUserToResize = false;
            this.gcMultiRow2.AllowUserToZoom = false;
            this.gcMultiRow2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.gcMultiRow2.EditMode = GrapeCity.Win.MultiRow.EditMode.EditProgrammatically;
            this.gcMultiRow2.Location = new System.Drawing.Point(573, 94);
            this.gcMultiRow2.Name = "gcMultiRow2";
            this.gcMultiRow2.ScrollBarMode = GrapeCity.Win.MultiRow.ScrollBarMode.Automatic;
            this.gcMultiRow2.Size = new System.Drawing.Size(598, 183);
            this.gcMultiRow2.TabIndex = 1;
            this.gcMultiRow2.Template = this.template42;
            this.gcMultiRow2.Text = "gcMultiRow2";
            this.gcMultiRow2.CellValueChanged += new System.EventHandler<GrapeCity.Win.MultiRow.CellEventArgs>(this.gcMultiRow2_CellValueChanged);
            this.gcMultiRow2.CellEnter += new System.EventHandler<GrapeCity.Win.MultiRow.CellEventArgs>(this.gcMultiRow2_CellEnter);
            this.gcMultiRow2.CellLeave += new System.EventHandler<GrapeCity.Win.MultiRow.CellEventArgs>(this.gcMultiRow2_CellLeave);
            this.gcMultiRow2.EditingControlShowing += new System.EventHandler<GrapeCity.Win.MultiRow.EditingControlShowingEventArgs>(this.gcMultiRow2_EditingControlShowing);
            this.gcMultiRow2.CurrentCellDirtyStateChanged += new System.EventHandler(this.gcMultiRow2_CurrentCellDirtyStateChanged);
            // 
            // gcMultiRow1
            // 
            this.gcMultiRow1.AllowUserToAddRows = false;
            this.gcMultiRow1.AllowUserToDeleteRows = false;
            this.gcMultiRow1.AllowUserToResize = false;
            this.gcMultiRow1.AllowUserToZoom = false;
            this.gcMultiRow1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.gcMultiRow1.EditMode = GrapeCity.Win.MultiRow.EditMode.EditProgrammatically;
            this.gcMultiRow1.Location = new System.Drawing.Point(573, 13);
            this.gcMultiRow1.Name = "gcMultiRow1";
            this.gcMultiRow1.ScrollBarMode = GrapeCity.Win.MultiRow.ScrollBarMode.Automatic;
            this.gcMultiRow1.Size = new System.Drawing.Size(598, 52);
            this.gcMultiRow1.TabIndex = 0;
            this.gcMultiRow1.Template = this.template32;
            this.gcMultiRow1.Text = "gcMultiRow1";
            this.gcMultiRow1.CellValueChanged += new System.EventHandler<GrapeCity.Win.MultiRow.CellEventArgs>(this.gcMultiRow1_CellValueChanged);
            this.gcMultiRow1.CellEnter += new System.EventHandler<GrapeCity.Win.MultiRow.CellEventArgs>(this.gcMultiRow1_CellEnter);
            this.gcMultiRow1.EditingControlShowing += new System.EventHandler<GrapeCity.Win.MultiRow.EditingControlShowingEventArgs>(this.gcMultiRow1_EditingControlShowing);
            // 
            // button1
            // 
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Location = new System.Drawing.Point(530, 589);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(37, 34);
            this.button1.TabIndex = 12;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Font = new System.Drawing.Font("メイリオ", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.checkBox1.Location = new System.Drawing.Point(730, 605);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(71, 25);
            this.checkBox1.TabIndex = 3;
            this.checkBox1.Text = "確認済";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // lnkDel
            // 
            this.lnkDel.Font = new System.Drawing.Font("Meiryo UI", 9.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lnkDel.Image = ((System.Drawing.Image)(resources.GetObject("lnkDel.Image")));
            this.lnkDel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lnkDel.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.lnkDel.Location = new System.Drawing.Point(127, 11);
            this.lnkDel.Name = "lnkDel";
            this.lnkDel.Size = new System.Drawing.Size(134, 37);
            this.lnkDel.TabIndex = 1;
            this.lnkDel.TabStop = true;
            this.lnkDel.Text = "応援移動票削除";
            this.lnkDel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lnkDel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkDel_LinkClicked);
            // 
            // lnkRtn
            // 
            this.lnkRtn.Font = new System.Drawing.Font("Meiryo UI", 9.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lnkRtn.Image = ((System.Drawing.Image)(resources.GetObject("lnkRtn.Image")));
            this.lnkRtn.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lnkRtn.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.lnkRtn.Location = new System.Drawing.Point(271, 11);
            this.lnkRtn.Name = "lnkRtn";
            this.lnkRtn.Size = new System.Drawing.Size(74, 37);
            this.lnkRtn.TabIndex = 2;
            this.lnkRtn.TabStop = true;
            this.lnkRtn.Text = "終了";
            this.lnkRtn.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lnkRtn.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkRtn_LinkClicked);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lnkRtn);
            this.groupBox1.Controls.Add(this.lnkDel);
            this.groupBox1.Controls.Add(this.lnkErrCheck);
            this.groupBox1.Location = new System.Drawing.Point(818, 585);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(351, 50);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            // 
            // lnkErrCheck
            // 
            this.lnkErrCheck.Font = new System.Drawing.Font("Meiryo UI", 9.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lnkErrCheck.Image = ((System.Drawing.Image)(resources.GetObject("lnkErrCheck.Image")));
            this.lnkErrCheck.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lnkErrCheck.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.lnkErrCheck.Location = new System.Drawing.Point(6, 11);
            this.lnkErrCheck.Name = "lnkErrCheck";
            this.lnkErrCheck.Size = new System.Drawing.Size(111, 37);
            this.lnkErrCheck.TabIndex = 0;
            this.lnkErrCheck.TabStop = true;
            this.lnkErrCheck.Text = "エラーチェック";
            this.lnkErrCheck.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lnkErrCheck.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkErrCheck_LinkClicked);
            // 
            // frmOuenCorrect
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1181, 642);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.lnkIP);
            this.Controls.Add(this.lblNoImage);
            this.Controls.Add(this.gcMultiRow3);
            this.Controls.Add(this.gcMultiRow2);
            this.Controls.Add(this.gcMultiRow1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.hScrollBar1);
            this.Controls.Add(this.btnEnd);
            this.Controls.Add(this.btnNext);
            this.Controls.Add(this.btnBefore);
            this.Controls.Add(this.btnFirst);
            this.Controls.Add(this.btnPlus);
            this.Controls.Add(this.btnMinus);
            this.Controls.Add(this.leadImg);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmOuenCorrect";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "応援移動票データ作成";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmOuenCorrect_FormClosing);
            this.Load += new System.EventHandler(this.frmOuenCorrectcs_Load);
            this.Shown += new System.EventHandler(this.frmOuenCorrect_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Template3 template31;
        private Template4 template41;
        private Template5 template51;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private Template6 template61;
        private Leadtools.WinForms.RasterImageViewer leadImg;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.HScrollBar hScrollBar1;
        private System.Windows.Forms.Button btnEnd;
        private System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.Button btnBefore;
        private System.Windows.Forms.Button btnFirst;
        private System.Windows.Forms.Button btnPlus;
        private System.Windows.Forms.Button btnMinus;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label lblErrMsg;
        private GrapeCity.Win.MultiRow.GcMultiRow gcMultiRow1;
        private GrapeCity.Win.MultiRow.GcMultiRow gcMultiRow2;
        private GrapeCity.Win.MultiRow.GcMultiRow gcMultiRow3;
        private Template3 template32;
        private Template4 template42;
        private Template5 template52;
        private Template6 template62;
        private System.Windows.Forms.Label lblNoImage;
        private System.Windows.Forms.LinkLabel lnkIP;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.LinkLabel lnkDel;
        private System.Windows.Forms.LinkLabel lnkRtn;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.LinkLabel lnkErrCheck;
    }
}