namespace SZDS_TIMECARD.config
{
    partial class frmKitakugoWork
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmKitakugoWork));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.linkLabel4 = new System.Windows.Forms.LinkLabel();
            this.lnkLblClr = new System.Windows.Forms.LinkLabel();
            this.lnkLblDelete = new System.Windows.Forms.LinkLabel();
            this.lnkLblUpdate = new System.Windows.Forms.LinkLabel();
            this.gcMultiRow1 = new GrapeCity.Win.MultiRow.GcMultiRow();
            this.tmpKitakugo1 = new SZDS_TIMECARD.config.tmpKitakugo();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 12);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 21;
            this.dataGridView1.Size = new System.Drawing.Size(489, 180);
            this.dataGridView1.StandardTab = true;
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.TabStop = false;
            this.dataGridView1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);
            // 
            // linkLabel4
            // 
            this.linkLabel4.Image = ((System.Drawing.Image)(resources.GetObject("linkLabel4.Image")));
            this.linkLabel4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.linkLabel4.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.linkLabel4.Location = new System.Drawing.Point(430, 483);
            this.linkLabel4.Name = "linkLabel4";
            this.linkLabel4.Size = new System.Drawing.Size(71, 37);
            this.linkLabel4.TabIndex = 5;
            this.linkLabel4.TabStop = true;
            this.linkLabel4.Text = "終了";
            this.linkLabel4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.linkLabel4.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel4_LinkClicked);
            // 
            // lnkLblClr
            // 
            this.lnkLblClr.Image = ((System.Drawing.Image)(resources.GetObject("lnkLblClr.Image")));
            this.lnkLblClr.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lnkLblClr.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.lnkLblClr.Location = new System.Drawing.Point(343, 483);
            this.lnkLblClr.Name = "lnkLblClr";
            this.lnkLblClr.Size = new System.Drawing.Size(71, 37);
            this.lnkLblClr.TabIndex = 4;
            this.lnkLblClr.TabStop = true;
            this.lnkLblClr.Text = "取消";
            this.lnkLblClr.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lnkLblClr.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkLblClr_LinkClicked);
            // 
            // lnkLblDelete
            // 
            this.lnkLblDelete.Image = ((System.Drawing.Image)(resources.GetObject("lnkLblDelete.Image")));
            this.lnkLblDelete.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lnkLblDelete.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.lnkLblDelete.Location = new System.Drawing.Point(251, 483);
            this.lnkLblDelete.Name = "lnkLblDelete";
            this.lnkLblDelete.Size = new System.Drawing.Size(71, 37);
            this.lnkLblDelete.TabIndex = 3;
            this.lnkLblDelete.TabStop = true;
            this.lnkLblDelete.Text = "削除";
            this.lnkLblDelete.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lnkLblDelete.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkLblDelete_LinkClicked);
            // 
            // lnkLblUpdate
            // 
            this.lnkLblUpdate.Image = ((System.Drawing.Image)(resources.GetObject("lnkLblUpdate.Image")));
            this.lnkLblUpdate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lnkLblUpdate.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.lnkLblUpdate.Location = new System.Drawing.Point(163, 483);
            this.lnkLblUpdate.Name = "lnkLblUpdate";
            this.lnkLblUpdate.Size = new System.Drawing.Size(67, 37);
            this.lnkLblUpdate.TabIndex = 2;
            this.lnkLblUpdate.TabStop = true;
            this.lnkLblUpdate.Text = "登録";
            this.lnkLblUpdate.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lnkLblUpdate.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkLblUpdate_LinkClicked);
            // 
            // gcMultiRow1
            // 
            this.gcMultiRow1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.gcMultiRow1.Location = new System.Drawing.Point(12, 206);
            this.gcMultiRow1.Name = "gcMultiRow1";
            this.gcMultiRow1.ScrollBarMode = GrapeCity.Win.MultiRow.ScrollBarMode.Automatic;
            this.gcMultiRow1.ScrollBars = System.Windows.Forms.ScrollBars.None;
            this.gcMultiRow1.Size = new System.Drawing.Size(489, 268);
            this.gcMultiRow1.TabIndex = 1;
            this.gcMultiRow1.Template = this.tmpKitakugo1;
            this.gcMultiRow1.Text = "gcMultiRow1";
            this.gcMultiRow1.CellValueChanged += new System.EventHandler<GrapeCity.Win.MultiRow.CellEventArgs>(this.gcMultiRow1_CellValueChanged);
            this.gcMultiRow1.CellEnter += new System.EventHandler<GrapeCity.Win.MultiRow.CellEventArgs>(this.gcMultiRow1_CellEnter);
            this.gcMultiRow1.EditingControlShowing += new System.EventHandler<GrapeCity.Win.MultiRow.EditingControlShowingEventArgs>(this.gcMultiRow1_EditingControlShowing);
            this.gcMultiRow1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.gcMultiRow1_KeyPress);
            // 
            // frmKitakugoWork
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(514, 524);
            this.Controls.Add(this.linkLabel4);
            this.Controls.Add(this.lnkLblClr);
            this.Controls.Add(this.lnkLblDelete);
            this.Controls.Add(this.lnkLblUpdate);
            this.Controls.Add(this.gcMultiRow1);
            this.Controls.Add(this.dataGridView1);
            this.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "frmKitakugoWork";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "帰宅後勤務登録";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmKitakugoWork_FormClosing);
            this.Load += new System.EventHandler(this.frmKitakugoWork_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private GrapeCity.Win.MultiRow.GcMultiRow gcMultiRow1;
        private tmpKitakugo tmpKitakugo1;
        private System.Windows.Forms.LinkLabel linkLabel4;
        private System.Windows.Forms.LinkLabel lnkLblClr;
        private System.Windows.Forms.LinkLabel lnkLblDelete;
        private System.Windows.Forms.LinkLabel lnkLblUpdate;
    }
}