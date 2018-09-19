namespace SZDS_TIMECARD.OCR
{
    [System.ComponentModel.ToolboxItem(true)]
    partial class Template6
    {
        /// <summary> 
        /// 必要なデザイナ変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースが破棄される場合 true、破棄されない場合は false です。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region MultiRow Template Designer generated code

        /// <summary> 
        /// デザイナ サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディタで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            GrapeCity.Win.MultiRow.CellStyle cellStyle3 = new GrapeCity.Win.MultiRow.CellStyle();
            GrapeCity.Win.MultiRow.RoundedBorder roundedBorder1 = new GrapeCity.Win.MultiRow.RoundedBorder();
            GrapeCity.Win.MultiRow.CellStyle cellStyle1 = new GrapeCity.Win.MultiRow.CellStyle();
            GrapeCity.Win.MultiRow.Border border1 = new GrapeCity.Win.MultiRow.Border();
            GrapeCity.Win.MultiRow.CellStyle cellStyle2 = new GrapeCity.Win.MultiRow.CellStyle();
            GrapeCity.Win.MultiRow.Border border2 = new GrapeCity.Win.MultiRow.Border();
            this.columnHeaderSection1 = new GrapeCity.Win.MultiRow.ColumnHeaderSection();
            this.labelCell1 = new GrapeCity.Win.MultiRow.LabelCell();
            this.labelCell2 = new GrapeCity.Win.MultiRow.LabelCell();
            this.labelCell3 = new GrapeCity.Win.MultiRow.LabelCell();
            // 
            // Row
            // 
            this.Row.Cells.Add(this.labelCell2);
            this.Row.Cells.Add(this.labelCell3);
            this.Row.Height = 21;
            // 
            // columnHeaderSection1
            // 
            this.columnHeaderSection1.Cells.Add(this.labelCell1);
            this.columnHeaderSection1.Height = 23;
            this.columnHeaderSection1.Name = "columnHeaderSection1";
            // 
            // labelCell1
            // 
            this.labelCell1.Location = new System.Drawing.Point(2, 2);
            this.labelCell1.Name = "labelCell1";
            this.labelCell1.Size = new System.Drawing.Size(191, 21);
            roundedBorder1.Bottom = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            roundedBorder1.Left = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            roundedBorder1.Right = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            roundedBorder1.Top = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            roundedBorder1.TopLeftCornerLine = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            roundedBorder1.TopLeftCornerRadius = 0.14F;
            roundedBorder1.TopRightCornerLine = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            roundedBorder1.TopRightCornerRadius = 0.14F;
            cellStyle3.Border = roundedBorder1;
            cellStyle3.ImeMode = System.Windows.Forms.ImeMode.Off;
            cellStyle3.TextAlign = GrapeCity.Win.MultiRow.MultiRowContentAlignment.MiddleCenter;
            this.labelCell1.Style = cellStyle3;
            this.labelCell1.TabIndex = 0;
            this.labelCell1.Value = "残業理由コード一覧";
            // 
            // labelCell2
            // 
            this.labelCell2.Location = new System.Drawing.Point(2, 0);
            this.labelCell2.Name = "labelCell2";
            this.labelCell2.Size = new System.Drawing.Size(161, 21);
            border1.Bottom = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            border1.Left = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            cellStyle1.Border = border1;
            cellStyle1.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            cellStyle1.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.labelCell2.Style = cellStyle1;
            this.labelCell2.TabIndex = 0;
            // 
            // labelCell3
            // 
            this.labelCell3.Location = new System.Drawing.Point(163, 0);
            this.labelCell3.Name = "labelCell3";
            this.labelCell3.Size = new System.Drawing.Size(30, 21);
            border2.Bottom = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            border2.Left = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            border2.Right = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            cellStyle2.Border = border2;
            cellStyle2.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            cellStyle2.ImeMode = System.Windows.Forms.ImeMode.Off;
            cellStyle2.TextAlign = GrapeCity.Win.MultiRow.MultiRowContentAlignment.MiddleCenter;
            this.labelCell3.Style = cellStyle2;
            this.labelCell3.TabIndex = 1;
            // 
            // Template6
            // 
            this.ColumnHeaders.AddRange(new GrapeCity.Win.MultiRow.ColumnHeaderSection[] {
            this.columnHeaderSection1});
            this.Width = 194;

        }

        #endregion

        private GrapeCity.Win.MultiRow.ColumnHeaderSection columnHeaderSection1;
        private GrapeCity.Win.MultiRow.LabelCell labelCell1;
        private GrapeCity.Win.MultiRow.LabelCell labelCell2;
        private GrapeCity.Win.MultiRow.LabelCell labelCell3;
    }
}
