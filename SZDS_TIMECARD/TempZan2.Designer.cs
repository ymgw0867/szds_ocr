namespace SZDS_TIMECARD
{
    [System.ComponentModel.ToolboxItem(true)]
    partial class TempZan2
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
            GrapeCity.Win.MultiRow.CellStyle cellStyle5 = new GrapeCity.Win.MultiRow.CellStyle();
            GrapeCity.Win.MultiRow.Border border5 = new GrapeCity.Win.MultiRow.Border();
            GrapeCity.Win.MultiRow.CellStyle cellStyle6 = new GrapeCity.Win.MultiRow.CellStyle();
            GrapeCity.Win.MultiRow.Border border6 = new GrapeCity.Win.MultiRow.Border();
            GrapeCity.Win.MultiRow.CellStyle cellStyle7 = new GrapeCity.Win.MultiRow.CellStyle();
            GrapeCity.Win.MultiRow.Border border7 = new GrapeCity.Win.MultiRow.Border();
            GrapeCity.Win.MultiRow.CellStyle cellStyle8 = new GrapeCity.Win.MultiRow.CellStyle();
            GrapeCity.Win.MultiRow.Border border8 = new GrapeCity.Win.MultiRow.Border();
            GrapeCity.Win.MultiRow.CellStyle cellStyle1 = new GrapeCity.Win.MultiRow.CellStyle();
            GrapeCity.Win.MultiRow.Border border1 = new GrapeCity.Win.MultiRow.Border();
            GrapeCity.Win.MultiRow.CellStyle cellStyle2 = new GrapeCity.Win.MultiRow.CellStyle();
            GrapeCity.Win.MultiRow.Border border2 = new GrapeCity.Win.MultiRow.Border();
            GrapeCity.Win.MultiRow.CellStyle cellStyle3 = new GrapeCity.Win.MultiRow.CellStyle();
            GrapeCity.Win.MultiRow.Border border3 = new GrapeCity.Win.MultiRow.Border();
            GrapeCity.Win.MultiRow.CellStyle cellStyle4 = new GrapeCity.Win.MultiRow.CellStyle();
            GrapeCity.Win.MultiRow.Border border4 = new GrapeCity.Win.MultiRow.Border();
            this.columnHeaderSection1 = new GrapeCity.Win.MultiRow.ColumnHeaderSection();
            this.labelCell1 = new GrapeCity.Win.MultiRow.LabelCell();
            this.labelCell2 = new GrapeCity.Win.MultiRow.LabelCell();
            this.labelCell3 = new GrapeCity.Win.MultiRow.LabelCell();
            this.labelCell4 = new GrapeCity.Win.MultiRow.LabelCell();
            this.lblTittle = new GrapeCity.Win.MultiRow.LabelCell();
            this.lblZengetsu = new GrapeCity.Win.MultiRow.LabelCell();
            this.lblHikaku = new GrapeCity.Win.MultiRow.LabelCell();
            this.lblTougetsu = new GrapeCity.Win.MultiRow.LabelCell();
            // 
            // Row
            // 
            this.Row.Cells.Add(this.lblTittle);
            this.Row.Cells.Add(this.lblZengetsu);
            this.Row.Cells.Add(this.lblHikaku);
            this.Row.Cells.Add(this.lblTougetsu);
            this.Row.Height = 19;
            // 
            // columnHeaderSection1
            // 
            this.columnHeaderSection1.Cells.Add(this.labelCell1);
            this.columnHeaderSection1.Cells.Add(this.labelCell2);
            this.columnHeaderSection1.Cells.Add(this.labelCell3);
            this.columnHeaderSection1.Cells.Add(this.labelCell4);
            this.columnHeaderSection1.Height = 20;
            this.columnHeaderSection1.Name = "columnHeaderSection1";
            // 
            // labelCell1
            // 
            this.labelCell1.Location = new System.Drawing.Point(2, 2);
            this.labelCell1.Name = "labelCell1";
            this.labelCell1.Size = new System.Drawing.Size(131, 19);
            border5.Bottom = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            border5.Left = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            border5.Top = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            cellStyle5.Border = border5;
            cellStyle5.ImeMode = System.Windows.Forms.ImeMode.Off;
            cellStyle5.TextAlign = GrapeCity.Win.MultiRow.MultiRowContentAlignment.MiddleCenter;
            this.labelCell1.Style = cellStyle5;
            this.labelCell1.TabIndex = 0;
            this.labelCell1.Value = "前月比較";
            // 
            // labelCell2
            // 
            this.labelCell2.Location = new System.Drawing.Point(209, 2);
            this.labelCell2.Name = "labelCell2";
            this.labelCell2.Size = new System.Drawing.Size(76, 19);
            border6.Bottom = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            border6.Left = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            border6.Right = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            border6.Top = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            cellStyle6.Border = border6;
            cellStyle6.ImeMode = System.Windows.Forms.ImeMode.Off;
            cellStyle6.TextAlign = GrapeCity.Win.MultiRow.MultiRowContentAlignment.MiddleCenter;
            this.labelCell2.Style = cellStyle6;
            this.labelCell2.TabIndex = 1;
            this.labelCell2.Value = "当月";
            // 
            // labelCell3
            // 
            this.labelCell3.Location = new System.Drawing.Point(285, 2);
            this.labelCell3.Name = "labelCell3";
            this.labelCell3.Size = new System.Drawing.Size(79, 19);
            border7.Bottom = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            border7.Right = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            border7.Top = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            cellStyle7.Border = border7;
            cellStyle7.ImeMode = System.Windows.Forms.ImeMode.Off;
            cellStyle7.TextAlign = GrapeCity.Win.MultiRow.MultiRowContentAlignment.MiddleCenter;
            this.labelCell3.Style = cellStyle7;
            this.labelCell3.TabIndex = 2;
            this.labelCell3.Value = "比較";
            // 
            // labelCell4
            // 
            this.labelCell4.Location = new System.Drawing.Point(133, 2);
            this.labelCell4.Name = "labelCell4";
            this.labelCell4.Size = new System.Drawing.Size(76, 19);
            border8.Bottom = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            border8.Left = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            border8.Right = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            border8.Top = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            cellStyle8.Border = border8;
            cellStyle8.ImeMode = System.Windows.Forms.ImeMode.Off;
            cellStyle8.TextAlign = GrapeCity.Win.MultiRow.MultiRowContentAlignment.MiddleCenter;
            this.labelCell4.Style = cellStyle8;
            this.labelCell4.TabIndex = 3;
            this.labelCell4.Value = "前月";
            // 
            // lblTittle
            // 
            this.lblTittle.Location = new System.Drawing.Point(2, 1);
            this.lblTittle.Name = "lblTittle";
            this.lblTittle.Size = new System.Drawing.Size(131, 19);
            cellStyle1.BackColor = System.Drawing.SystemColors.Window;
            border1.Bottom = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            border1.Left = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            border1.Right = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            border1.Top = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            cellStyle1.Border = border1;
            cellStyle1.ImeMode = System.Windows.Forms.ImeMode.Off;
            cellStyle1.TextAlign = GrapeCity.Win.MultiRow.MultiRowContentAlignment.MiddleCenter;
            this.lblTittle.Style = cellStyle1;
            this.lblTittle.TabIndex = 0;
            // 
            // lblZengetsu
            // 
            this.lblZengetsu.Location = new System.Drawing.Point(133, 1);
            this.lblZengetsu.Name = "lblZengetsu";
            this.lblZengetsu.Size = new System.Drawing.Size(76, 19);
            cellStyle2.BackColor = System.Drawing.SystemColors.Window;
            border2.Bottom = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            border2.Left = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            border2.Right = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            border2.Top = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            cellStyle2.Border = border2;
            cellStyle2.ImeMode = System.Windows.Forms.ImeMode.Off;
            cellStyle2.TextAlign = GrapeCity.Win.MultiRow.MultiRowContentAlignment.MiddleCenter;
            this.lblZengetsu.Style = cellStyle2;
            this.lblZengetsu.TabIndex = 1;
            // 
            // lblHikaku
            // 
            this.lblHikaku.Location = new System.Drawing.Point(285, 1);
            this.lblHikaku.Name = "lblHikaku";
            this.lblHikaku.Size = new System.Drawing.Size(79, 19);
            cellStyle3.BackColor = System.Drawing.SystemColors.Window;
            border3.Bottom = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            border3.Left = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            border3.Right = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            border3.Top = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            cellStyle3.Border = border3;
            cellStyle3.ImeMode = System.Windows.Forms.ImeMode.Off;
            cellStyle3.TextAlign = GrapeCity.Win.MultiRow.MultiRowContentAlignment.MiddleCenter;
            this.lblHikaku.Style = cellStyle3;
            this.lblHikaku.TabIndex = 2;
            // 
            // lblTougetsu
            // 
            this.lblTougetsu.Location = new System.Drawing.Point(209, 1);
            this.lblTougetsu.Name = "lblTougetsu";
            this.lblTougetsu.Size = new System.Drawing.Size(76, 19);
            cellStyle4.BackColor = System.Drawing.SystemColors.Window;
            border4.Bottom = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            border4.Left = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            border4.Right = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            border4.Top = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            cellStyle4.Border = border4;
            cellStyle4.ImeMode = System.Windows.Forms.ImeMode.Off;
            cellStyle4.TextAlign = GrapeCity.Win.MultiRow.MultiRowContentAlignment.MiddleCenter;
            this.lblTougetsu.Style = cellStyle4;
            this.lblTougetsu.TabIndex = 3;
            // 
            // TempZan2
            // 
            this.ColumnHeaders.AddRange(new GrapeCity.Win.MultiRow.ColumnHeaderSection[] {
            this.columnHeaderSection1});
            this.Width = 364;

        }

        #endregion

        private GrapeCity.Win.MultiRow.ColumnHeaderSection columnHeaderSection1;
        private GrapeCity.Win.MultiRow.LabelCell labelCell1;
        private GrapeCity.Win.MultiRow.LabelCell labelCell2;
        private GrapeCity.Win.MultiRow.LabelCell labelCell3;
        private GrapeCity.Win.MultiRow.LabelCell lblTittle;
        private GrapeCity.Win.MultiRow.LabelCell lblZengetsu;
        private GrapeCity.Win.MultiRow.LabelCell lblHikaku;
        private GrapeCity.Win.MultiRow.LabelCell lblTougetsu;
        private GrapeCity.Win.MultiRow.LabelCell labelCell4;
    }
}
