namespace SZDS_TIMECARD.OCR
{
    [System.ComponentModel.ToolboxItem(true)]
    partial class Template3
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
            GrapeCity.Win.MultiRow.CellStyle cellStyle1 = new GrapeCity.Win.MultiRow.CellStyle();
            GrapeCity.Win.MultiRow.RoundedBorder roundedBorder1 = new GrapeCity.Win.MultiRow.RoundedBorder();
            GrapeCity.Win.MultiRow.CellStyle cellStyle2 = new GrapeCity.Win.MultiRow.CellStyle();
            GrapeCity.Win.MultiRow.Border border1 = new GrapeCity.Win.MultiRow.Border();
            GrapeCity.Win.MultiRow.CellStyle cellStyle3 = new GrapeCity.Win.MultiRow.CellStyle();
            GrapeCity.Win.MultiRow.Border border2 = new GrapeCity.Win.MultiRow.Border();
            GrapeCity.Win.MultiRow.CellStyle cellStyle4 = new GrapeCity.Win.MultiRow.CellStyle();
            GrapeCity.Win.MultiRow.Border border3 = new GrapeCity.Win.MultiRow.Border();
            GrapeCity.Win.MultiRow.CellStyle cellStyle5 = new GrapeCity.Win.MultiRow.CellStyle();
            GrapeCity.Win.MultiRow.Border border4 = new GrapeCity.Win.MultiRow.Border();
            GrapeCity.Win.MultiRow.CellStyle cellStyle6 = new GrapeCity.Win.MultiRow.CellStyle();
            GrapeCity.Win.MultiRow.Border border5 = new GrapeCity.Win.MultiRow.Border();
            GrapeCity.Win.MultiRow.CellStyle cellStyle7 = new GrapeCity.Win.MultiRow.CellStyle();
            GrapeCity.Win.MultiRow.Border border6 = new GrapeCity.Win.MultiRow.Border();
            GrapeCity.Win.MultiRow.CellStyle cellStyle8 = new GrapeCity.Win.MultiRow.CellStyle();
            GrapeCity.Win.MultiRow.RoundedBorder roundedBorder2 = new GrapeCity.Win.MultiRow.RoundedBorder();
            GrapeCity.Win.MultiRow.CellStyle cellStyle9 = new GrapeCity.Win.MultiRow.CellStyle();
            GrapeCity.Win.MultiRow.RoundedBorder roundedBorder3 = new GrapeCity.Win.MultiRow.RoundedBorder();
            GrapeCity.Win.MultiRow.CellStyle cellStyle10 = new GrapeCity.Win.MultiRow.CellStyle();
            GrapeCity.Win.MultiRow.Border border7 = new GrapeCity.Win.MultiRow.Border();
            GrapeCity.Win.MultiRow.CellStyle cellStyle11 = new GrapeCity.Win.MultiRow.CellStyle();
            GrapeCity.Win.MultiRow.RoundedBorder roundedBorder4 = new GrapeCity.Win.MultiRow.RoundedBorder();
            this.columnHeaderSection1 = new GrapeCity.Win.MultiRow.ColumnHeaderSection();
            this.txtYear = new GrapeCity.Win.MultiRow.TextBoxCell();
            this.txtMonth = new GrapeCity.Win.MultiRow.TextBoxCell();
            this.txtDay = new GrapeCity.Win.MultiRow.TextBoxCell();
            this.labelCell1 = new GrapeCity.Win.MultiRow.LabelCell();
            this.labelCell2 = new GrapeCity.Win.MultiRow.LabelCell();
            this.labelCell3 = new GrapeCity.Win.MultiRow.LabelCell();
            this.lblWeek = new GrapeCity.Win.MultiRow.LabelCell();
            this.lblShozoku = new GrapeCity.Win.MultiRow.LabelCell();
            this.labelCell7 = new GrapeCity.Win.MultiRow.LabelCell();
            this.lblPage = new GrapeCity.Win.MultiRow.LabelCell();
            this.txtBushoCode = new GrapeCity.Win.MultiRow.TextBoxCell();
            // 
            // Row
            // 
            this.Row.BackColor = System.Drawing.SystemColors.Control;
            this.Row.Cells.Add(this.txtYear);
            this.Row.Cells.Add(this.txtMonth);
            this.Row.Cells.Add(this.txtDay);
            this.Row.Cells.Add(this.labelCell1);
            this.Row.Cells.Add(this.labelCell2);
            this.Row.Cells.Add(this.labelCell3);
            this.Row.Cells.Add(this.lblWeek);
            this.Row.Cells.Add(this.lblShozoku);
            this.Row.Cells.Add(this.labelCell7);
            this.Row.Cells.Add(this.lblPage);
            this.Row.Cells.Add(this.txtBushoCode);
            this.Row.Height = 51;
            // 
            // columnHeaderSection1
            // 
            this.columnHeaderSection1.Height = 1;
            this.columnHeaderSection1.Name = "columnHeaderSection1";
            // 
            // txtYear
            // 
            this.txtYear.Location = new System.Drawing.Point(3, 2);
            this.txtYear.MaxLength = 4;
            this.txtYear.Name = "txtYear";
            this.txtYear.Size = new System.Drawing.Size(50, 24);
            cellStyle1.BackColor = System.Drawing.SystemColors.Window;
            roundedBorder1.Bottom = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            roundedBorder1.Left = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            roundedBorder1.Top = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            roundedBorder1.TopLeftCornerLine = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            roundedBorder1.TopLeftCornerRadius = 0.14F;
            cellStyle1.Border = roundedBorder1;
            cellStyle1.Font = new System.Drawing.Font("ＭＳ ゴシック", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            cellStyle1.ForeColor = System.Drawing.Color.Navy;
            cellStyle1.TextAlign = GrapeCity.Win.MultiRow.MultiRowContentAlignment.MiddleRight;
            this.txtYear.Style = cellStyle1;
            this.txtYear.TabIndex = 0;
            // 
            // txtMonth
            // 
            this.txtMonth.Location = new System.Drawing.Point(72, 2);
            this.txtMonth.MaxLength = 2;
            this.txtMonth.Name = "txtMonth";
            this.txtMonth.Size = new System.Drawing.Size(22, 24);
            cellStyle2.BackColor = System.Drawing.SystemColors.Window;
            border1.Top = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            cellStyle2.Border = border1;
            cellStyle2.Font = new System.Drawing.Font("ＭＳ ゴシック", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            cellStyle2.ForeColor = System.Drawing.Color.Navy;
            cellStyle2.TextAlign = GrapeCity.Win.MultiRow.MultiRowContentAlignment.MiddleCenter;
            this.txtMonth.Style = cellStyle2;
            this.txtMonth.TabIndex = 2;
            // 
            // txtDay
            // 
            this.txtDay.Location = new System.Drawing.Point(112, 2);
            this.txtDay.MaxLength = 2;
            this.txtDay.Name = "txtDay";
            this.txtDay.Size = new System.Drawing.Size(27, 24);
            cellStyle3.BackColor = System.Drawing.SystemColors.Window;
            border2.Top = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            cellStyle3.Border = border2;
            cellStyle3.Font = new System.Drawing.Font("ＭＳ ゴシック", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            cellStyle3.ForeColor = System.Drawing.Color.Navy;
            cellStyle3.TextAlign = GrapeCity.Win.MultiRow.MultiRowContentAlignment.MiddleRight;
            this.txtDay.Style = cellStyle3;
            this.txtDay.TabIndex = 4;
            // 
            // labelCell1
            // 
            this.labelCell1.Location = new System.Drawing.Point(53, 2);
            this.labelCell1.Name = "labelCell1";
            this.labelCell1.Selectable = false;
            this.labelCell1.Size = new System.Drawing.Size(19, 24);
            cellStyle4.BackColor = System.Drawing.SystemColors.Window;
            border3.Bottom = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.Gray);
            border3.Top = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            cellStyle4.Border = border3;
            cellStyle4.Font = new System.Drawing.Font("メイリオ", 8.25F);
            cellStyle4.ImeMode = System.Windows.Forms.ImeMode.Off;
            cellStyle4.TextAlign = GrapeCity.Win.MultiRow.MultiRowContentAlignment.BottomLeft;
            this.labelCell1.Style = cellStyle4;
            this.labelCell1.TabIndex = 1;
            this.labelCell1.TabStop = false;
            this.labelCell1.Value = "年";
            // 
            // labelCell2
            // 
            this.labelCell2.Location = new System.Drawing.Point(94, 2);
            this.labelCell2.Name = "labelCell2";
            this.labelCell2.Selectable = false;
            this.labelCell2.Size = new System.Drawing.Size(18, 24);
            cellStyle5.BackColor = System.Drawing.SystemColors.Window;
            border4.Bottom = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            border4.Top = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            cellStyle5.Border = border4;
            cellStyle5.Font = new System.Drawing.Font("メイリオ", 8.25F);
            cellStyle5.ImeMode = System.Windows.Forms.ImeMode.Off;
            cellStyle5.TextAlign = GrapeCity.Win.MultiRow.MultiRowContentAlignment.BottomLeft;
            this.labelCell2.Style = cellStyle5;
            this.labelCell2.TabIndex = 3;
            this.labelCell2.TabStop = false;
            this.labelCell2.Value = "月";
            // 
            // labelCell3
            // 
            this.labelCell3.Location = new System.Drawing.Point(139, 2);
            this.labelCell3.Name = "labelCell3";
            this.labelCell3.Selectable = false;
            this.labelCell3.Size = new System.Drawing.Size(18, 24);
            cellStyle6.BackColor = System.Drawing.SystemColors.Window;
            border5.Bottom = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            border5.Top = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            cellStyle6.Border = border5;
            cellStyle6.Font = new System.Drawing.Font("メイリオ", 8.25F);
            cellStyle6.ImeMode = System.Windows.Forms.ImeMode.Off;
            cellStyle6.TextAlign = GrapeCity.Win.MultiRow.MultiRowContentAlignment.BottomLeft;
            this.labelCell3.Style = cellStyle6;
            this.labelCell3.TabIndex = 5;
            this.labelCell3.TabStop = false;
            this.labelCell3.Value = "日";
            // 
            // lblWeek
            // 
            this.lblWeek.Location = new System.Drawing.Point(157, 2);
            this.lblWeek.Name = "lblWeek";
            this.lblWeek.Selectable = false;
            this.lblWeek.Size = new System.Drawing.Size(31, 24);
            cellStyle7.BackColor = System.Drawing.SystemColors.Window;
            border6.Top = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            cellStyle7.Border = border6;
            cellStyle7.Font = new System.Drawing.Font("ＭＳ ゴシック", 12F);
            cellStyle7.ImeMode = System.Windows.Forms.ImeMode.Off;
            cellStyle7.TextAlign = GrapeCity.Win.MultiRow.MultiRowContentAlignment.MiddleCenter;
            this.lblWeek.Style = cellStyle7;
            this.lblWeek.TabIndex = 6;
            this.lblWeek.TabStop = false;
            // 
            // lblShozoku
            // 
            this.lblShozoku.Location = new System.Drawing.Point(64, 26);
            this.lblShozoku.Name = "lblShozoku";
            this.lblShozoku.Selectable = false;
            this.lblShozoku.Size = new System.Drawing.Size(534, 24);
            roundedBorder2.Bottom = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.Gray);
            roundedBorder2.BottomRightCornerLine = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.Gray);
            roundedBorder2.BottomRightCornerRadius = 0.14F;
            roundedBorder2.Left = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.Gray);
            roundedBorder2.Right = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.Gray);
            roundedBorder2.Top = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.Gray);
            roundedBorder2.TopRightCornerLine = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.Gray);
            roundedBorder2.TopRightCornerRadius = 0.14F;
            cellStyle8.Border = roundedBorder2;
            cellStyle8.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Bold);
            cellStyle8.ForeColor = System.Drawing.Color.Blue;
            cellStyle8.ImeMode = System.Windows.Forms.ImeMode.Off;
            cellStyle8.TextAlign = GrapeCity.Win.MultiRow.MultiRowContentAlignment.MiddleCenter;
            this.lblShozoku.Style = cellStyle8;
            this.lblShozoku.TabIndex = 9;
            this.lblShozoku.TabStop = false;
            // 
            // labelCell7
            // 
            this.labelCell7.Location = new System.Drawing.Point(188, 2);
            this.labelCell7.Name = "labelCell7";
            this.labelCell7.Selectable = false;
            this.labelCell7.Size = new System.Drawing.Size(28, 24);
            cellStyle9.BackColor = System.Drawing.SystemColors.Window;
            roundedBorder3.Bottom = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            roundedBorder3.Right = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            roundedBorder3.Top = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            roundedBorder3.TopRightCornerLine = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            roundedBorder3.TopRightCornerRadius = 0.14F;
            cellStyle9.Border = roundedBorder3;
            cellStyle9.Font = new System.Drawing.Font("メイリオ", 8.25F);
            cellStyle9.ImeMode = System.Windows.Forms.ImeMode.Off;
            cellStyle9.TextAlign = GrapeCity.Win.MultiRow.MultiRowContentAlignment.BottomLeft;
            this.labelCell7.Style = cellStyle9;
            this.labelCell7.TabIndex = 7;
            this.labelCell7.TabStop = false;
            this.labelCell7.Value = "曜";
            // 
            // lblPage
            // 
            this.lblPage.Location = new System.Drawing.Point(512, 2);
            this.lblPage.Name = "lblPage";
            this.lblPage.Selectable = false;
            cellStyle10.Border = border7;
            cellStyle10.Font = new System.Drawing.Font("ＭＳ ゴシック", 9.75F);
            cellStyle10.ImeMode = System.Windows.Forms.ImeMode.Off;
            cellStyle10.TextAlign = GrapeCity.Win.MultiRow.MultiRowContentAlignment.MiddleRight;
            this.lblPage.Style = cellStyle10;
            this.lblPage.TabIndex = 10;
            this.lblPage.TabStop = false;
            // 
            // txtBushoCode
            // 
            this.txtBushoCode.Location = new System.Drawing.Point(3, 26);
            this.txtBushoCode.MaxLength = 5;
            this.txtBushoCode.Name = "txtBushoCode";
            this.txtBushoCode.Size = new System.Drawing.Size(61, 24);
            cellStyle11.BackColor = System.Drawing.SystemColors.Window;
            roundedBorder4.Bottom = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            roundedBorder4.BottomLeftCornerLine = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            roundedBorder4.BottomLeftCornerRadius = 0.14F;
            roundedBorder4.Left = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            roundedBorder4.Right = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Thin, System.Drawing.Color.DimGray);
            roundedBorder4.Top = new GrapeCity.Win.MultiRow.Line(GrapeCity.Win.MultiRow.LineStyle.Medium, System.Drawing.Color.DimGray);
            cellStyle11.Border = roundedBorder4;
            cellStyle11.Font = new System.Drawing.Font("ＭＳ ゴシック", 12F);
            cellStyle11.ForeColor = System.Drawing.Color.Navy;
            cellStyle11.TextAlign = GrapeCity.Win.MultiRow.MultiRowContentAlignment.MiddleCenter;
            this.txtBushoCode.Style = cellStyle11;
            this.txtBushoCode.TabIndex = 8;
            // 
            // Template3
            // 
            this.ColumnHeaders.AddRange(new GrapeCity.Win.MultiRow.ColumnHeaderSection[] {
            this.columnHeaderSection1});
            this.Width = 598;

        }

        #endregion

        private GrapeCity.Win.MultiRow.ColumnHeaderSection columnHeaderSection1;
        private GrapeCity.Win.MultiRow.TextBoxCell txtYear;
        private GrapeCity.Win.MultiRow.TextBoxCell txtMonth;
        private GrapeCity.Win.MultiRow.TextBoxCell txtDay;
        private GrapeCity.Win.MultiRow.LabelCell labelCell1;
        private GrapeCity.Win.MultiRow.LabelCell labelCell2;
        private GrapeCity.Win.MultiRow.LabelCell labelCell3;
        private GrapeCity.Win.MultiRow.TextBoxCell txtBushoCode;
        private GrapeCity.Win.MultiRow.LabelCell lblWeek;
        private GrapeCity.Win.MultiRow.LabelCell lblShozoku;
        private GrapeCity.Win.MultiRow.LabelCell labelCell7;
        private GrapeCity.Win.MultiRow.LabelCell lblPage;
    }
}
