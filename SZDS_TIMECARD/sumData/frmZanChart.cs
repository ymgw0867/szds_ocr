using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using SZDS_TIMECARD.Common;

namespace SZDS_TIMECARD.sumData
{
    public partial class frmZanChart : Form
    {
        public frmZanChart(string dbName)
        {
            InitializeComponent();

            hAdp.Fill(dts.過去勤務票ヘッダ);
            dAdp.Fill(dts.休日);
        }

        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.残業集計TableAdapter adp = new DataSet1TableAdapters.残業集計TableAdapter();
        DataSet1TableAdapters.過去勤務票ヘッダTableAdapter hAdp = new DataSet1TableAdapters.過去勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.休日TableAdapter dAdp = new DataSet1TableAdapters.休日TableAdapter();

        xlsData bs;
        string _dbName = string.Empty;

        string[] colChartData = null;

        clsZanSum[] z = null;
        DateTime dtEnd = DateTime.Today;

        private void frmZanChart_Load(object sender, EventArgs e)
        {
            bs = new xlsData();
            bs.zpArray = bs.getZanPlan();
        }

        private void setColChartData(ref string [] c, string ymd, int bushoCode)
        {
            DateTime dt = DateTime.Today;

            if (!DateTime.TryParse(ymd, out dt))
            {
                return;
            }

            DateTime dt01 = DateTime.Parse(dt.Year.ToString() + "/" + dt.Month.ToString() + "/01");
            dtEnd = dt01.AddMonths(1).AddDays(-1); // 当月末日

            DateTime dtTo = dt01;
            int iX = 0;

            z = new clsZanSum[dt.Day];

            while (dtTo <= dt)
            {
                //if (!dts.休日.Any(a => a.年月日 == dtTo))
                //{
                //    Array.Resize(ref c, iX + 1);

                //    if (dtTo <= dt)
                //    {
                //        // 出勤簿データ日付範囲内のとき初期値「０」
                //        c[iX] = dtTo.Day.ToString() + ",0,0";
                //    }
                //    else
                //    {
                //        // 出勤簿データ日付範囲外のとき初期値「値なし」
                //        c[iX] = dtTo.Day.ToString() + ",,0";
                //    }

                //    iX++;

                z[iX] = new clsZanSum();

                z[iX].sSzCode = bushoCode.ToString();
                z[iX].sDay = dtTo.Day;
                z[iX].sZangyo = 0;
                z[iX].sMonthPlan = 0;
                z[iX].sPlanbyDay = 0;
                z[iX].sZissekibyDay = "";
                z[iX].sYear = dt.Year;
                z[iX].sMonth = dt.Month;
                z[iX].sEndDay = dtEnd.Day;

                if (!dts.休日.Any(a => a.年月日 == dtTo))
                {
                    z[iX].sZangyo = 0;
                    z[iX].sMonthPlan = 0;
                    z[iX].sPlanbyDay = 0;
                    z[iX].sZissekibyDay = "";
                    z[iX].sHoliday = 0;
                }
                else
                {
                    z[iX].sHoliday = 1;
                }

                dtTo = dtTo.AddDays(1);

                iX++;
            }
        }


        private void setChart1()
        {
            chart1.BackColor = Color.LightGray;

            // チャートエリアのインスタンス
            ChartArea area1 = new ChartArea();
            area1.Name = "main";

            // チャートエリアを追加
            chart1.ChartAreas.Add(area1);

            // 軸ラベルの設定
            //chart1.ChartAreas["area1"].AxisX.Title = "日";
            //chart1.ChartAreas["area1"].AxisY.Title = "残業";
            //※Axis.TitleFontでフォントも指定できるがこれはデザイナで変更したほうが楽

            // X軸最小値、最大値、目盛間隔の設定
            area1.AxisX.Minimum = 0;
            area1.AxisX.Maximum = dtEnd.Day;
            area1.AxisX.Interval = 1;

            //chart1.ChartAreas["area1"].AxisX.Minimum = 0;
            //chart1.ChartAreas["area1"].AxisX.Maximum = dtEnd.Day;
            //chart1.ChartAreas["area1"].AxisX.Interval = 2;

            ////int iX = 0;

            ////var s = z.Where(a => a.sHoliday == global.flgOff)
            ////    .Select (a => new {
            ////        day = a.sDay}).OrderBy(a => a.day);


            ////for (int i = 0; i < z.Length; i++)
            ////{
            ////    if (z[i].sHoliday == global.flgOff)
            ////    {
            ////        chart1.ChartAreas["area1"].AxisX.CustomLabels.Add(new CustomLabel(iX + 2, 10, z[i].sDay.ToString(), 0,  LabelMarkStyle.None));
            ////        //chart1.ChartAreas["area1"].AxisX.CustomLabels[iX].Text =  z[i].sDay.ToString();
            ////        iX++;
            ////    }
            ////}

            // Y軸最小値、最大値、目盛間隔の設定
            area1.AxisY.Minimum = 0;
            area1.AxisY.Maximum = 300;
            area1.AxisY.Interval = 50;

            // Y軸右側 最小値、最大値、目盛間隔の設定
            area1.AxisY2.Enabled = AxisEnabled.True;    // 第2Y軸を有効にする
            area1.AxisY2.Minimum = 0;
            area1.AxisY2.Maximum = 40;
            area1.AxisY2.Interval = 5;

            // 目盛線の消去
            area1.AxisX.MajorGrid.Enabled = false;
            area1.AxisY.MajorGrid.Enabled = false;
            area1.AxisY2.MajorGrid.Enabled = false;

            // 日別残業実績グラフ
            Series byDay = new Series("日々の残業時間");         // インスタンス
            byDay.ChartArea = area1.Name;                       // チャートエリアを指定する
            byDay.ChartType = SeriesChartType.Column;           // グラフの種類を棒グラフに設定
            byDay.XAxisType = AxisType.Primary;                 // 下側のＸ軸の目盛を使用する
            byDay.YAxisType = AxisType.Primary;                 // 左側のＹ軸の目盛を使用する

            // 目標累積時間グラフ
            Series byKeikakuSum = new Series("目標累積時間"); // インスタンス
            byKeikakuSum.ChartArea = area1.Name;            // チャートエリアを指定する
            byKeikakuSum.ChartType = SeriesChartType.Line;  // グラフの種類を折れ線グラフにする
            byKeikakuSum.XAxisType = AxisType.Primary;      // 下側のＸ軸の目盛を使用する
            byKeikakuSum.YAxisType = AxisType.Primary;      // 左側のＹ軸の目盛を使用する
            byKeikakuSum.MarkerStyle = MarkerStyle.Diamond; // マーカー

            // 実績累積時間グラフ
            Series byZissekiSum = new Series("実績累積時間"); // インスタンス
            byZissekiSum.ChartArea = area1.Name;            // チャートエリアを指定する
            byZissekiSum.ChartType = SeriesChartType.Line;  // グラフの種類を折れ線グラフにする
            byZissekiSum.XAxisType = AxisType.Primary;      // 下側のＸ軸の目盛を使用する
            byZissekiSum.YAxisType = AxisType.Primary;      // 左側のＹ軸の目盛を使用する
            byZissekiSum.MarkerStyle = MarkerStyle.Square;  // マーカー

            // 日々残業目標時間グラフ
            Series byKeikakuAve = new Series("日々の目標時間");    // インスタンス
            byKeikakuAve.ChartArea = area1.Name;                // チャートエリアを指定する
            byKeikakuAve.ChartType = SeriesChartType.Line;      // グラフの種類を折れ線グラフにする
            byKeikakuAve.XAxisType = AxisType.Primary;          // 下側のＸ軸の目盛を使用する
            byKeikakuAve.YAxisType = AxisType.Secondary;        // 右側のＹ軸の目盛を使用する

            // グラフを追加する（後の方が上になる）
            chart1.Series.Add(byKeikakuAve);
            chart1.Series.Add(byZissekiSum);
            chart1.Series.Add(byKeikakuSum);
            chart1.Series.Add(byDay);

            // データをセットする
            for (int i = 0; i < z.Length; i++)
            {
                if (z[i].sHoliday == global.flgOff)
                {
                    // 日別残業グラフのデータを追加
                    byDay.Points.AddXY(z[i].sDay, z[i].sZangyo);

                    // 当月計画日割りグラフのデータを追加
                    byKeikakuSum.Points.AddXY(z[i].sDay, z[i].sPlanbyDay);

                    // 当月実績日割りグラフのデータを追加
                    byZissekiSum.Points.AddXY(z[i].sDay, z[i].sZissekibyDay);

                    // 日々計画グラフのデータを追加
                    byKeikakuAve.Points.AddXY(z[i].sDay, z[i].sMonthPlan);
                }
            }
        }
        
        private void yymmChanged()
        {
            DateTime dt;
            string str = txtYear.Text + "/" + txtMonth.Text + "/1";
            if (!DateTime.TryParse(str, out dt))
            {
                return;
            }

            if (dts.過去勤務票ヘッダ.Any(a => a.年 == Utility.StrtoInt(txtYear.Text) &&
                                                   a.月 == Utility.StrtoInt(txtMonth.Text)))
            {
                // 最新の出勤簿日付を取得・表示
                var s = dts.過去勤務票ヘッダ.Where(a => a.年 == Utility.StrtoInt(txtYear.Text) &&
                                                       a.月 == Utility.StrtoInt(txtMonth.Text))
                                           .Max(a => a.日);

                lblKDays.Text = getKadouDays(Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text)).ToString();
                lblDate.Text = txtYear.Text + "/" + txtMonth.Text.PadLeft(2, '0') + "/" + s.ToString().PadLeft(2, '0');
                lblWdays.Text = getWorkDays(DateTime.Parse(lblDate.Text)).ToString();

                linkLabel1.Enabled = true;
            }
            else
            {
                lblKDays.Text = "--";
                lblDate.Text = "出勤簿なし";
                lblWdays.Text = "--";

                linkLabel1.Enabled = false;
            }
        }

        private void txtYear_TextChanged(object sender, EventArgs e)
        {
            yymmChanged();
        }

        private void txtMonth_TextChanged(object sender, EventArgs e)
        {
            yymmChanged();
        }

        private int getWorkDays(DateTime dt)
        {
            int rtn = 0;

            //　該当月の該当日までの休日を取得
            int s = dts.休日.Count(a => a.年月日.Year == dt.Year && a.年月日.Month == dt.Month && a.年月日 <= dt);

            // 実働日数
            rtn = dt.Day - s;

            return rtn;
        }

        private int getKadouDays(int yy, int mm)
        {
            int rtn = 0;

            //　該当月の該当日までの休日を取得
            int s = dts.休日.Count(a => a.年月日.Year == yy && a.年月日.Month == mm);

            DateTime dt = new DateTime(yy, mm, 1);
            dt = dt.AddMonths(1);
            dt = dt.AddDays(-1);

            // 稼働日数
            rtn = dt.Day - s;

            return rtn;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (Utility.StrtoInt(txtYear.Text) < 2017)
            {
                MessageBox.Show("対象年が正しくありません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtYear.Focus();
                return;
            }

            if (Utility.StrtoInt(txtMonth.Text) < 1 || Utility.StrtoInt(txtMonth.Text) > 12)
            {
                MessageBox.Show("対象月が正しくありません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonth.Focus();
                return;
            }

            this.Cursor = Cursors.WaitCursor;
            setColChartData(ref colChartData, lblDate.Text, 11111);
            showZangyoTotal(Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text), "11111");
            setChart1();
            this.Cursor = Cursors.Default;
        }

        private void showZangyoTotal(int yy, int mm, string bushoCode)
        {
            // 残業集計データ読みこみ
            adp.Fill(dts.残業集計, yy, mm, yy, mm, yy, mm, yy, mm, yy, mm, yy, mm);

            // nullに「０」をセット
            foreach (var item in dts.残業集計)
            {
                if (item.Is残業時Null())
                {
                    item.残業時 = 0;
                }

                if (item.Is残業分Null())
                {
                    item.残業分 = 0;
                }
            }

            // 部署指定理由別で残業時間を集計
            var s = dts.残業集計.Where(a => a.部署コード == bushoCode).GroupBy(a => a.残業理由)
                .Select(g => new
                {
                    zanRe = g.Key,
                    zanH = g.Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10))
                })
                .OrderBy(a => a.zanRe);

            foreach (var t in s)
            {
                double zanZisseki = 0;　// 実績時間
                double kaDays = 0;      // 当月稼働日数
                int r = 0;

                //for (int rI = 0; rI < gr.RowCount; rI++)
                //{
                //    if (gr[ColSz, rI].Value.ToString() == t.buCode)
                //    {
                //        double zan = i.zanH / 60;

                //        if (i.zanRe == 1) gr[Col1, rI].Value = zan;
                //        if (i.zanRe == 2) gr[Col2, rI].Value = zan;
                //        if (i.zanRe == 3) gr[Col3, rI].Value = zan;
                //        if (i.zanRe == 4) gr[Col4, rI].Value = zan;
                //        if (i.zanRe == 5) gr[Col5, rI].Value = zan;
                //        if (i.zanRe == 6) gr[Col6, rI].Value = zan;
                //        if (i.zanRe == 7) gr[Col7, rI].Value = zan;
                //        if (i.zanRe == 8) gr[Col8, rI].Value = zan;
                //        if (i.zanRe == 9) gr[Col9, rI].Value = zan;
                //        if (i.zanRe >= 10) gr[Col10, rI].Value = zan;
                //        zanZisseki += zan;
                //        r = rI;
                //    }
                //
            }
            
            double zanTotal = 0;

            // 部署指定日別で残業時間を集計
            var d = dts.残業集計.Where(a => a.部署コード == bushoCode).GroupBy(a => a.日)
                .Select(g => new
                {
                    day = g.Key,
                    zanH = g.Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10))
                })
                .OrderBy(a => a.day);

            // 日別の残業時間を配列にセット
            foreach (var t in d)
            {
                //for (int i = 0; i < colChartData.Length ; i++)
                //{
                    //string[] arr = colChartData[i].Split(',');

                    //if (arr[0] == t.day.ToString())
                    //{
                    //    arr[1] = (t.zanH / 60).ToString();
                    //    colChartData[i] = arr[0] + "," + arr[1] + ",0";
                    //    break;
                    //}
                //}

                // 月間残業合計
                zanTotal += t.zanH;

                for (int i = 0; i < z.Length; i++)
                {
                    if (z[i].sDay == t.day)
                    {
                        z[i].sZangyo = Utility.StrtoDouble(((t.zanH / 60).ToString("#,##0.0")));
                        break;
                    }
                }
            }

            // 月間残業合計を時間単位に変換
            zanTotal = Utility.StrtoDouble(((zanTotal / 60).ToString("#,##0.0")));

            double zanPlan = 0;

            for (int i = 1; i < bs.zpArray.GetLength(0); i++)
            {
                // 対象年月以外のときは対象外
                if (Utility.StrtoInt(bs.zpArray[i, 2].ToString()) != yy || Utility.StrtoInt(bs.zpArray[i, 3].ToString()) != mm)
                {
                    continue;
                }

                // 該当部署の当月計画値を取得する
                if (bs.zpArray[i, 1].ToString() == bushoCode)
                {
                    zanPlan = Utility.StrtoDouble(Utility.NulltoStr(bs.zpArray[i, 5]));
                    break;
                }                
            }

            for (int i = 0; i < z.Length; i++)
            {
                // 当月計画値の稼働日数割りをセット
                z[i].sPlanbyDay = Utility.StrtoDouble((zanPlan / Utility.StrtoDouble(lblKDays.Text) * (i + 1)).ToString("#,##0.0"));

                // 当月実績値の稼働日数割りをセット
                z[i].sZissekibyDay = (Utility.StrtoDouble((zanTotal / Utility.StrtoDouble(lblKDays.Text) * (i + 1)).ToString("#,##0.0"))).ToString();

                // 当月計画値の日々目標時間をセット
                z[i].sMonthPlan = Utility.StrtoDouble((zanTotal / Utility.StrtoDouble(lblKDays.Text)).ToString("#,##0.0"));
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Close();
        }

        private void frmZanChart_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
        }
    }
}
