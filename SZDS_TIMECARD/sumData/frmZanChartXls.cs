using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections;
using SZDS_TIMECARD.Common;
using Excel = Microsoft.Office.Interop.Excel;

namespace SZDS_TIMECARD.sumData
{
    public partial class frmZanChartXls : Form
    {
        public frmZanChartXls(string dbName)
        {
            InitializeComponent();

            hAdp.Fill(dts.過去勤務票ヘッダ);
            dAdp.Fill(dts.休日);

            _dbName = dbName;
        }

        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.残業集計TableAdapter adp = new DataSet1TableAdapters.残業集計TableAdapter();
        DataSet1TableAdapters.過去勤務票ヘッダTableAdapter hAdp = new DataSet1TableAdapters.過去勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.休日TableAdapter dAdp = new DataSet1TableAdapters.休日TableAdapter();

        xlsData bs = new xlsData();
        string _dbName = string.Empty;

        int sNin = 0;                       // 人数
        int sSeisan = 0;                    // 生産数
        int zNin = 0;                       // 前月人数
        int zSeisan = 0;                    // 前月生産数
        double zenZan = 0;                  // 前月残業合計
        int zenKaDays = 0;                  // 前月稼働日数       

        private void button1_Click(object sender, EventArgs e)
        {
        }

        private void prtReport()
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

            if (!dts.過去勤務票ヘッダ.Any(a => a.年 == Utility.StrtoInt(txtYear.Text) && a.月 == Utility.StrtoInt(txtMonth.Text)))
            {
                MessageBox.Show("対象データがありません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonth.Focus();
                return;
            }

            if (MessageBox.Show("残業推移グラフを発行します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }
            
            this.Cursor = Cursors.WaitCursor;

            // 前月
            int zYY = 0;
            int zMM = 0;

            if (Utility.StrtoInt(txtMonth.Text) == 1)
            {
                zMM = 12;
                zYY = Utility.StrtoInt(txtYear.Text) - 1;
            }
            else
            {
                zMM = Utility.StrtoInt(txtMonth.Text) - 1;
                zYY = Utility.StrtoInt(txtYear.Text);
            }

            if (rBtn1.Checked)
            {
                // 部署別集計処理
                zanSum(Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text), zYY, zMM);
            }
            else if (rBtn2.Checked)
            {
                // 部門別集計処理
                zanSumBumon(Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text), zYY, zMM);
            }
            else if (rBtn3.Checked)
            {
                // 全社集計処理
                zanSumAll(Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text), zYY, zMM);
            }

            MessageBox.Show("処理が終了しました", "残業推移グラフ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            
            // カーソル戻す
            this.Cursor = Cursors.Default;
        }


        private void frmZanChartXls_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            // 部署名コンボボックスのデータソースをセットする
            Utility.ComboBumon.loadBusho(comboBox2, _dbName);

            label1.Visible = false;
            toolStripProgressBar1.Visible = false;

            txtYear.Text = DateTime.Today.Year.ToString();
            txtMonth.Text = DateTime.Today.Month.ToString();

            rBtn1.Checked = true;
            comboBox2.Enabled = true;

            comboBox1.SelectedIndex = 0;
        }

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     日付別の配列を生成する </summary>
        /// <param name="yy">
        ///     対象年</param>
        /// <param name="mm">
        ///     対象月</param>
        /// <param name="bushoCode">
        ///     部署または部門コード（0:全社、1:製造部門、2:間接部門）</param>
        /// <param name="z">
        ///     日付別配列</param>
        ///-----------------------------------------------------------------------
        private void dayArrayNew(int yy, int mm, string bushoCode, ref clsZanSum[] z)
        {            
            DateTime dt = DateTime.Today;
            DateTime dtEnd = DateTime.Today;

            DateTime dt01 = DateTime.Parse(yy + "/" + mm + "/01");
            dtEnd = dt01.AddMonths(1).AddDays(-1); // 対象月末日

            DateTime dtTo = dt01;
            int iX = 0;
            
            while (dtTo <= dtEnd)
            {
                if (iX > 0)
                {
                    Array.Resize(ref z, iX + 1);
                }

                z[iX] = new clsZanSum();

                z[iX].sSzCode = bushoCode;
                z[iX].sDay = dtTo.Day;
                z[iX].sZangyo = 0;
                z[iX].sMonthPlan = 0;
                z[iX].sPlanbyDay = 0;
                z[iX].sZissekibyDay = "";
                z[iX].sYear = dt01.Year;
                z[iX].sMonth = dt01.Month;
                z[iX].sEndDay = dtEnd.Day;

                if (!dts.休日.Any(a => a.年月日 == dtTo))
                {
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

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     全社メイン集計処理 </summary>
        /// <param name="yy">
        ///     対象年</param>
        /// <param name="mm">
        ///     対象月</param>
        /// <param name="zYY">
        ///     前月の年</param>
        /// <param name="zMM">
        ///     前月</param>
        ///-----------------------------------------------------------------------
        private void zanSumAll(int yy, int mm, int zYY, int zMM)
        {
            // 前月データから前月実績配列を作成
            string[,] zengetsuArray = null;
            setZengetsuZan(zYY, zMM, ref zengetsuArray, dts);

            // !!!!!!!!!!!! デバッグ用前月データがないので当月で動作確認。必ず戻すこと !!!!!!!!!!!!
            //setZengetsuZan(yy, mm, ref zengetsuArray, dts);
            // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

            // 当月データ取得
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

            // エクセルオブジェクト
            Excel.Application oXls = new Excel.Application();
            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(Properties.Settings.Default.xlsZanChart,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

            //// 奉行データベース接続
            //string sc = sqlControl.obcConnectSting.get(_dbName);
            //sqlControl.DataControl sdCon = new Common.sqlControl.DataControl(sc);

            try
            {
                int pCnt = 1;

                // 部門コード
                string bmnCode = string.Empty;

                // 部署別残業理由シートの内容を配列に取得する
                object[,] zReSeizou = bs.getZanReason();

                // 部署別残業理由別残業計画シートの内容を配列に取得する
                object[,] zRe = bs.getZanReasonPlan();

                // 部署別残業計画シートの内容を配列に取得する
                bs.zpArray = bs.getZanPlan();

                // progressBar
                int nMax = 1;
                toolStripProgressBar1.Maximum = nMax;
                toolStripProgressBar1.Minimum = 0;
                toolStripProgressBar1.Visible = true;

                // 一番最近の勤務票日付
                int maxDay = dts.過去勤務票ヘッダ.Where(a => a.年 == yy && a.月 == mm).Max(a => a.日);

                // progressBar表示
                label1.Visible = true;
                label1.Text = "残業推移グラフ作成中..." + "1/" + nMax;
                label1.Text = "残業推移グラフ作成中...";
                toolStripProgressBar1.Value = 1;
                this.Refresh(); // ← 追加
                
                bmnCode = global.FLGOFF;   // 部門コードを取得

                // 全社の人数、生産数取得
                sNin = 0;
                sSeisan = 0;
                getBumonNin(bmnCode, yy, mm, ref sNin, ref sSeisan);

                // 前月の残業合計、生産数、人数を取得
                zSeisan = 0;
                zNin = 0;
                zenZan = 0;
                zenKaDays = 0;
                zenSeisanNinBumon(zengetsuArray, bmnCode, zYY, zMM, ref zSeisan, ref zNin, ref zenZan, ref zenKaDays);

                // 該当部門の当月計画値を取得する
                double zanPlan = getZanPlanBumon(bs.zpArray, yy, mm, Utility.StrtoInt(bmnCode));

                // 日付別の配列を生成
                clsZanSum[] z = new clsZanSum[1];
                dayArrayNew(yy, mm, bmnCode, ref z);

                // 日別残業時間を配列にセットする・残業月間合計を取得する
                double zanTotal = 0;
                setDaybyZanBumon(bmnCode, ref z, ref zanTotal);

                // 残業計画時間の稼働日数割りと日々目標値を配列にセットする
                setDaybyPlan(ref z, zanPlan);

                // 残業時間の実績累積を配列にセットする
                setDaybyZisseki(ref z, yy, mm, maxDay);

                // 月間残業合計を時間単位に変換
                zanTotal = Utility.StrtoDouble(((zanTotal / 60).ToString("#,##0.0")));

                // エクセルシート出力
                // テンプレートシートを追加する
                pCnt++;
                oxlsSheet.Copy(Type.Missing, oXlsBook.Sheets[pCnt - 1]);
                oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt];

                // シートにデータを貼り付ける
                xlsOutPutBumon(oXls, ref oXlsBook, ref oxlsSheet, z, Properties.Settings.Default.xlsZanChart, yy, mm, zReSeizou, zRe, bmnCode);
             
                System.Threading.Thread.Sleep(1000);

                // 1枚目はテンプレートシートなので印刷時には削除する
                oXls.DisplayAlerts = false;
                oXlsBook.Sheets[1].Delete();

                // 1枚目のシートが表示されるようにする
                oxlsSheet = oXlsBook.Sheets[1];
                oxlsSheet.Select();

                //マウスポインタを元に戻す
                this.Cursor = Cursors.Default;

                // 確認のためExcelのウィンドウを表示する
                oXls.Visible = true;

                //印刷
                oXlsBook.PrintOutEx(Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //oXlsBook.PrintOutEx();
                //oXlsBook.PrintPreview(true);

                // ウィンドウを非表示にする
                oXls.Visible = false;

                //保存処理
                oXls.DisplayAlerts = false;

                DialogResult ret;

                //ダイアログボックスの初期設定
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Title = "残業推移グラフ";
                saveFileDialog1.OverwritePrompt = true;
                saveFileDialog1.RestoreDirectory = true;
                DateTime dt = DateTime.Now;
                saveFileDialog1.FileName = yy + "年" + mm.ToString().PadLeft(2, ' ') + "月 残業推移グラフ_全社";
                saveFileDialog1.Filter = "Microsoft Office Excelファイル(*.xlsx)|*.xlsx|全てのファイル(*.*)|*.*";

                //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
                string fileName;
                ret = saveFileDialog1.ShowDialog();

                if (ret == System.Windows.Forms.DialogResult.OK)
                {
                    fileName = saveFileDialog1.FileName;
                    oXlsBook.SaveAs(fileName, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing,
                                    Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }

                // プログレスバーを非表示
                toolStripProgressBar1.Visible = false;
                label1.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                //// 奉行データベース接続切断
                //if (sdCon.Cn.State == ConnectionState.Open)
                //{
                //    sdCon.Close();
                //}

                // Bookをクローズ
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                // Excelを終了
                oXls.Quit();

                // COM オブジェクトの参照カウントを解放する 
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                oXls = null;
                oXlsBook = null;
                oxlsSheet = null;

                GC.Collect();
            }
        }

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     部門別メイン集計処理 </summary>
        /// <param name="yy">
        ///     対象年</param>
        /// <param name="mm">
        ///     対象月</param>
        /// <param name="zYY">
        ///     前月の年</param>
        /// <param name="zMM">
        ///     前月</param>
        ///-----------------------------------------------------------------------
        private void zanSumBumon(int yy, int mm, int zYY, int zMM)
        {
            // 前月データから前月実績配列を作成
            string[,] zengetsuArray = null;
            setZengetsuZan(zYY, zMM, ref zengetsuArray, dts);

            //// !!!!!!!!!!!! デバッグ用前月データがないので当月で動作確認。必ず戻すこと !!!!!!!!!!!!
            //setZengetsuZan(yy, mm, ref zengetsuArray, dts);
            //// !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                        
            // 当月データ取得
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
            
            // エクセルオブジェクト
            Excel.Application oXls = new Excel.Application();
            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(Properties.Settings.Default.xlsZanChart, 
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

            //// 奉行データベース接続
            //string sc = sqlControl.obcConnectSting.get(_dbName);
            //sqlControl.DataControl sdCon = new Common.sqlControl.DataControl(sc);

            try
            {
                int pCnt = 1;

                // 部門コード
                string bmnCode = string.Empty;    

                // 部署別残業理由シートの内容を配列に取得する
                object[,] zReSeizou = bs.getZanReason();

                // 部署別残業理由別残業計画シートの内容を配列に取得する
                object[,] zRe = bs.getZanReasonPlan();

                // 部署別残業計画シートの内容を配列に取得する
                bs.zpArray = bs.getZanPlan();

                // progressBar
                int nMax = 2;
                toolStripProgressBar1.Maximum = nMax;
                toolStripProgressBar1.Minimum = 0;
                toolStripProgressBar1.Visible = true;

                // 一番最近の勤務票日付
                int maxDay = dts.過去勤務票ヘッダ.Where(a => a.年 == yy && a.月 == mm).Max(a => a.日);

                // 部署別残業計画配列を順番に読む
                for (int i = 1; i <= 2; i++)    // 製造部門1以下、間接部門2以上
                {
                    // progressBar表示
                    label1.Visible = true;
                    label1.Text = "残業推移グラフ作成中..." + i + "/" + nMax;
                    toolStripProgressBar1.Value = i;
                    this.Refresh(); // ← 追加

                    bmnCode = i.ToString();   // 部門コードを取得

                    // 部門の人数、生産数取得
                    sNin = 0;
                    sSeisan = 0;
                    getBumonNin(bmnCode, yy, mm, ref sNin, ref sSeisan);

                    // 前月の残業合計、生産数、人数を取得
                    zSeisan = 0;
                    zNin = 0;
                    zenZan = 0;
                    zenKaDays = 0;
                    zenSeisanNinBumon(zengetsuArray, bmnCode, zYY, zMM, ref zSeisan, ref zNin, ref zenZan, ref zenKaDays);

                    // 該当部門の当月計画値を取得する
                    double zanPlan = getZanPlanBumon(bs.zpArray, yy, mm, Utility.StrtoInt(bmnCode));

                    // 日付別の配列を生成
                    clsZanSum[] z = new clsZanSum[1];
                    dayArrayNew(yy, mm, bmnCode, ref z);

                    // 日別残業時間を配列にセットする・残業月間合計を取得する
                    double zanTotal = 0;
                    setDaybyZanBumon(bmnCode, ref z, ref zanTotal);

                    // 残業計画時間の稼働日数割りと日々目標値を配列にセットする
                    setDaybyPlan(ref z, zanPlan);

                    // 残業時間の実績累積を配列にセットする
                    setDaybyZisseki(ref z, yy, mm, maxDay);

                    // 月間残業合計を時間単位に変換
                    zanTotal = Utility.StrtoDouble(((zanTotal / 60).ToString("#,##0.0")));

                    // エクセルシート出力
                    // テンプレートシートを追加する
                    pCnt++;
                    oxlsSheet.Copy(Type.Missing, oXlsBook.Sheets[pCnt - 1]);
                    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt];

                    // シートにデータを貼り付ける
                    xlsOutPutBumon(oXls, ref oXlsBook, ref oxlsSheet, z, Properties.Settings.Default.xlsZanChart, yy, mm, zReSeizou, zRe, bmnCode);
                }

                System.Threading.Thread.Sleep(1000);

                // 1枚目はテンプレートシートなので印刷時には削除する
                oXls.DisplayAlerts = false;
                oXlsBook.Sheets[1].Delete();

                // 1枚目のシートが表示されるようにする
                oxlsSheet = oXlsBook.Sheets[1];
                oxlsSheet.Select();

                //マウスポインタを元に戻す
                this.Cursor = Cursors.Default;

                // 確認のためExcelのウィンドウを表示する
                oXls.Visible = true;

                //印刷
                oXlsBook.PrintOutEx(Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //oXlsBook.PrintOutEx();
                //oXlsBook.PrintPreview(true);

                // ウィンドウを非表示にする
                oXls.Visible = false;

                //保存処理
                oXls.DisplayAlerts = false;

                DialogResult ret;

                //ダイアログボックスの初期設定
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Title = "残業推移グラフ";
                saveFileDialog1.OverwritePrompt = true;
                saveFileDialog1.RestoreDirectory = true;
                DateTime dt = DateTime.Now;
                saveFileDialog1.FileName = yy + "年" + mm.ToString().PadLeft(2, ' ') + "月 残業推移グラフ_製造・間接";
                saveFileDialog1.Filter = "Microsoft Office Excelファイル(*.xlsx)|*.xlsx|全てのファイル(*.*)|*.*";

                //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
                string fileName;
                ret = saveFileDialog1.ShowDialog();

                if (ret == System.Windows.Forms.DialogResult.OK)
                {
                    fileName = saveFileDialog1.FileName;
                    oXlsBook.SaveAs(fileName, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing,
                                    Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }

                // プログレスバーを非表示
                toolStripProgressBar1.Visible = false;
                label1.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                //// 奉行データベース接続切断
                //if (sdCon.Cn.State == ConnectionState.Open)
                //{
                //    sdCon.Close();
                //}

                // Bookをクローズ
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                // Excelを終了
                oXls.Quit();

                // COM オブジェクトの参照カウントを解放する 
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                oXls = null;
                oXlsBook = null;
                oxlsSheet = null;

                GC.Collect();
            }
        }

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     部署別メイン集計処理 </summary>
        /// <param name="yy">
        ///     対象年</param>
        /// <param name="mm">
        ///     対象月</param>
        /// <param name="zYY">
        ///     前月の年</param>
        /// <param name="zMM">
        ///     前月</param>
        ///-----------------------------------------------------------------------
        private void zanSum(int yy, int mm, int zYY, int zMM)
        {
            // 前月データから前月実績配列を作成
            string[,] zengetsuArray = null;
            setZengetsuZan(zYY, zMM, ref zengetsuArray, dts);

            // !!!!!!!!!!!! デバッグ用前月データがないので当月で動作確認。必ず戻すこと !!!!!!!!!!!!
            //setZengetsuZan(yy, mm, ref zengetsuArray, dts);
            // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

            // 当月データ取得
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

            // エクセルオブジェクト
            Excel.Application oXls = new Excel.Application();
            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(Properties.Settings.Default.xlsZanChart,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

            // 奉行データベース接続
            string sc = sqlControl.obcConnectSting.get(_dbName);
            sqlControl.DataControl sdCon = new Common.sqlControl.DataControl(sc);

            try
            {
                int pCnt = 1;

                // 部署コード
                string bushoCode = string.Empty;

                // 部署別残業理由シートの内容を配列に取得する
                object[,] zReSeizou = bs.getZanReason();

                // 部署別残業理由別残業計画シートの内容を配列に取得する
                object[,] zRe = bs.getZanReasonPlan();

                // 部署別残業計画シートの内容を配列に取得する
                bs.zpArray = bs.getZanPlan();

                // progressBar
                int nMax = bs.zpArray.GetLength(0);
                toolStripProgressBar1.Maximum = nMax;
                toolStripProgressBar1.Minimum = 0;
                toolStripProgressBar1.Visible = true;

                // 一番最近の勤務票日付
                int maxDay = dts.過去勤務票ヘッダ.Where(a => a.年 == yy && a.月 == mm).Max(a => a.日);

                // 部署別残業計画配列を順番に読む
                for (int i = 1; i <= nMax; i++)
                {
                    // progressBar表示
                    label1.Visible = true;
                    label1.Text = "残業推移グラフ作成中..." + i + "/" + nMax;
                    toolStripProgressBar1.Value = i;
                    this.Refresh(); // ← 追加

                    // 部署コードが5桁未満のときは対象外
                    if (bs.zpArray[i, 1].ToString().Length < 5)
                    {
                        continue;
                    }

                    // 対象年月以外のとき
                    if (Utility.StrtoInt(bs.zpArray[i, 2].ToString()) != yy || Utility.StrtoInt(bs.zpArray[i, 3].ToString()) != mm)
                    {
                        continue;
                    }

                    // 人員数が「０」のときは対象外
                    if (Utility.StrtoInt(bs.zpArray[i, 4].ToString()) == global.flgOff)
                    {
                        continue;
                    }

                    // 部署指定のとき該当するか
                    if (comboBox2.SelectedIndex != -1)
                    {
                        Utility.ComboBumon cmb = (Utility.ComboBumon)comboBox2.SelectedItem;
                        if (bs.zpArray[i, 1].ToString() != cmb.code.ToString())
                        {
                            continue;
                        }
                    }

                    bushoCode = bs.zpArray[i, 1].ToString();                        // 部署コードを取得
                    sNin = Utility.StrtoInt(bs.zpArray[i, 4].ToString());           // 人数取得
                    sSeisan = Utility.StrtoInt(bs.zpArray[i, 6].ToString());        // 生産数取得

                    // 前月の残業合計、生産数、人数を取得
                    zenSeisanNin(zengetsuArray, bushoCode, zYY, zMM, ref zSeisan, ref zNin, ref zenZan, ref zenKaDays);

                    // 該当部署の当月計画値を取得する
                    double zanPlan = getZanPlan(bs.zpArray, yy, mm, bushoCode);

                    // 日付別の配列を生成
                    clsZanSum[] z = new clsZanSum[1];
                    dayArrayNew(yy, mm, bushoCode, ref z);

                    // 日別残業時間を配列にセットする・残業月間合計を取得する
                    double zanTotal = 0;
                    setDaybyZan(bushoCode, ref z, ref zanTotal);

                    // 残業計画時間の稼働日数割りと日々目標値を配列にセットする
                    setDaybyPlan(ref z, zanPlan);

                    // 残業時間の実績累積を配列にセットする
                    setDaybyZisseki(ref z, yy, mm, maxDay);

                    // 月間残業合計を時間単位に変換
                    zanTotal = Utility.StrtoDouble(((zanTotal / 60).ToString("#,##0.0")));

                    // エクセルシート出力
                    // テンプレートシートを追加する
                    pCnt++;
                    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];
                    oxlsSheet.Copy(Type.Missing, oXlsBook.Sheets[pCnt - 1]);
                    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt];

                    // シートにデータを貼り付ける
                    xlsOutPut(oXls, ref oXlsBook, ref oxlsSheet, z, Properties.Settings.Default.xlsZanChart, yy, mm, zReSeizou, zRe, sdCon);
                }

                System.Threading.Thread.Sleep(1000);

                // 1枚目はテンプレートシートなので印刷時には削除する
                oXls.DisplayAlerts = false;
                oXlsBook.Sheets[1].Delete();

                // 1枚目のシートが表示されるようにする
                oxlsSheet = oXlsBook.Sheets[1];
                oxlsSheet.Select();

                //マウスポインタを元に戻す
                this.Cursor = Cursors.Default;

                // 確認のためExcelのウィンドウを表示する
                oXls.Visible = true;

                //印刷
                oXlsBook.PrintOutEx(Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //oXlsBook.PrintOutEx();
                //oXlsBook.PrintPreview(true);

                // ウィンドウを非表示にする
                oXls.Visible = false;

                //保存処理
                oXls.DisplayAlerts = false;

                DialogResult ret;

                //ダイアログボックスの初期設定
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Title = "残業推移グラフ";
                saveFileDialog1.OverwritePrompt = true;
                saveFileDialog1.RestoreDirectory = true;
                DateTime dt = DateTime.Now;
                saveFileDialog1.FileName = yy + "年" + mm.ToString().PadLeft(2, ' ') + "月 残業推移グラフ";
                saveFileDialog1.Filter = "Microsoft Office Excelファイル(*.xlsx)|*.xlsx|全てのファイル(*.*)|*.*";

                //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
                string fileName;
                ret = saveFileDialog1.ShowDialog();

                if (ret == System.Windows.Forms.DialogResult.OK)
                {
                    fileName = saveFileDialog1.FileName;
                    oXlsBook.SaveAs(fileName, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing,
                                    Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }

                // プログレスバーを非表示
                toolStripProgressBar1.Visible = false;
                label1.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                // 奉行データベース接続切断
                if (sdCon.Cn.State == ConnectionState.Open)
                {
                    sdCon.Close();
                }

                // Bookをクローズ
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                // Excelを終了
                oXls.Quit();

                // COM オブジェクトの参照カウントを解放する 
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                oXls = null;
                oXlsBook = null;
                oxlsSheet = null;

                GC.Collect();
            }
        }

        ///---------------------------------------------------------------------------------
        /// <summary>
        ///     前月の生産数、人数、残業合計、稼働日数を取得する </summary>
        /// <param name="zen">
        ///     前月残業実績配列</param>
        /// <param name="bCode">
        ///     部署コード</param>
        /// <param name="zyy">
        ///     前月年</param>
        /// <param name="zmm">
        ///     前月</param>
        /// <param name="zSei">
        ///     前月生産数</param>
        /// <param name="zNi">
        ///     前月人数</param>
        /// <param name="zz">
        ///     前月残業合計</param>
        /// <param name="kDays">
        ///     前月稼働日数</param>
        ///---------------------------------------------------------------------------------
        private void zenSeisanNin(string [,] zen, string bCode, int zyy, int zmm, ref int zSei, ref int zNi, ref double zz, ref int kDays)
        {
            for (int i = 1; i <= bs.zpArray.GetLength(0); i++)
            {
                if (bs.zpArray[i, 1].ToString() == bCode && 
                    Utility.StrtoInt(bs.zpArray[i, 2].ToString()) == zyy && 
                    Utility.StrtoInt(bs.zpArray[i, 3].ToString()) == zmm)
                {
                    // 生産数、人数を取得
                    zNi = Utility.StrtoInt(bs.zpArray[i, 4].ToString());
                    zSei = Utility.StrtoInt(bs.zpArray[i, 6].ToString());
                    break;
                }
            }

            // 前月残業合計を取得
            zz = 0;
            for (int i = 0; i < zen.GetLength(0); i++)
			{
                if (zen[i, 0] == bCode)
                {
                    zz += Utility.StrtoDouble(zen[i, 1]);
                }
			}

            // 月間残業合計を時間単位に変換
            zz = Utility.StrtoDouble(((zz / 60).ToString("#,##0.0")));

            // 前月稼働日数を求める
            DateTime dt01 = DateTime.Parse(zyy + "/" + zmm + "/01");
            DateTime dtEnd = dt01.AddMonths(1).AddDays(-1); // 対象月末日

            DateTime dtTo = dt01;
            int iX = 0;

            while (dtTo <= dtEnd)
            {
                if (!dts.休日.Any(a => a.年月日 == dtTo))
                {
                    // 休日でない日をカウント
                    iX++;
                }

                dtTo = dtTo.AddDays(1);
            }

            // 稼働日数
            kDays = iX;
        }

        ///---------------------------------------------------------------------------------
        /// <summary>
        ///     前月の生産数、人数、残業合計、稼働日数を取得する </summary>
        /// <param name="zen">
        ///     前月残業実績配列</param>
        /// <param name="bCode">
        ///     部署コード</param>
        /// <param name="zyy">
        ///     前月年</param>
        /// <param name="zmm">
        ///     前月</param>
        /// <param name="zSei">
        ///     前月生産数</param>
        /// <param name="zNi">
        ///     前月人数</param>
        /// <param name="zz">
        ///     前月残業合計</param>
        /// <param name="kDays">
        ///     前月稼働日数</param>
        ///---------------------------------------------------------------------------------
        private void zenSeisanNinBumon(string[,] zen, string bmnCode, int zyy, int zmm, ref int zSei, ref int zNi, ref double zz, ref int kDays)
        {
            for (int i = 1; i <= bs.zpArray.GetLength(0); i++)
            {
                if (bmnCode == "0")
                {
                    // 全社
                    if (Utility.StrtoInt(bs.zpArray[i, 2].ToString()) == zyy &&
                        Utility.StrtoInt(bs.zpArray[i, 3].ToString()) == zmm)
                    {
                        // 生産数、人数を取得
                        zNi += Utility.StrtoInt(bs.zpArray[i, 4].ToString());
                        zSei += Utility.StrtoInt(bs.zpArray[i, 6].ToString());
                    }
                }
                else if (bmnCode == "1")
                {
                    // 直接部門
                    if (bs.zpArray[i, 1].ToString().Substring(1, 1) == bmnCode &&
                        Utility.StrtoInt(bs.zpArray[i, 2].ToString()) == zyy &&
                        Utility.StrtoInt(bs.zpArray[i, 3].ToString()) == zmm)
                    {
                        // 生産数、人数を取得
                        zNi += Utility.StrtoInt(bs.zpArray[i, 4].ToString());
                        zSei += Utility.StrtoInt(bs.zpArray[i, 6].ToString());
                    }
                }
                else if (bmnCode == "2")
                {
                    // 間接部門
                    if (bs.zpArray[i, 1].ToString().Substring(1, 1) != "0" &&
                        bs.zpArray[i, 1].ToString().Substring(1, 1) != "1" && 
                        Utility.StrtoInt(bs.zpArray[i, 2].ToString()) == zyy &&
                        Utility.StrtoInt(bs.zpArray[i, 3].ToString()) == zmm)
                    {
                        // 生産数、人数を取得
                        zNi += Utility.StrtoInt(bs.zpArray[i, 4].ToString());
                        zSei += Utility.StrtoInt(bs.zpArray[i, 6].ToString());
                    }
                }
            }

            // 前月残業合計を取得
            zz = 0;
            for (int i = 0; i < zen.GetLength(0); i++)
            {
                if (zen[i, 0] == bmnCode)
                {
                    zz += Utility.StrtoDouble(zen[i, 1]);
                }
            }

            // 月間残業合計を時間単位に変換
            zz = Utility.StrtoDouble(((zz / 60).ToString("#,##0.0")));

            // 前月稼働日数を求める
            DateTime dt01 = DateTime.Parse(zyy + "/" + zmm + "/01");
            DateTime dtEnd = dt01.AddMonths(1).AddDays(-1); // 対象月末日

            DateTime dtTo = dt01;
            int iX = 0;

            while (dtTo <= dtEnd)
            {
                if (!dts.休日.Any(a => a.年月日 == dtTo))
                {
                    // 休日でない日をカウント
                    iX++;
                }

                dtTo = dtTo.AddDays(1);
            }

            // 稼働日数
            kDays = iX;
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     エクセルシートへ出力 </summary>
        /// <param name="oXls">
        ///     Excel.Applicationオブジェクト</param>
        /// <param name="oXlsBook">
        ///     Excel.Workbookオブジェクト</param>
        /// <param name="oxlsSheet">
        ///     Excel.Worksheetオブジェクト</param>
        /// <param name="zd">
        ///     日別配列</param>
        /// <param name="xlsFile">
        ///     エクセルファイルパス</param>
        /// <param name="yy">
        ///     対象年 </param>
        /// <param name="mm">
        ///     対象月 </param>
        /// <param name="zReSeizou">
        ///     部署別残業理由配列</param>
        /// <param name="zRe">
        ///     部署別理由別残業計画配列</param>
        /// <param name="sdCon">
        ///     sqlControl.DataControl</param>
        ///-------------------------------------------------------------------
        private void xlsOutPut(Excel.Application oXls, ref Excel.Workbook oXlsBook, ref Excel.Worksheet oxlsSheet, clsZanSum[] zd, string xlsFile, int yy, int mm, object[,] zReSeizou, object [,] zRe, sqlControl.DataControl sdCon)
        {
            oxlsSheet.Select(Type.Missing);

            Excel.Range rng = null;

            int iX = 1;

            // 実稼働日数を取得
            int kDays = zd.Where(a => a.sHoliday == 0 || a.sZangyo > 0).Count();

            // 前回の書き込みセルを初期化する
            rng = oxlsSheet.Range[oxlsSheet.Cells[3, 31], oxlsSheet.Cells[7, 61]];
            rng.Value2 = "";

            // 2次元配列
            object[,] xlsArray = rng.Value2;

            foreach (var t in zd.Where(a => a.sHoliday == 0 || a.sZangyo > 0).OrderBy(a => a.sDay))
            {
                xlsArray[1, iX] = t.sDay.ToString();
                xlsArray[2, iX] = t.sZangyo.ToString();
                xlsArray[3, iX] = t.sMonthPlan.ToString();
                xlsArray[4, iX] = t.sZissekibyDay.ToString();
                xlsArray[5, iX] = t.sPlanbyDay.ToString();
                iX++;
            }
            
            try
            {
                // ウィンドウを非表示にする
                oXls.Visible = false;
                oXls.DisplayAlerts = false;

                oxlsSheet.Cells[1, 2] = yy + "年" + mm.ToString().PadLeft(2, ' ') + "月 残業推移グラフ";

                // 部署コード、名称
                string szCode = zd.First().sSzCode.ToString();
                oxlsSheet.Cells[1, 11] = szCode;
                oxlsSheet.Cells[1, 13] = getDepartmentName(_dbName, szCode, sdCon);
                
                // シート名に部署名をつける
                oxlsSheet.Name = szCode + " " + oxlsSheet.Cells[1, 13].value;

                // 実稼働日数を求める
                int d = zd.Where(a => a.sHoliday == 0 || a.sZangyo > 0).Count();
                oxlsSheet.Cells[1, 21] = d.ToString();

                // 集計期間
                int maxDay = 0;
                if (dts.過去勤務票ヘッダ.Any(a => a.年 == yy && a.月 == mm && a.部署コード == szCode))
                {
                    maxDay = dts.過去勤務票ヘッダ.Where(a => a.年 == yy && a.月 == mm && a.部署コード == szCode).Max(a => a.日);
                    oxlsSheet.Cells[1, 24] = mm.ToString().PadLeft(2, ' ') + "月" + maxDay.ToString().PadLeft(2, ' ') + "日まで";
                }
                else
                {
                    maxDay = 0;
                    oxlsSheet.Cells[1, 24] = "勤怠データなし";
                }

                // 集計日数
                int n = zd.Count(a => (a.sHoliday == 0 || a.sZangyo > 0) && a.sDay <= maxDay);
                oxlsSheet.Cells[1, 28] = n;

                // グラフ用データ一括書き込み　← 書き込むと数値が文字扱いとなりグラフが描画されない
                //rng = oxlsSheet.Range[oxlsSheet.Cells[4, 31], oxlsSheet.Cells[7, 30 + d]];
                //rng.NumberFormatLocal = "0.0";
                rng = oxlsSheet.Range[oxlsSheet.Cells[3, 31], oxlsSheet.Cells[7, 61]];
                rng.Value2 = xlsArray;

                iX = 0;

                //// グラフ用データをセルに書き込む（セル個別）
                //foreach (var t in zd.Where(a => a.sHoliday == 0 || a.sZangyo > 0).OrderBy(a => a.sDay))
                //{
                //    oxlsSheet.Cells[3, 31 + iX] = t.sDay.ToString();
                //    oxlsSheet.Cells[4, 31 + iX] = t.sZangyo.ToString();
                //    oxlsSheet.Cells[5, 31 + iX] = t.sMonthPlan.ToString();
                //    oxlsSheet.Cells[6, 31 + iX] = t.sZissekibyDay.ToString();
                //    oxlsSheet.Cells[7, 31 + iX] = t.sPlanbyDay.ToString();

                //    iX++;
                //}
                
                int iR = 0;
                double zRePlanTl = 0;

                // 理由別残業計画シートの内容を配列に取得
                if (Utility.StrtoInt(szCode.Substring(1, 1)) <= global.flgOn)
                {
                    // 製造部門：理由別残業計画はなし
                    zReSeizou = bs.getZanReason();      // 残業理由配列

                    for (int i = 1; i <= zReSeizou.GetLength(0); i++)
                    {
                        // 部署コードが一致しているか？
                        if (zReSeizou[i, 1].ToString() != szCode.ToString())
                        {
                            continue;
                        }

                        oxlsSheet.Cells[5 + iR, 20] = zReSeizou[i, 2];    // 理由コード
                        oxlsSheet.Cells[5 + iR, 21] = zReSeizou[i, 3];    // 残業理由
                        
                        // 理由別残業時間を集計
                        if (comboBox1.SelectedIndex == 0)
                        {
                            double s = dts.残業集計.Where(a => a.部署コード == szCode && a.残業理由 == (double)zReSeizou[i, 2])
                                .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));

                            s = s / 60;
                            oxlsSheet.Cells[5 + iR, 27] = s;
                        }
                        else
                        {
                            double s = dts.残業集計.Where(a => a.応援先 == szCode && a.残業理由 == (double)zReSeizou[i, 2])
                                .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));

                            s = s / 60;
                            oxlsSheet.Cells[5 + iR, 27] = s;
                        }

                        iR += 2;

                        if (iR == 20)
                        {
                            // 合計欄に当月計画時間の稼働日数割りを表示する
                            if (zd.Any(a => a.sDay == maxDay))
                            {
                                foreach (var t in zd.Where(a => a.sDay == maxDay))
                                {
                                    oxlsSheet.Cells[25, 25] = t.sMonthPlan;
                                    break;
                                }
                            }
                            else
                            {
                                oxlsSheet.Cells[25, 25] = 0;
                            }
                            break;
                        }
                    }
                }
                else
                {
                    // 間接部門：理由別残業計画シートの内容を配列に取得
                    zRe = bs.getZanReasonPlan();

                    for (int i = 1; i <= zRe.GetLength(0); i++)
                    {
                        // 部署コードが一致しているか？
                        if (zRe[i, 1].ToString() != szCode.ToString())
                        {
                            continue;
                        }

                        // 年月コードが一致しているか？
                        if (zRe[i, 2].ToString() != yy.ToString() && zRe[i, 3].ToString() != mm.ToString())
                        {
                            continue;
                        }

                        oxlsSheet.Cells[5 + iR, 20] = zRe[i, 4];    // 理由コード
                        oxlsSheet.Cells[5 + iR, 21] = zRe[i, 5];    // 残業理由
                        oxlsSheet.Cells[5 + iR, 25] = Utility.StrtoDouble(zRe[i, 6].ToString()) / (double)kDays * n;    // 現在の日付で稼働日割りした理由別計画値

                        if (comboBox1.SelectedIndex == 0)
                        {
                            // 理由別残業時間を集計
                            double s = dts.残業集計.Where(a => a.部署コード == szCode && a.残業理由 == (double)zRe[i, 4])
                                .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));

                            s = s / 60;
                            oxlsSheet.Cells[5 + iR, 27] = s;
                        }
                        else
                        {
                            // 理由別残業時間を集計
                            double s = dts.残業集計.Where(a => a.応援先 == szCode && a.残業理由 == (double)zRe[i, 4])
                                .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));

                            s = s / 60;
                            oxlsSheet.Cells[5 + iR, 27] = s;
                        }
                        
                        iR += 2;
                        //zRePlanTl += Utility.StrtoDouble(zRe[i, 6].ToString()); // 計画合計に加算
                        zRePlanTl += Utility.StrtoDouble(zRe[i, 6].ToString()) / (double)kDays * n;     // 計画合計に加算

                        if (iR == 20)
                        {
                            // 現在の日付で稼働日割りした理由別計画の合計
                            oxlsSheet.Cells[25, 25] = zRePlanTl;
                            break;
                        }
                    }
                }

                // 前月との比較要素のセット
                oxlsSheet.Cells[35, 23] = zSeisan;              // 前月生産数
                oxlsSheet.Cells[37, 23] = zNin;                 // 前月人数
                //oxlsSheet.Cells[44, 23] = zenZan;               // 前月残業合計　※チェック用
                oxlsSheet.Cells[41, 23] = zenZan / zNin;        // 前月1人当り残業
                //oxlsSheet.Cells[45, 23] = zenKaDays;            // 前月稼働日数　※チェック用
                oxlsSheet.Cells[39, 23] = zenZan / zenKaDays;   // 前月一日当たり合計

                oxlsSheet.Cells[35, 25] = sSeisan;              // 生産数
                oxlsSheet.Cells[37, 25] = sNin;                 // 人数

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "残業グラフエクセル出力", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); 
            }
            finally
            {
            }
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     エクセルシートへ出力 </summary>
        /// <param name="oXls">
        ///     Excel.Applicationオブジェクト</param>
        /// <param name="oXlsBook">
        ///     Excel.Workbookオブジェクト</param>
        /// <param name="oxlsSheet">
        ///     Excel.Worksheetオブジェクト</param>
        /// <param name="zd">
        ///     日別配列</param>
        /// <param name="xlsFile">
        ///     エクセルファイルパス</param>
        /// <param name="yy">
        ///     対象年 </param>
        /// <param name="mm">
        ///     対象月 </param>
        /// <param name="zReSeizou">
        ///     部署別残業理由配列</param>
        /// <param name="zRe">
        ///     部署別理由別残業計画配列</param>
        /// <param name="sBmn">
        ///     部門コード</param>
        ///-------------------------------------------------------------------
        private void xlsOutPutBumon(Excel.Application oXls, ref Excel.Workbook oXlsBook, ref Excel.Worksheet oxlsSheet, clsZanSum[] zd, string xlsFile, int yy, int mm, object[,] zReSeizou, object[,] zRe, string sBmn)
        {
            oxlsSheet.Select(Type.Missing);

            Excel.Range rng = null;

            int iX = 1;

            // 実稼働日数を取得
            int kDays = zd.Where(a => a.sHoliday == 0 || a.sZangyo > 0).Count();

            // 前回の書き込みセルを初期化する
            rng = oxlsSheet.Range[oxlsSheet.Cells[3, 31], oxlsSheet.Cells[7, 61]];
            rng.Value2 = "";

            // 2次元配列
            object[,] xlsArray = rng.Value2;

            foreach (var t in zd.Where(a => a.sHoliday == 0 || a.sZangyo > 0).OrderBy(a => a.sDay))
            {
                xlsArray[1, iX] = t.sDay.ToString();
                xlsArray[2, iX] = t.sZangyo.ToString();
                xlsArray[3, iX] = t.sMonthPlan.ToString();
                xlsArray[4, iX] = t.sZissekibyDay.ToString();
                xlsArray[5, iX] = t.sPlanbyDay.ToString();
                iX++;
            }
            
            try
            {
                // ウィンドウを非表示にする
                oXls.Visible = false;
                oXls.DisplayAlerts = false;

                oxlsSheet.Cells[1, 2] = yy + "年" + mm.ToString().PadLeft(2, ' ') + "月 残業推移グラフ";

                // 部門、名称
                string szCode = zd.First().sSzCode.ToString();
                oxlsSheet.Cells[1, 11] = szCode.PadLeft(6, '0');
                                
                if (szCode == "0")
                {
                    oxlsSheet.Cells[1, 13] = global.BMN_ALL;
                }
                else if (szCode == "1")
                {
                    oxlsSheet.Cells[1, 13] = global.BMN_SEIZOU;
                }
                else if (szCode == "2")
                {
                    oxlsSheet.Cells[1, 13] = global.BMN_KANSETSU;
                }

                // シート名に部署名をつける
                oxlsSheet.Name = szCode.PadLeft(6, '0') + " " + oxlsSheet.Cells[1, 13].value;

                // 実稼働日数を求める
                int d = zd.Where(a => a.sHoliday == 0 || a.sZangyo > 0).Count();
                oxlsSheet.Cells[1, 21] = d.ToString();

                // 集計期間
                int maxDay = 0;

                switch (szCode)
                {
                    case "0":

                        // 全社
                        if (dts.過去勤務票ヘッダ.Any(a => a.年 == yy && a.月 == mm))
                        {
                            maxDay = dts.過去勤務票ヘッダ.Where(a => a.年 == yy && a.月 == mm).Max(a => a.日);
                            oxlsSheet.Cells[1, 24] = mm.ToString().PadLeft(2, ' ') + "月" + maxDay.ToString().PadLeft(2, ' ') + "日まで";
                        }
                        else
                        {
                            maxDay = 0;
                            oxlsSheet.Cells[1, 24] = "勤怠データなし";
                        }

                        break;

                    case "1":

                        // 製造部門
                        if (dts.過去勤務票ヘッダ.Any(a => a.年 == yy && a.月 == mm && a.部署コード.Substring(1, 1) == szCode))
                        {
                            maxDay = dts.過去勤務票ヘッダ.Where(a => a.年 == yy && a.月 == mm && a.部署コード.Substring(1, 1) == szCode).Max(a => a.日);
                            oxlsSheet.Cells[1, 24] = mm.ToString().PadLeft(2, ' ') + "月" + maxDay.ToString().PadLeft(2, ' ') + "日まで";
                        }
                        else
                        {
                            maxDay = 0;
                            oxlsSheet.Cells[1, 24] = "勤怠データなし";
                        }
                        break;

                    case "2":

                        // 間接部門
                        if (dts.過去勤務票ヘッダ.Any(a => a.年 == yy && a.月 == mm && a.部署コード.Substring(1, 1) != "1" && a.部署コード.Substring(1, 1) != "0"))
                        {
                            maxDay = dts.過去勤務票ヘッダ.Where(a => a.年 == yy && a.月 == mm && a.部署コード.Substring(1, 1) != "1" && a.部署コード.Substring(1, 1) != "0").Max(a => a.日);
                            oxlsSheet.Cells[1, 24] = mm.ToString().PadLeft(2, ' ') + "月" + maxDay.ToString().PadLeft(2, ' ') + "日まで";
                        }
                        else
                        {
                            maxDay = 0;
                            oxlsSheet.Cells[1, 24] = "勤怠データなし";
                        }

                        break;

                    default:
                        break;
                }
                
                // 集計日数
                int n = zd.Count(a => (a.sHoliday == 0 || a.sZangyo > 0) && a.sDay <= maxDay);
                oxlsSheet.Cells[1, 28] = n;
                
                // グラフ用データ一括書き込み　← 書き込むと数値が文字扱いとなりグラフが描画されない
                //rng = oxlsSheet.Range[oxlsSheet.Cells[4, 31], oxlsSheet.Cells[7, 61]];
                //rng.NumberFormatLocal = "0.0";
                rng = oxlsSheet.Range[oxlsSheet.Cells[3, 31], oxlsSheet.Cells[7, 61]];
                rng.Value2 = xlsArray;

                iX = 0;

                //// グラフ用データをセルに書き込む（セル個別）
                //foreach (var t in zd.Where(a => a.sHoliday == 0 || a.sZangyo > 0).OrderBy(a => a.sDay))
                //{
                //    oxlsSheet.Cells[3, 31 + iX] = t.sDay.ToString();
                //    oxlsSheet.Cells[4, 31 + iX] = t.sZangyo.ToString();
                //    oxlsSheet.Cells[5, 31 + iX] = t.sMonthPlan.ToString();
                //    oxlsSheet.Cells[6, 31 + iX] = t.sZissekibyDay.ToString();
                //    oxlsSheet.Cells[7, 31 + iX] = t.sPlanbyDay.ToString();

                //    iX++;
                //}

                int iR = 0;
                double zRePlanTl = 0;
                double zOver10 = 0;
                double zzz = 0;

                // 理由別残業計画
                if (szCode == "0" || szCode == "1")
                {
                    // 製造部門：理由別残業計画はなし                    
                    oxlsSheet.Cells[5 + iR, 20] = "";   // 理由コード
                    oxlsSheet.Cells[5 + iR, 21] = "";   // 残業理由

                    // 理由別残業時間を集計
                    IEnumerable<riyuZan> s = getRiyuZanSum(szCode);

                    foreach (var t in s)
                    {
                        if (t.riyu >= 10)
                        {
                            iR = 23;
                            zzz = t.zan / 60;
                            zOver10 += zzz;
                            oxlsSheet.Cells[iR, 27] = zOver10;
                        }
                        else
                        {
                            iR = (int)t.riyu * 2 + 3;
                            zzz = t.zan / 60;
                            oxlsSheet.Cells[iR, 27] = zzz;
                        }

                        oxlsSheet.Cells[iR, 20] = "";   // 理由コード
                        oxlsSheet.Cells[iR, 21] = "";   // 残業理由
                    }

                    // 合計欄に当月計画時間の稼働日数割りを表示する
                    if (zd.Any(a => a.sDay == maxDay))
                    {
                        foreach (var t in zd.Where(a => a.sDay == maxDay))
                        {
                            oxlsSheet.Cells[25, 25] = t.sMonthPlan;
                            break;
                        }
                    }
                    else
                    {
                        oxlsSheet.Cells[25, 25] = 0;
                    }
                }
                else if (Utility.StrtoInt(szCode) >= 2)
                {
                    // 間接部門：理由別残業計画シートの内容を配列に取得
                    zRe = bs.getZanReasonPlan();

                    double[] keikaku = new double[10];
                    for (int ind = 0; ind < keikaku.Length; ind++)
                    {
                        keikaku[ind] = 0;
                    }

                    // 理由別計画集計値を配列にセット
                    for (int i = 1; i <= zRe.GetLength(0); i++)
                    {
                        // 部署コードが一致しているか？
                        if (zRe[i, 1].ToString().Substring(1, 1) == "0" ||  
                            zRe[i, 1].ToString().Substring(1, 1) == "1")
                        {
                            continue;
                        }

                        // 年月コードが一致しているか？
                        if (zRe[i, 2].ToString() != yy.ToString() || zRe[i, 3].ToString() != mm.ToString())
                        {
                            continue;
                        }
                        
                        int p = Utility.StrtoInt(zRe[i, 4].ToString());

                        if (p > 10) p = 10;

                        keikaku[p - 1] += Utility.StrtoInt(zRe[i, 6].ToString());
                    }

                    // 理由別計画集計値配列を順次読む
                    for (int ind = 0; ind < keikaku.Length; ind++)
                    {
                        oxlsSheet.Cells[5 + iR, 20] = ind + 1;          // 理由コード
                        //oxlsSheet.Cells[5 + iR, 21] = zRe[i, 5];      // 残業理由
                        oxlsSheet.Cells[5 + iR, 21] = "";               // 残業理由
                        oxlsSheet.Cells[5 + iR, 25] = keikaku[ind] / (double)kDays * n;     // 現在の日付で稼働日割りした理由別計画値

                        double s = 0;

                        // 理由別残業時間を集計
                        if (comboBox1.SelectedIndex == 0)
                        {
                            // 社員所属で集計
                            if (ind < 9)
                            {
                                s = dts.残業集計.Where(a => a.部署コード.Substring(1, 1) != "1" && a.部署コード.Substring(1, 1) != "0" && a.残業理由 == (double)ind + 1)
                                    .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));
                            }
                            else
                            {
                                s = dts.残業集計.Where(a => a.部署コード.Substring(1, 1) != "1" && a.部署コード.Substring(1, 1) != "0" && a.残業理由 >= 10)
                                    .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));
                            }

                            s = s / 60;
                            oxlsSheet.Cells[5 + iR, 27] = s;
                        }
                        else
                        {
                            // 応援先で集計
                            if (ind < 9)
                            {
                                s = dts.残業集計.Where(a => a.応援先.Substring(1, 1) != "1" && a.応援先.Substring(1, 1) != "0" && a.残業理由 == (double)ind + 1)
                                    .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));
                            }
                            else
                            {
                                s = dts.残業集計.Where(a => a.応援先.Substring(1, 1) != "1" && a.応援先.Substring(1, 1) != "0" && a.残業理由 >= 10)
                                    .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));
                            }

                            s = s / 60;
                            oxlsSheet.Cells[5 + iR, 27] = s;
                        }

                        iR += 2;
                        zRePlanTl += keikaku[ind] / (double)kDays * n; // 計画合計に加算

                        if (iR == 20)
                        {
                            // 計画合計
                            oxlsSheet.Cells[25, 25] = zRePlanTl;
                        }
                    }
                }

                // 前月との比較要素のセット
                oxlsSheet.Cells[35, 23] = zSeisan;              // 前月生産数
                oxlsSheet.Cells[37, 23] = zNin;                 // 前月人数
                //oxlsSheet.Cells[44, 23] = zenZan;               // 前月残業合計　※チェック用
                oxlsSheet.Cells[41, 23] = zenZan / zNin;        // 前月1人当り残業
                //oxlsSheet.Cells[45, 23] = zenKaDays;            // 前月稼働日数　※チェック用
                oxlsSheet.Cells[39, 23] = zenZan / zenKaDays;   // 前月一日当たり合計

                oxlsSheet.Cells[35, 25] = sSeisan;              // 生産数
                oxlsSheet.Cells[37, 25] = sNin;                 // 人数

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "残業グラフ エクセル出力", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
            }
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     日々目標値をセットする </summary>
        /// <param name="zd">
        ///     日々配列</param>
        /// <param name="zPlan">
        ///     月計画残業時間</param>
        ///----------------------------------------------------------------------
        private void setDaybyGoal(ref clsZanSum[] zd, double zPlan)
        {
            // 稼働日数を取得
            int kDays = zd.Where(a => a.sHoliday == 0).Count();

            // 日々目標値
            int val = (int)(zPlan / kDays);

            // 日々目標値
            foreach (var t in zd.Where(a => a.sHoliday == 0 || a.sZangyo > 0).OrderBy(a => a.sDay))
            {
                t.sPlanbyDay = val;
            }
        }


        ///-----------------------------------------------------------------------
        /// <summary>
        ///     残業時間の実績累積を配列にセットする </summary>
        /// <param name="zd">
        ///     日別配列</param>
        /// <param name="yy">
        ///     対象年</param>
        /// <param name="mm">
        ///     対象月</param>
        /// <param name="sMaxDay">
        ///     一番最近の勤務票日付</param>
        ///-----------------------------------------------------------------------
        private void setDaybyZisseki(ref clsZanSum[] zd, int yy, int mm, int sMaxDay)
        {
            double v = 0;

            foreach (var t in zd.Where(a => (a.sHoliday == 0 || a.sZangyo > 0) && a.sDay <= sMaxDay).OrderBy(a => a.sDay))
            {
                t.sZissekibyDay = (v + t.sZangyo).ToString();
                v += t.sZangyo;
            }
        }

        ///------------------------------------------------------------------------------
        /// <summary>
        ///     残業計画時間の稼働日数割りと日々目標値を配列にセットする </summary>
        /// <param name="bCode">
        ///     部署コード</param>
        /// <param name="z">
        ///     日別配列</param>
        /// <param name="zPlan">
        ///     月残業計画値</param>
        ///------------------------------------------------------------------------------
        private void setDaybyPlan(ref clsZanSum[] zd, double zPlan)
        {
            // 稼働日数を取得
            int kDays = zd.Where(a => a.sHoliday == 0).Count();

            // 日々目標値
            double val = zPlan / (double)kDays;

            int i = 0;

            foreach (var t in zd.Where(a => a.sHoliday == 0 || a.sZangyo > 0).OrderBy(a => a.sDay))
            {
                t.sMonthPlan = Utility.StrtoDouble((zPlan / (double)kDays * (i + 1)).ToString("#,##0.0"));
                t.sPlanbyDay = val;
                i++;
            }
        }


        ///----------------------------------------------------------------------
        /// <summary>
        ///     日別残業時間集計 </summary>
        /// <param name="bCode">
        ///     部署コード</param>
        /// <param name="z">
        ///     日付別配列</param>
        /// <param name="zTotal">
        ///     残業時間月合計</param>
        ///----------------------------------------------------------------------
        private void setDaybyZan(string bCode, ref clsZanSum[] z, ref double zTotal)
        {
            if (comboBox1.SelectedIndex == 0)
            {
                // 日別に残業時間を集計 ※社員所属で集計
                var d = dts.残業集計.Where(a => a.部署コード == bCode).GroupBy(a => a.日)
                    .Select(g => new
                    {
                        day = g.Key,
                        zanH = g.Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10))
                    })
                    .OrderBy(a => a.day);

                foreach (var t in d)
                {
                    // 月間残業合計
                    zTotal += t.zanH;

                    // 日別の残業時間を配列にセット
                    for (int iZ = 0; iZ < z.Length; iZ++)
                    {
                        if (z[iZ].sDay == t.day)
                        {
                            z[iZ].sZangyo = Utility.StrtoDouble(((t.zanH / 60).ToString("#,##0.0")));
                            break;
                        }
                    }
                }
            }
            else
            {
                // 日別に残業時間を集計 ※応援先部署で集計
                var d = dts.残業集計.Where(a => a.応援先 == bCode).GroupBy(a => a.日)
                    .Select(g => new
                    {
                        day = g.Key,
                        zanH = g.Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10))
                    })
                    .OrderBy(a => a.day);

                foreach (var t in d)
                {
                    // 月間残業合計
                    zTotal += t.zanH;

                    // 日別の残業時間を配列にセット
                    for (int iZ = 0; iZ < z.Length; iZ++)
                    {
                        if (z[iZ].sDay == t.day)
                        {
                            z[iZ].sZangyo = Utility.StrtoDouble(((t.zanH / 60).ToString("#,##0.0")));
                            break;
                        }
                    }
                }
            }
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     日別残業時間集計 </summary>
        /// <param name="bCode">
        ///     部門コード（0:全社、1:製造部門、2:間接部門）</param>
        /// <param name="z">
        ///     日付別配列</param>
        /// <param name="zTotal">
        ///     残業時間月合計</param>
        ///----------------------------------------------------------------------
        private void setDaybyZanBumon(string bCode, ref clsZanSum[] z, ref double zTotal)
        {
            IEnumerable<dayZan> d = getDaybyZanSum(bCode);

            foreach (var t in d)
            {
                // 月間残業合計
                zTotal += t.zanH;

                // 日別の残業時間を配列にセット
                for (int iZ = 0; iZ < z.Length; iZ++)
                {
                    if (z[iZ].sDay == t.day)
                    {
                        z[iZ].sZangyo = Utility.StrtoDouble(((t.zanH / 60).ToString("#,##0.0")));
                        break;
                    }
                }
            }
        }
        

        ///----------------------------------------------------------------------
        /// <summary>
        ///     該当部署の当月残業計画時間を取得する </summary>
        /// <param name="zz">
        ///     部署別残業計画配列</param>
        /// <param name="yy">
        ///     対象年</param>
        /// <param name="mm">
        ///     対象月</param>
        /// <param name="bushoCode">
        ///     部署コード</param>
        /// <returns>
        ///     当月残業計画時間</returns>
        ///----------------------------------------------------------------------
        private double getZanPlan(object[,] zz, int yy, int mm, string bushoCode)
        {
            double zanPlan = 0;

            for (int iZ = 1; iZ < zz.GetLength(0); iZ++)
            {
                // 対象年月以外のときは対象外
                if (Utility.StrtoInt(zz[iZ, 2].ToString()) != yy ||
                    Utility.StrtoInt(zz[iZ, 3].ToString()) != mm)
                {
                    continue;
                }

                // 該当部署の当月計画値を取得する
                if (zz[iZ, 1].ToString() == bushoCode)
                {
                    zanPlan = Utility.StrtoDouble(Utility.NulltoStr(zz[iZ, 5]));
                    break;
                }
            }

            return zanPlan;
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     該当部門の当月残業計画時間を取得する </summary>
        /// <param name="zz">
        ///     部署別残業計画配列</param>
        /// <param name="yy">
        ///     対象年</param>
        /// <param name="mm">
        ///     対象月</param>
        /// <param name="bushoCode">
        ///     部門コード</param>
        /// <returns>
        ///     当月残業計画時間</returns>
        ///----------------------------------------------------------------------
        private double getZanPlanBumon(object[,] zz, int yy, int mm, int bmnCode)
        {
            double zanPlan = 0;

            for (int iZ = 1; iZ < zz.GetLength(0); iZ++)
            {
                // 対象年月以外のときは対象外
                if (Utility.StrtoInt(zz[iZ, 2].ToString()) != yy ||
                    Utility.StrtoInt(zz[iZ, 3].ToString()) != mm)
                {
                    continue;
                }

                // 部署コードの頭2桁目が部門（１：製造、2以上：間接）
                int bmn = Utility.StrtoInt(zz[iZ, 1].ToString().Substring(1, 1));

                // 全社または部門の当月計画値を取得する
                switch (bmnCode)
                {
                    case 0: // 全社
                        zanPlan += Utility.StrtoDouble(Utility.NulltoStr(zz[iZ, 5]));
                        break;

                    case 1: // 製造部門
                        if (bmn == bmnCode)
                        {
                            zanPlan += Utility.StrtoDouble(Utility.NulltoStr(zz[iZ, 5]));
                        }
                        break;

                    case 2: // 間接部門
                        if (bmn >= bmnCode)
                        {
                            zanPlan += Utility.StrtoDouble(Utility.NulltoStr(zz[iZ, 5]));
                        }
                        break;

                    default:
                        break;
                }
            }

            return zanPlan;
        }

        ///--------------------------------------------------------------------
        /// <summary>
        ///     奉行マスターから部署名を取得する </summary>
        /// <param name="_dbName">
        ///     データベース名</param>
        /// <param name="dCode">
        ///     部署コード</param>
        /// <param name="sdCon">
        ///     sqlControl.DataControl</param>
        /// <returns>
        ///     部署名 </returns>
        ///--------------------------------------------------------------------
        private string getDepartmentName(string _dbName, string dCode, sqlControl.DataControl sdCon)
        {
            string b = string.Empty;
            string dName = string.Empty;

            // 検索用部署コード
            if (Utility.StrtoInt(dCode) != global.flgOff)
            {
                b = dCode.Trim().PadLeft(15, '0');
            }
            else
            {
                b = dCode.Trim().PadRight(15, ' ');
            }

            string dt = DateTime.Today.ToShortDateString();
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT DepartmentID, DepartmentCode, DepartmentName ");
            sb.Append("FROM tbDepartment ");
            sb.Append("where EstablishDate <= '").Append(dt).Append("'");
            sb.Append(" and AbolitionDate >= '").Append(dt).Append("'");
            sb.Append(" and ValidDate <= '").Append(dt).Append("'");
            sb.Append(" and InValidDate >= '").Append(dt).Append("'");
            sb.Append(" and DepartmentCode = '").Append(b).Append("'");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                dName = dR["DepartmentName"].ToString().Trim();
                break;
            }

            dR.Close();

            return dName;
        }

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     前月残業実績配列作成 </summary>
        /// <param name="yy">
        ///     前月の年</param>
        /// <param name="mm">
        ///     前月</param>
        /// <param name="zenArray">
        ///     配列</param>
        /// <param name="dts">
        ///     DataSet</param>
        ///-----------------------------------------------------------------------
        private void setZengetsuZan(int yy, int mm, ref string[,] zenArray, DataSet1 dts)
        {
            // 前月残業データ取得
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

            // リストでとる
            if (comboBox1.SelectedIndex == 0)
            {
                // 所属部署で集計
                var zenGetsu = dts.残業集計.Select(a => new
                {
                    busho = a.部署コード,
                    zan = (a.残業時 * 60) + (a.残業分 * 60 / 10)
                }).ToList();

                // 前月残業実績配列を作成
                zenArray = new string[zenGetsu.Count(), 2];

                int ix = 0;
                foreach (var t in zenGetsu)
                {
                    zenArray[ix, 0] = t.busho;
                    zenArray[ix, 1] = t.zan.ToString();
                    ix++;
                }
            }
            else
            {
                // 応援先で集計
                var zenGetsu = dts.残業集計.Select(a => new
                {
                    busho = a.応援先,
                    zan = (a.残業時 * 60) + (a.残業分 * 60 / 10)
                }).ToList();

                // 前月残業実績配列を作成
                zenArray = new string[zenGetsu.Count(), 2];

                int ix = 0;
                foreach (var t in zenGetsu)
                {
                    zenArray[ix, 0] = t.busho;
                    zenArray[ix, 1] = t.zan.ToString();
                    ix++;
                }
            }
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     人数、生産数を取得する </summary>
        /// <param name="bmnCode">
        ///     0:全社、1:製造部門、2:間接部門</param>
        /// <param name="yy">
        ///     対象年</param>
        /// <param name="mm">
        ///     対象月</param>
        /// <param name="zNi">
        ///     人数</param>
        /// <param name="zSei">
        ///     生産数</param>
        ///------------------------------------------------------------------
        private void getBumonNin(string bmnCode, int yy, int mm, ref int zNi, ref int zSei)
        {
            for (int i = 1; i <= bs.zpArray.GetLength(0); i++)
            {
                if (bmnCode == "0")
                {
                    // 全社
                    if (Utility.StrtoInt(bs.zpArray[i, 2].ToString()) == yy &&
                        Utility.StrtoInt(bs.zpArray[i, 3].ToString()) == mm)
                    {
                        // 生産数、人数を取得
                        zNi += Utility.StrtoInt(bs.zpArray[i, 4].ToString());
                        zSei += Utility.StrtoInt(bs.zpArray[i, 6].ToString());
                    }
                }
                else if (bmnCode == "1")
                {
                    // 製造部門
                    if (bs.zpArray[i, 1].ToString().Substring(1, 1) == bmnCode &&
                        Utility.StrtoInt(bs.zpArray[i, 2].ToString()) == yy &&
                        Utility.StrtoInt(bs.zpArray[i, 3].ToString()) == mm)
                    {
                        // 生産数、人数を取得
                        zNi += Utility.StrtoInt(bs.zpArray[i, 4].ToString());
                        zSei += Utility.StrtoInt(bs.zpArray[i, 6].ToString());
                    }
                }
                else if (bmnCode == "2")
                {
                    // 間接部門
                    if (bs.zpArray[i, 1].ToString().Substring(1, 1) != "0" &&
                        bs.zpArray[i, 1].ToString().Substring(1, 1) != "1" &&
                        Utility.StrtoInt(bs.zpArray[i, 2].ToString()) == yy &&
                        Utility.StrtoInt(bs.zpArray[i, 3].ToString()) == mm)
                    {
                        // 生産数、人数を取得
                        zNi += Utility.StrtoInt(bs.zpArray[i, 4].ToString());
                        zSei += Utility.StrtoInt(bs.zpArray[i, 6].ToString());
                    }
                }
            }
        }

        ///--------------------------------------------------------------------
        /// <summary>
        ///     理由別残業時間 </summary>
        /// <param name="bmnCode">
        ///     部門コード</param>
        /// <returns>
        ///     IEnumerable<riyuZan> </returns>
        ///--------------------------------------------------------------------
        private IEnumerable<riyuZan> getRiyuZanSum(string bmnCode)
        {
            EnumerableRowCollection<DataSet1.残業集計Row> ss;

            if (comboBox1.SelectedIndex == 0)
            {
                // 社員所属で集計
                if (bmnCode == "0")
                {
                    ss = dts.残業集計.Where(a => a.部署コード != "");
                }
                else
                {
                    ss = dts.残業集計.Where(a => a.部署コード.Substring(1, 1) == bmnCode);
                }
            }
            else
            {
                // 応援先で集計
                if (bmnCode == "0")
                {
                    ss = dts.残業集計.Where(a => a.応援先 != "");
                }
                else
                {
                    ss = dts.残業集計.Where(a => a.応援先.Substring(1, 1) == bmnCode);
                }
            }
            
            // 理由別残業時間を集計
            IEnumerable<riyuZan> s = ss
                .OrderBy(a => a.残業理由)
                .GroupBy(a => a.残業理由)
                .Select(r => new riyuZan 
                {
                    riyu = r.Key,
                    zan = r.Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10))
                });

            return s;
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     日別の残業時間を集計 </summary>
        /// <param name="bmnCode">
        ///     部門コード</param>
        /// <returns>
        ///     IEnumerable<dayZan> </returns>
        ///----------------------------------------------------------------
        private IEnumerable<dayZan> getDaybyZanSum(string bmnCode)
        {
            EnumerableRowCollection<DataSet1.残業集計Row> d;

            if (bmnCode == "0")
            {
                // 全社
                d = dts.残業集計.Where(a => a.部署コード != ""); 
            }
            else
            {
                if (comboBox1.SelectedIndex == 0)
                {
                    // 部門で絞込み
                    d = dts.残業集計.Where(a => a.部署コード.Substring(1, 1) == bmnCode);
                }
                else
                {
                    // 応援先で絞込み
                    d = dts.残業集計.Where(a => a.応援先.Substring(1, 1) == bmnCode);
                }
            }
            
            // 日別の残業時間を集計
            IEnumerable<dayZan> s = d.GroupBy(a => a.日)
                .Select(g => new dayZan
                {
                    day = g.Key,
                    zanH = g.Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10))
                })
                .OrderBy(a => a.day);

            return s;
        }

        ///----------------------------------------------
        /// <summary>
        ///     理由別集計残業時間クラス </summary>
        ///----------------------------------------------
        private class riyuZan
        {
            public double riyu;
            public double zan;
        }

        ///----------------------------------------------
        /// <summary>
        ///     日別集計残業クラス </summary>
        ///----------------------------------------------
        private class dayZan
        {
            public int day;
            public double zanH;
        }

        private void rBtn1_CheckedChanged(object sender, EventArgs e)
        {
            if (rBtn1.Checked)
            {
                comboBox2.Enabled = true;
            }
            else
            {
                comboBox2.Enabled = false;
            }
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            prtReport();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Close();
        }

        private void frmZanChartXls_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }
    }
}
