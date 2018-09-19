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
using LINQtoCSV;

namespace SZDS_TIMECARD.sumData
{
    public partial class frmZanChartXls_New201804 : Form
    {
        public frmZanChartXls_New201804(string dbName)
        {
            InitializeComponent();

            hAdp.Fill(dts.過去勤務票ヘッダ);
            dAdp.Fill(dts.休日);

            _dbName = dbName;
        }

        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.残業集計TableAdapter adp = new DataSet1TableAdapters.残業集計TableAdapter();
        DataSet1TableAdapters.残業集計1TableAdapter adp_ka = new DataSet1TableAdapters.残業集計1TableAdapter();
        DataSet1TableAdapters.残業集計2TableAdapter adp_kakari = new DataSet1TableAdapters.残業集計2TableAdapter();     // 2018/04/09
        DataSet1TableAdapters.残業集計3TableAdapter adp_Han = new DataSet1TableAdapters.残業集計3TableAdapter();        // 2018/04/16
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

        // 奉行から出力した勤怠データ配列
        string[] workArray = null;

        // 奉行から出力した前月勤怠データ配列
        string[] zenArray = null;

        // 対象年月と最大日付
        int yy = 0, mm = 0, days = 0;
        
        // 前月
        int zYY = 0, zMM = 0;

        string outZangyoFile = @"c:\SZDS_OCR\xls\zangyo.csv";                   // 当月勤務実績出力ファイル
        string outZengetsuFile = @"c:\SZDS_OCR\xls\zengetsu.csv";               // 前月勤務実績出力ファイル
        string outKaZanPlanFile = @"c:\SZDS_OCR\xls\kaZanPlan.csv";             // 部署別残業計画出力ファイル
        string outKaByReByZanPlanFile = @"c:\SZDS_OCR\xls\kaByreByZanPlan.csv"; // 部署別理由別残業計画出力ファイル 2018/04/11

        bool openCsvZen = false;
        bool openCsvTou = false;
        int chart_Status = global.flgOff;
        int chart_StatusZen = global.flgOff;

        // 課別残業計画配列 2017/10/18
        string[,] zpArrayNew = null;

        // 部署別理由別残業計画の班・係・課単位での集計結果配列 2018/04/11
        string[] szByreByzanPlan = null;

        private void button1_Click(object sender, EventArgs e)
        {
        }

        private void prtReport()
        {
            this.Cursor = Cursors.WaitCursor;
                            
            if (rBtn2.Checked)
            {
                // 製造、間接部門別集計処理
                zanSumBumon(Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text), zYY, zMM);
            }
            else if (rBtn3.Checked)
            {
                // 全社集計処理
                zanSumAll(Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text), zYY, zMM);
            }
            else if (rBtn4.Checked)
            {
                // 課集計処理
                zanSum_KA(Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text), zYY, zMM);
            }
            else if (rBtn6.Checked)
            {
                // 係集計処理：2018/04/09
                zanSum_KAKARI(Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text), zYY, zMM);
            }
            else if (rBtn5.Checked)
            {
                // 班集計処理：2018/04/09
                zanSum_HAN(Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text), zYY, zMM);
            }

            MessageBox.Show("処理が終了しました", "残業推移グラフ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            
            // カーソル戻す
            this.Cursor = Cursors.Default;
        }


        private void frmZanChartXls_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //// 部署名コンボボックスのデータソースをセットする
            //Utility.ComboBumon.loadKa(comboBox2, _dbName);

            label1.Visible = false;
            toolStripProgressBar1.Visible = false;

            txtYear.Text = DateTime.Today.Year.ToString();
            txtMonth.Text = DateTime.Today.Month.ToString();

            rBtn4.Checked = true;
            comboBox2.Enabled = true;
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
            setZengetsuZan(ref zengetsuArray, 3);

            // 当月データ取得
            adp_ka.Fill(dts.残業集計1, yy, mm, yy, mm, yy, mm, yy, mm, yy, mm, yy, mm);

            // nullに「０」をセット
            foreach (var item in dts.残業集計1)
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

                // 部署コード
                string bmnCode = string.Empty;

                // 部署別残業理由シートの内容を配列に取得する
                object[,] zReSeizou = bs.getZanReason();

                // 部署別理由別残業計画シートの内容を配列に取得する
                //object[,] zRe = bs.getZanReasonPlan();
                bs.zrpArray = bs.getZanReasonPlan();

                // 部署別理由別残業計画シート(bs.zrpArray)の内容をCSV(kaByreByZanPlan.csv)に出力する : 2018/04/11
                kaByreByZanPlanSet();

                /* 部署別理由別残業計画.csv(kaByreByZanPlan.csv)から
                   課毎に残業計画時間を集計した配列を作成する : 2018/04/11 */
                kaByreByZanPlanSum(3);

                // 部署別残業計画シートの内容を配列に取得する
                bs.zpArray = bs.getZanPlan();

                // 部署別残業計画シート(bs.zpArray)の内容をCSV(kaZanPlan.csv)に出力する : 2017/10/18
                kaZanPlanSet();

                // 課別年月別に集計した人数・残業計画・生産数の2次元配列を生成する : 2018/04/09
                // zpArrayNew[]配列の生成
                kaZanPlanSum(3);

                // progressBar
                int nMax = 1;
                toolStripProgressBar1.Maximum = nMax;
                toolStripProgressBar1.Minimum = 0;
                toolStripProgressBar1.Visible = true;

                // 一番最近の勤務票日付
                int maxDay = days;

                // progressBar表示
                label1.Visible = true;
                label1.Text = "残業推移グラフ作成中です...";
                toolStripProgressBar1.Value = 1;
                //this.Refresh(); // ← 追加
                
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
                double zanPlan = getZanPlanBumon(zpArrayNew, yy, mm, Utility.StrtoInt(bmnCode));

                // 日付別の配列を生成
                clsZanSum[] z = new clsZanSum[1];
                dayArrayNew(yy, mm, bmnCode, ref z);

                // 日別残業時間を配列にセットする・残業月間合計を取得する
                decimal zanTotal = 0;
                setDaybyZanBumon(bmnCode, ref z, ref zanTotal);

                // 残業計画時間の稼働日数割りと日々目標値を配列にセットする : 2018/01/17
                setDaybyPlan(ref z, zanPlan);

                // 残業時間の実績累積を配列にセットする
                setDaybyZisseki(ref z, yy, mm, maxDay);

                //// 月間残業合計を時間単位に変換
                //zanTotal = Utility.StrtoDouble(((zanTotal / 60).ToString("#,##0.0")));

                // エクセルシート出力
                // テンプレートシートを追加する
                pCnt++;
                oxlsSheet.Copy(Type.Missing, oXlsBook.Sheets[pCnt - 1]);
                oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt];

                // シートにデータを貼り付ける
                xlsOutPutBumon(oXls, ref oXlsBook, ref oxlsSheet, z, Properties.Settings.Default.xlsZanChart, yy, mm, zReSeizou, szByreByzanPlan, bmnCode);
             
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
            setZengetsuZan(ref zengetsuArray, 3);

            // 2017/10/10
            // 当月データ取得　：理由別実績取得用
            adp_ka.Fill(dts.残業集計1, yy, mm, yy, mm, yy, mm, yy, mm, yy, mm, yy, mm);

            // nullに「０」をセット
            foreach (var item in dts.残業集計1)
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

                // 部署コード
                string bmnCode = string.Empty;

                // 部署別残業理由シートの内容を配列に取得する
                object[,] zReSeizou = bs.getZanReason();

                // 部署別理由別残業計画シートの内容を配列に取得する
                //object[,] zRe = bs.getZanReasonPlan();
                bs.zrpArray = bs.getZanReasonPlan();

                // 部署別理由別残業計画シート(bs.zrpArray)の内容をCSV(kaByreByZanPlan.csv)に出力する : 2018/04/18
                kaByreByZanPlanSet();

                /* 部署別理由別残業計画.csv(kaByreByZanPlan.csv)から
                   課毎に残業計画時間を集計した配列を作成する : 2018/04/18 */
                kaByreByZanPlanSum(3);

                // 部署別残業計画シートの内容を配列に取得する
                bs.zpArray = bs.getZanPlan();

                // 部署別残業計画シート(bs.zpArray)の内容をCSV(kaZanPlan.csv)に出力する : 2017/10/18
                kaZanPlanSet();

                // 課別年月別に集計した人数・残業計画・生産数の2次元配列を生成する : 2018/04/09
                // zpArrayNew[]配列の生成
                kaZanPlanSum(3);

                // progressBar
                int nMax = 2;
                toolStripProgressBar1.Maximum = nMax;
                toolStripProgressBar1.Minimum = 0;
                toolStripProgressBar1.Visible = true;

                // 一番最近の勤務票日付
                int maxDay = days;

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
                    double zanPlan = getZanPlanBumon(zpArrayNew, yy, mm, Utility.StrtoInt(bmnCode));

                    // 日付別の配列を生成
                    clsZanSum[] z = new clsZanSum[1];
                    dayArrayNew(yy, mm, bmnCode, ref z);

                    // 日別残業時間を配列にセットする・残業月間合計を取得する
                    decimal zanTotal = 0;
                    setDaybyZanBumon(bmnCode, ref z, ref zanTotal);

                    // 残業計画時間の稼働日数割りと日々目標値を配列にセットする : 2018/01/17
                    setDaybyPlan(ref z, zanPlan);

                    // 残業時間の実績累積を配列にセットする
                    setDaybyZisseki(ref z, yy, mm, maxDay);

                    //// 月間残業合計を時間単位に変換
                    //zanTotal = Utility.StrtoDouble(((zanTotal / 60).ToString("#,##0.0")));

                    // エクセルシート出力
                    // テンプレートシートを追加する
                    pCnt++;
                    oxlsSheet.Copy(Type.Missing, oXlsBook.Sheets[pCnt - 1]);
                    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt];

                    // シートにデータを貼り付ける
                    xlsOutPutBumon(oXls, ref oXlsBook, ref oxlsSheet, z, Properties.Settings.Default.xlsZanChart, yy, mm, zReSeizou, szByreByzanPlan, bmnCode);
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
            //setZengetsuZan(ref zengetsuArray);

            // !!!!!!!!!!!! デバッグ用前月データがないので当月で動作確認。必ず戻すこと !!!!!!!!!!!!
            setZengetsuZan(ref zengetsuArray, 3);
            // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

             //2017/10/10
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
                //int maxDay = dts.過去勤務票ヘッダ.Where(a => a.年 == yy && a.月 == mm).Max(a => a.日);
                int maxDay = days;

                // 部署別残業計画配列を順番に読む
                for (int i = 1; i <= nMax; i++)
                {
                    // progressBar表示
                    label1.Visible = true;
                    label1.Text = "残業推移グラフ作成中..." + i + "/" + nMax;
                    toolStripProgressBar1.Value = i;
                    this.Refresh(); // ← 追加

                    //// 部署コードが5桁未満のときは対象外
                    //if (bs.zpArray[i, 1].ToString().Length < 5)
                    //{
                    //    continue;
                    //}

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
                    decimal zanTotal = 0;
                    setDaybyZan(bushoCode, ref z, ref zanTotal);

                    // 残業計画時間の稼働日数割りと日々目標値を配列にセットする
                    setDaybyPlan(ref z, zanPlan);

                    // 残業時間の実績累積を配列にセットする
                    setDaybyZisseki(ref z, yy, mm, maxDay);

                    //// 月間残業合計を時間単位に変換
                    //zanTotal = Utility.StrtoDouble(((zanTotal / 60).ToString("#,##0.0")));

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

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     班別メイン集計処理 : 2018/04/16</summary>
        /// <param name="yy">
        ///     対象年</param>
        /// <param name="mm">
        ///     対象月</param>
        /// <param name="zYY">
        ///     前月の年</param>
        /// <param name="zMM">
        ///     前月</param>
        ///-----------------------------------------------------------------------
        private void zanSum_HAN(int yy, int mm, int zYY, int zMM)
        {
            // 前月データから前月実績配列を作成
            string[,] zengetsuArray = null;
            setZengetsuZan(ref zengetsuArray, 5);

            // 当月データ取得：理由別残業時間用として
            adp_Han.Fill(dts.残業集計3, yy, mm, yy, mm, yy, mm, yy, mm, yy, mm, yy, mm);

            // nullに「０」をセット
            foreach (var item in dts.残業集計3)
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

                // 部署別理由別残業計画シートの内容を配列に取得する
                //object[,] zRe = bs.getZanReasonPlan();
                bs.zrpArray = bs.getZanReasonPlan();
                
                // 部署別理由別残業計画シート(bs.zrpArray)の内容をCSV(kaByreByZanPlan.csv)に出力する : 2018/04/11
                kaByreByZanPlanSet();

                /* 部署別理由別残業計画.csv(kaByreByZanPlan.csv)から
                   班、係、課毎に残業計画時間を集計した配列を作成する : 2018/04/11 */
                kaByreByZanPlanSum(5);

                // 部署別残業計画シートの内容を配列に取得する
                bs.zpArray = bs.getZanPlan();

                // 部署別残業計画シート(bs.zpArray)の内容をCSV(kaZanPlan.csv)に出力する : 2017/10/18
                kaZanPlanSet();

                // 班別年月別に集計した人数・残業計画・生産数の2次元配列を生成する : 2018/04/09
                // zpArrayNew[]配列の生成
                kaZanPlanSum(5);

                // progressBar
                int nMax = zpArrayNew.GetLength(0);
                toolStripProgressBar1.Maximum = nMax;
                toolStripProgressBar1.Minimum = 0;
                toolStripProgressBar1.Visible = true;

                label1.Visible = true;
                //label1.Text = "残業推移グラフ作成中です...";

                // 一番最近の勤務票日付
                int maxDay = days;

                // 係別残業実績配列（zpArrayNew[]）を読む
                for (int i = 0; i < nMax; i++)
                {
                    // progressBar表示
                    label1.Text = "残業推移グラフ作成中..." + i + "/" + nMax;
                    toolStripProgressBar1.Value = i;
                    //this.Refresh(); // ← 追加

                    //// 部署コードが5桁未満のときは対象外
                    //if (bs.zpArray[i, 1].ToString().Length < 5)
                    //{
                    //    continue;
                    //}

                    // 対象年月以外のとき
                    if (Utility.StrtoInt(zpArrayNew[i, 1].ToString()) != yy || Utility.StrtoInt(zpArrayNew[i, 2].ToString()) != mm)
                    {
                        continue;
                    }

                    // 人員数が「０」のときは対象外
                    if (Utility.StrtoInt(zpArrayNew[i, 3].ToString()) == global.flgOff)
                    {
                        continue;
                    }

                    //// 課以上の階層は対象外 : 2018/04/09
                    //if (zpArrayNew[i, 0].ToString().Substring(3, 1) == global.FLGOFF)
                    //{
                    //    continue;
                    //}

                    // 班指定のとき該当するか
                    if (comboBox2.SelectedIndex != -1)
                    {
                        Utility.ComboBumon cmb = (Utility.ComboBumon)comboBox2.SelectedItem;
                        if (zpArrayNew[i, 0].ToString() != cmb.code.Substring(0, 5))
                        {
                            continue;
                        }
                    }

                    bushoCode = zpArrayNew[i, 0].ToString();                        // 係コードを取得
                    sNin = Utility.StrtoInt(zpArrayNew[i, 3].ToString());           // 人数取得
                    sSeisan = Utility.StrtoInt(zpArrayNew[i, 5].ToString());        // 生産数取得

                    // 前月の残業合計、生産数、人数を取得
                    zenSeisanNin(zengetsuArray, bushoCode, zYY, zMM, ref zSeisan, ref zNin, ref zenZan, ref zenKaDays);

                    // 該当係の当月計画値を取得する
                    double zanPlan = getZanPlan(zpArrayNew, yy, mm, bushoCode);

                    // 日付別の配列を生成
                    clsZanSum[] z = new clsZanSum[1];
                    dayArrayNew(yy, mm, bushoCode, ref z);

                    // 日別残業時間を配列にセットする・残業月間合計を取得する
                    decimal zanTotal = 0;
                    setDaybyZan(bushoCode, ref z, ref zanTotal);

                    // 残業計画時間の稼働日数割りと日々目標値を配列にセットする : 2018/01/17
                    setDaybyPlan(ref z, zanPlan);

                    // 残業時間の実績累積を配列にセットする
                    setDaybyZisseki(ref z, yy, mm, maxDay);

                    //// 月間残業合計を時間単位に変換
                    //zanTotal = Utility.StrtoDouble(((zanTotal / 60).ToString("#,##0.0")));

                    // エクセルシート出力
                    // テンプレートシートを追加する
                    pCnt++;
                    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];
                    oxlsSheet.Copy(Type.Missing, oXlsBook.Sheets[pCnt - 1]);
                    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt];

                    // シートにデータを貼り付ける
                    //xlsOutPut_KAKARI(oXls, ref oXlsBook, ref oxlsSheet, z, Properties.Settings.Default.xlsZanChart, yy, mm, zReSeizou, zRe, sdCon);
                    xlsOutPut_HAN(oXls, ref oXlsBook, ref oxlsSheet, z, Properties.Settings.Default.xlsZanChart, yy, mm, zReSeizou, szByreByzanPlan, sdCon);
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

                // 課指定のときファイル名に課名称を付加する
                string hKa = string.Empty;
                if (comboBox2.SelectedIndex != -1)
                {
                    Utility.ComboBumon cmb = (Utility.ComboBumon)comboBox2.SelectedItem;
                    hKa = cmb.Name + " ";
                }

                // ファイル名
                saveFileDialog1.FileName = yy + "年" + mm.ToString().PadLeft(2, ' ') + "月 " + hKa + "残業推移グラフ";
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

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     係別メイン集計処理 : 2018/04/09</summary>
        /// <param name="yy">
        ///     対象年</param>
        /// <param name="mm">
        ///     対象月</param>
        /// <param name="zYY">
        ///     前月の年</param>
        /// <param name="zMM">
        ///     前月</param>
        ///-----------------------------------------------------------------------
        private void zanSum_KAKARI(int yy, int mm, int zYY, int zMM)
        {
            // 前月データから前月実績配列を作成
            string[,] zengetsuArray = null;
            setZengetsuZan(ref zengetsuArray, 4);

            // 当月データ取得：理由別残業時間用として
            adp_kakari.Fill(dts.残業集計2, yy, mm, yy, mm, yy, mm, yy, mm, yy, mm, yy, mm);

            // nullに「０」をセット
            foreach (var item in dts.残業集計2)
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

                // 部署別理由別残業計画シートの内容を配列に取得する
                //object[,] zRe = bs.getZanReasonPlan();
                bs.zrpArray = bs.getZanReasonPlan();
                
                // 部署別理由別残業計画シート(bs.zrpArray)の内容をCSV(kaByreByZanPlan.csv)に出力する : 2018/04/11
                kaByreByZanPlanSet();

                /* 部署別理由別残業計画.csv(kaByreByZanPlan.csv)から
                   班、係、課毎に残業計画時間を集計した配列を作成する : 2018/04/11 */
                kaByreByZanPlanSum(4);
                
                // 部署別残業計画シートの内容を配列に取得する
                bs.zpArray = bs.getZanPlan();

                // 部署別残業計画シート(bs.zpArray)の内容をCSV(kaZanPlan.csv)に出力する : 2017/10/18
                kaZanPlanSet();

                // 係別年月別に集計した人数・残業計画・生産数の2次元配列を生成する : 2018/04/09
                // zpArrayNew[]配列の生成
                kaZanPlanSum(4);

                // progressBar
                int nMax = zpArrayNew.GetLength(0);
                toolStripProgressBar1.Maximum = nMax;
                toolStripProgressBar1.Minimum = 0;
                toolStripProgressBar1.Visible = true;

                label1.Visible = true;
                //label1.Text = "残業推移グラフ作成中です...";

                // 一番最近の勤務票日付
                int maxDay = days;

                // 係別残業実績配列（zpArrayNew[]）を読む
                for (int i = 0; i < nMax; i++)
                {
                    // progressBar表示
                    label1.Text = "残業推移グラフ作成中..." + i + "/" + nMax;
                    toolStripProgressBar1.Value = i;
                    //this.Refresh(); // ← 追加

                    //// 部署コードが5桁未満のときは対象外
                    //if (bs.zpArray[i, 1].ToString().Length < 5)
                    //{
                    //    continue;
                    //}

                    // 対象年月以外のとき
                    if (Utility.StrtoInt(zpArrayNew[i, 1].ToString()) != yy || Utility.StrtoInt(zpArrayNew[i, 2].ToString()) != mm)
                    {
                        continue;
                    }

                    // 人員数が「０」のときは対象外
                    if (Utility.StrtoInt(zpArrayNew[i, 3].ToString()) == global.flgOff)
                    {
                        continue;
                    }

                    //// 課以上の階層は対象外 : 2018/04/09
                    //if (zpArrayNew[i, 0].ToString().Substring(3, 1) == global.FLGOFF)
                    //{
                    //    continue;
                    //}

                    // 係指定のとき該当するか
                    if (comboBox2.SelectedIndex != -1)
                    {
                        Utility.ComboBumon cmb = (Utility.ComboBumon)comboBox2.SelectedItem;
                        if (zpArrayNew[i, 0].ToString() != cmb.code.Substring(0, 4))
                        {
                            continue;
                        }
                    }

                    bushoCode = zpArrayNew[i, 0].ToString();                        // 係コードを取得
                    sNin = Utility.StrtoInt(zpArrayNew[i, 3].ToString());           // 人数取得
                    sSeisan = Utility.StrtoInt(zpArrayNew[i, 5].ToString());        // 生産数取得

                    // 前月の残業合計、生産数、人数を取得
                    zenSeisanNin(zengetsuArray, bushoCode, zYY, zMM, ref zSeisan, ref zNin, ref zenZan, ref zenKaDays);

                    // 該当係の当月計画値を取得する
                    double zanPlan = getZanPlan(zpArrayNew, yy, mm, bushoCode);

                    // 日付別の配列を生成
                    clsZanSum[] z = new clsZanSum[1];
                    dayArrayNew(yy, mm, bushoCode, ref z);

                    // 日別残業時間を配列にセットする・残業月間合計を取得する
                    decimal zanTotal = 0;
                    setDaybyZan(bushoCode + "0", ref z, ref zanTotal);

                    // 残業計画時間の稼働日数割りと日々目標値を配列にセットする : 2018/01/17
                    setDaybyPlan(ref z, zanPlan);

                    // 残業時間の実績累積を配列にセットする
                    setDaybyZisseki(ref z, yy, mm, maxDay);

                    //// 月間残業合計を時間単位に変換
                    //zanTotal = Utility.StrtoDouble(((zanTotal / 60).ToString("#,##0.0")));

                    // エクセルシート出力
                    // テンプレートシートを追加する
                    pCnt++;
                    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];
                    oxlsSheet.Copy(Type.Missing, oXlsBook.Sheets[pCnt - 1]);
                    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt];

                    // シートにデータを貼り付ける
                    //xlsOutPut_KAKARI(oXls, ref oXlsBook, ref oxlsSheet, z, Properties.Settings.Default.xlsZanChart, yy, mm, zReSeizou, zRe, sdCon);
                    xlsOutPut_KAKARI(oXls, ref oXlsBook, ref oxlsSheet, z, Properties.Settings.Default.xlsZanChart, yy, mm, zReSeizou, szByreByzanPlan, sdCon);
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

                // 課指定のときファイル名に課名称を付加する
                string hKa = string.Empty;
                if (comboBox2.SelectedIndex != -1)
                {
                    Utility.ComboBumon cmb = (Utility.ComboBumon)comboBox2.SelectedItem;
                    hKa = cmb.Name + " ";
                }

                // ファイル名
                saveFileDialog1.FileName = yy + "年" + mm.ToString().PadLeft(2, ' ') + "月 " + hKa + "残業推移グラフ";
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

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     課別メイン集計処理 </summary>
        /// <param name="yy">
        ///     対象年</param>
        /// <param name="mm">
        ///     対象月</param>
        /// <param name="zYY">
        ///     前月の年</param>
        /// <param name="zMM">
        ///     前月</param>
        ///-----------------------------------------------------------------------
        private void zanSum_KA(int yy, int mm, int zYY, int zMM)
        {
            // 前月データから前月実績配列を作成
            string[,] zengetsuArray = null;
            setZengetsuZan(ref zengetsuArray, 3);

            //2017/10/10
            // 当月データ取得：理由別残業時間用として
            adp_ka.Fill(dts.残業集計1, yy, mm, yy, mm, yy, mm, yy, mm, yy, mm, yy, mm);

            // nullに「０」をセット
            foreach (var item in dts.残業集計1)
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

                // 部署別理由別残業計画シートの内容を配列に取得する
                //object[,] zRe = bs.getZanReasonPlan();
                bs.zrpArray = bs.getZanReasonPlan();

                // 部署別理由別残業計画シート(bs.zrpArray)の内容をCSV(kaByreByZanPlan.csv)に出力する : 2018/04/11
                kaByreByZanPlanSet();

                /* 部署別理由別残業計画.csv(kaByreByZanPlan.csv)から
                   課毎に残業計画時間を集計した配列を作成する : 2018/04/11 */
                kaByreByZanPlanSum(3);

                // 部署別残業計画シートの内容を配列に取得する
                bs.zpArray = bs.getZanPlan();

                // 部署別残業計画シート(bs.zpArray)の内容をCSV(kaZanPlan.csv)に出力する : 2017/10/18
                kaZanPlanSet();

                // 課別年月別に集計した人数・残業計画・生産数の2次元配列を生成する : 2018/04/09
                // zpArrayNew[]配列の生成
                kaZanPlanSum(3);
                
                // progressBar
                int nMax = zpArrayNew.GetLength(0);
                toolStripProgressBar1.Maximum = nMax;
                toolStripProgressBar1.Minimum = 0;
                toolStripProgressBar1.Visible = true;

                label1.Visible = true;
                //label1.Text = "残業推移グラフ作成中です...";

                // 一番最近の勤務票日付
                int maxDay = days;

                // 部署別残業計画配列を順番に読む
                for (int i = 0; i < nMax; i++)
                {
                    // progressBar表示
                    label1.Text = "残業推移グラフ作成中..." + i + "/" + nMax;
                    toolStripProgressBar1.Value = i;
                    //this.Refresh(); // ← 追加

                    //// 部署コードが5桁未満のときは対象外
                    //if (bs.zpArray[i, 1].ToString().Length < 5)
                    //{
                    //    continue;
                    //}

                    // 対象年月以外のとき
                    if (Utility.StrtoInt(zpArrayNew[i, 1].ToString()) != yy || Utility.StrtoInt(zpArrayNew[i, 2].ToString()) != mm)
                    {
                        continue;
                    }

                    // 人員数が「０」のときは対象外
                    if (Utility.StrtoInt(zpArrayNew[i, 3].ToString()) == global.flgOff)
                    {
                        continue;
                    }

                    // 部署指定のとき該当するか
                    if (comboBox2.SelectedIndex != -1)
                    {
                        Utility.ComboBumon cmb = (Utility.ComboBumon)comboBox2.SelectedItem;
                        if (zpArrayNew[i, 0].ToString() != cmb.code.Substring(0, 3))
                        {
                            continue;
                        }
                    }

                    bushoCode = zpArrayNew[i, 0].ToString();                        // 課コードを取得
                    sNin = Utility.StrtoInt(zpArrayNew[i, 3].ToString());           // 人数取得
                    sSeisan = Utility.StrtoInt(zpArrayNew[i, 5].ToString());        // 生産数取得

                    // 前月の残業合計、生産数、人数を取得
                    zenSeisanNin(zengetsuArray, bushoCode, zYY, zMM, ref zSeisan, ref zNin, ref zenZan, ref zenKaDays);

                    // 該当課の当月計画値を取得する
                    double zanPlan = getZanPlan(zpArrayNew, yy, mm, bushoCode);

                    // 日付別の配列を生成
                    clsZanSum[] z = new clsZanSum[1];
                    dayArrayNew(yy, mm, bushoCode, ref z);

                    // 日別残業時間を配列にセットする・残業月間合計を取得する
                    decimal zanTotal = 0;
                    setDaybyZan(bushoCode + "00", ref z, ref zanTotal);

                    // 残業計画時間の稼働日数割りと日々目標値を配列にセットする : 2018/01/17
                    setDaybyPlan(ref z, zanPlan);

                    // 残業時間の実績累積を配列にセットする
                    setDaybyZisseki(ref z, yy, mm, maxDay);

                    //// 月間残業合計を時間単位に変換
                    //zanTotal = Utility.StrtoDouble(((zanTotal / 60).ToString("#,##0.0")));

                    // エクセルシート出力
                    // テンプレートシートを追加する
                    pCnt++;
                    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];
                    oxlsSheet.Copy(Type.Missing, oXlsBook.Sheets[pCnt - 1]);
                    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt];
                                        
                    // シートにデータを貼り付ける
                    xlsOutPut_KA(oXls, ref oXlsBook, ref oxlsSheet, z, Properties.Settings.Default.xlsZanChart, yy, mm, zReSeizou, szByreByzanPlan, sdCon);
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

                // 課指定のときファイル名に課名称を付加する
                string hKa = string.Empty;
                if (comboBox2.SelectedIndex != -1)
                {
                    Utility.ComboBumon cmb = (Utility.ComboBumon)comboBox2.SelectedItem;
                    hKa = cmb.Name + " ";
                }

                // ファイル名
                saveFileDialog1.FileName = yy + "年" + mm.ToString().PadLeft(2, ' ') + "月 " + hKa + "残業推移グラフ";
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

        private string[] setZreSummary(object [,] zRe, int yy, int mm)
        {
            int iX = 0;
            string[] zReArray = null; 
            string wkCode = string.Empty;
            int wkNum = 0;
            int[] keikaku = new int[10];
            for (int i = 0; i < keikaku.Length; i++)
            {
                keikaku[i] = 0;
            }

            for (int i = 1; i < zRe.GetLength(0); i++)
            {
                // 対象年月以外のとき
                if (Utility.StrtoInt(zRe[i, 2].ToString()) != yy || Utility.StrtoInt(zRe[i, 3].ToString()) != mm)
                {
                    continue;
                }

                if (wkCode != string.Empty && wkCode != zRe[i, 1].ToString().Substring(0, 3))
                {
                    for (int z = 0; z < 10; z++)
                    {
                        Array.Resize(ref zReArray, iX + 1);
                        zReArray[iX] = wkCode + "," + (z + 1) + "," + keikaku[z];
                        iX++;
                    }

                    for (int iz = 0; iz < keikaku.Length; iz++)
                    {
                        keikaku[iz] = 0;
                    }
                }

                wkCode = zRe[i, 1].ToString().Substring(0, 3);
                wkNum = Utility.StrtoInt(zRe[i, 4].ToString().Trim()) - 1;
                if (wkNum >= 10)
                {
                    wkNum = 9;
                }

                keikaku[wkNum] += Utility.StrtoInt(zRe[i, 6].ToString().Trim());
            }

            return zReArray;
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
            for (int i = 0; i < zpArrayNew.GetLength(0); i++)
            {
                if (zpArrayNew[i, 0].ToString() == bCode && 
                    Utility.StrtoInt(zpArrayNew[i, 1].ToString()) == zyy && 
                    Utility.StrtoInt(zpArrayNew[i, 2].ToString()) == zmm)
                {
                    // 生産数、人数を取得
                    zNi = Utility.StrtoInt(zpArrayNew[i, 3].ToString());
                    zSei = Utility.StrtoInt(zpArrayNew[i, 5].ToString());
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

            //// 月間残業合計を時間単位に変換
            //zz = Utility.StrtoDouble(((zz / 60).ToString("#,##0.0")));

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
            for (int i = 0; i < zpArrayNew.GetLength(0); i++)
            {
                // 人数が0のときはネグる
                if (Utility.StrtoInt(zpArrayNew[i, 3].ToString()) == 0)
                {
                    continue;
                }

                if (bmnCode == "0")
                {
                    // 全社
                    if (Utility.StrtoInt(zpArrayNew[i, 1].ToString()) == zyy &&
                        Utility.StrtoInt(zpArrayNew[i, 2].ToString()) == zmm)
                    {
                        // 生産数、人数を取得
                        zNi += Utility.StrtoInt(zpArrayNew[i, 3].ToString());
                        zSei += Utility.StrtoInt(zpArrayNew[i, 5].ToString());
                    }
                }
                else if (bmnCode == "1")
                {
                    // 直接部門
                    if (zpArrayNew[i, 0].ToString().Substring(1, 1) == bmnCode &&
                        Utility.StrtoInt(zpArrayNew[i, 1].ToString()) == zyy &&
                        Utility.StrtoInt(zpArrayNew[i, 2].ToString()) == zmm)
                    {
                        // 生産数、人数を取得
                        zNi += Utility.StrtoInt(zpArrayNew[i, 3].ToString());
                        zSei += Utility.StrtoInt(zpArrayNew[i, 5].ToString());
                    }
                }
                else if (bmnCode == "2")
                {
                    // 間接部門
                    if (zpArrayNew[i, 0].ToString().Substring(1, 1) != "0" &&
                        zpArrayNew[i, 0].ToString().Substring(1, 1) != "1" && 
                        Utility.StrtoInt(zpArrayNew[i, 1].ToString()) == zyy &&
                        Utility.StrtoInt(zpArrayNew[i, 2].ToString()) == zmm)
                    {
                        // 生産数、人数を取得
                        zNi += Utility.StrtoInt(zpArrayNew[i, 3].ToString());
                        zSei += Utility.StrtoInt(zpArrayNew[i, 5].ToString());
                    }
                }
            }

            // 前月残業合計を取得
            zz = 0;
            for (int i = 0; i < zen.GetLength(0); i++)
            {
                if (bmnCode == "0")
                {
                    // 全社
                    if (zen[i, 0].Substring(1, 1) != "0")
                    {
                        zz += Utility.StrtoDouble(zen[i, 1]);
                    }
                }
                else if (bmnCode == "1")
                {
                    // 製造部門
                    if (zen[i, 0].Substring(1, 1) == bmnCode)
                    {
                        zz += Utility.StrtoDouble(zen[i, 1]);
                    }
                }
                else if (bmnCode == "2")
                {
                    // 間接部門
                    if (zen[i, 0].Substring(1, 1) != "0" && zen[i, 0].Substring(1, 1) != "1")
                    {
                        zz += Utility.StrtoDouble(zen[i, 1]);
                    }
                }
            }

            //// 月間残業合計を時間単位に変換
            //zz = Utility.StrtoDouble(((zz / 60).ToString("#,##0.0")));

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
                oxlsSheet.Name = szCode + oxlsSheet.Cells[1, 13].value;

                // 実稼働日数を求める
                int d = zd.Where(a => a.sHoliday == 0 || a.sZangyo > 0).Count();
                oxlsSheet.Cells[1, 21] = d.ToString();

                // 集計期間
                int maxDay = days;
                oxlsSheet.Cells[1, 24] = mm.ToString().PadLeft(2, ' ') + "月" + maxDay.ToString().PadLeft(2, ' ') + "日まで";
                
                //if (dts.過去勤務票ヘッダ.Any(a => a.年 == yy && a.月 == mm && a.部署コード == szCode))
                //{
                //    maxDay = dts.過去勤務票ヘッダ.Where(a => a.年 == yy && a.月 == mm && a.部署コード == szCode).Max(a => a.日);
                //    oxlsSheet.Cells[1, 24] = mm.ToString().PadLeft(2, ' ') + "月" + maxDay.ToString().PadLeft(2, ' ') + "日まで";
                //}
                //else
                //{
                //    maxDay = 0;
                //    oxlsSheet.Cells[1, 24] = "勤怠データなし";
                //}

                //// 集計日数
                //int n = zd.Count(a => (a.sHoliday == 0 || a.sZangyo > 0) && a.sDay <= maxDay);
                //oxlsSheet.Cells[1, 28] = n;


                // 集計日数：該当月の勤務日数 2018/01/12
                int n = zd.Count(a => a.sHoliday == 0);
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
                if (Utility.StrtoInt(szCode.Substring(1, 1)) <= global.flgOn) // 製造部門
                {
                    // 製造部門：理由別残業計画はなし
                    //zReSeizou = bs.getZanReason();      // 残業理由配列

                    for (int i = 1; i <= zReSeizou.GetLength(0); i++)
                    {
                        // 部署コードが一致しているか？
                        if (zReSeizou[i, 1].ToString() != szCode.ToString())
                        {
                            continue;
                        }

                        oxlsSheet.Cells[5 + iR, 20] = zReSeizou[i, 2];    // 理由コード
                        oxlsSheet.Cells[5 + iR, 21] = zReSeizou[i, 3];    // 残業理由

                        // 理由別残業時間を集計：最大日付範囲で取得 2017/11/22
                        double s = dts.残業集計1
                            .Where(a => a.部署コード == szCode && a.残業理由 == (double)zReSeizou[i, 2] && 
                                        a.日 <= maxDay)
                            .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));

                        s = s / 60;
                        oxlsSheet.Cells[5 + iR, 27] = s;

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
                    //zRe = bs.getZanReasonPlan();

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

                        // 理由別残業時間を集計：最大日付範囲で取得 2017/11/22
                        double s = dts.残業集計1
                            .Where(a => a.部署コード == szCode && a.残業理由 == (double)zRe[i, 4] && 
                                        a.日 <= maxDay)
                            .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));

                        s = s / 60;
                        oxlsSheet.Cells[5 + iR, 27] = s;
                        
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
                oxlsSheet.Cells[35, 23] = zSeisan;                  // 前月生産数
                oxlsSheet.Cells[37, 23] = zNin;                     // 前月人数
                //oxlsSheet.Cells[44, 23] = zenZan;                 // 前月残業合計　※チェック用
                //oxlsSheet.Cells[41, 23] = zenZan / zNin;          // 前月1人当り残業
                oxlsSheet.Cells[41, 23] = zenZan / zNin / zenKaDays;    // 前月1人当り残業（実績 / 人数 / 稼働日数）  : 2018/01/08
                //oxlsSheet.Cells[45, 23] = zenKaDays;                  // 前月稼働日数　※チェック用
                oxlsSheet.Cells[39, 23] = zenZan / zenKaDays;           // 前月一日当たり合計

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
        ///     エクセルシートへ出力：班別 2018/04/16 </summary>
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
        private void xlsOutPut_HAN(Excel.Application oXls, ref Excel.Workbook oXlsBook, ref Excel.Worksheet oxlsSheet, clsZanSum[] zd, string xlsFile, int yy, int mm, object[,] zReSeizou, object[] zRe, sqlControl.DataControl sdCon)
        {
            oxlsSheet.Select(Type.Missing);

            Excel.Range rng = null;

            int iX = 1;

            // 実稼働日数を取得
            //int kDays = zd.Where(a => a.sHoliday == 0 || a.sZangyo > 0).Count(); 2018/01/17
            int kDays = zd.Count(a => a.sHoliday == 0); // 2018/01/17 休日出勤は含まない

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
            string szCode = string.Empty;

            try
            {
                // ウィンドウを非表示にする
                oXls.Visible = false;
                oXls.DisplayAlerts = false;

                oxlsSheet.Cells[1, 2] = yy + "年" + mm.ToString().PadLeft(2, ' ') + "月 残業推移グラフ";

                // 部署コード、名称
                szCode = zd.First().sSzCode.ToString();
                oxlsSheet.Cells[1, 11] = szCode;
                oxlsSheet.Cells[1, 13] = getDepartmentName(_dbName, szCode, sdCon);

                // シート名に部署名をつける
                oxlsSheet.Name = szCode + oxlsSheet.Cells[1, 13].value;

                //// 実稼働日数を求める
                //int d = zd.Where(a => a.sHoliday == 0 || a.sZangyo > 0).Count();
                //oxlsSheet.Cells[1, 21] = d.ToString();

                // 稼働日数を求める:休日出勤を含まない集計日までの稼働日数 2018/01/08
                int d = zd.Where(a => a.sHoliday == 0 && a.sDay <= days).Count();
                //oxlsSheet.Cells[1, 21] = d.ToString(); 2018/01/17
                oxlsSheet.Cells[1, 28] = d.ToString();  // 2018/01/17

                // 集計期間
                int maxDay = days;
                oxlsSheet.Cells[1, 24] = mm.ToString().PadLeft(2, ' ') + "月" + maxDay.ToString().PadLeft(2, ' ') + "日まで";

                // 集計日数：該当月の勤務日数 2018/01/12
                //int n = zd.Count(a => a.sHoliday == 0); 2018/01/17
                //oxlsSheet.Cells[1, 28] = n;   2018/01/17
                oxlsSheet.Cells[1, 21] = kDays;

                // グラフ用データ一括書き込み
                rng = oxlsSheet.Range[oxlsSheet.Cells[3, 31], oxlsSheet.Cells[7, 61]];
                rng.Value2 = xlsArray;

                iX = 0;

                int iR = 0;
                double zRePlanTl = 0;

                // 製造部門と間接部門の処理共通化のため以下、コメント化 2018/06/23

                //// 製造部門
                //if (Utility.StrtoInt(szCode.Substring(1, 1)) <= global.flgOn)
                //{
                //    for (int i = 1; i <= zReSeizou.GetLength(0); i++)
                //    {
                //        // 有効な部署コードでないときはネグる：2018/04/12
                //        if (Utility.NulltoStr(zReSeizou[i, 1]).Length < 5)
                //        {
                //            continue;
                //        }

                //        // 部署コード５桁が一致しているか？
                //        if (zReSeizou[i, 1].ToString().Substring(0, 5) != szCode.ToString())
                //        {
                //            continue;
                //        }

                //        oxlsSheet.Cells[5 + iR, 20] = zReSeizou[i, 2];    // 理由コード
                //        oxlsSheet.Cells[5 + iR, 21] = zReSeizou[i, 3];    // 残業理由

                //        // 理由別残業時間を集計 : 最大日付範囲で取得 2017/11/22
                //        double s = 0;

                //        // 応援先で集計 2018/02/11
                //        if ((double)zReSeizou[i, 2] < 10)
                //        {
                //            s = dts.残業集計3
                //                .Where(a => a.応援先 == szCode &&
                //                       a.残業理由 == (double)zReSeizou[i, 2] &&
                //                       a.日 <= maxDay)
                //                .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));
                //        }
                //        else
                //        {
                //            s = dts.残業集計3
                //                .Where(a => a.応援先 == szCode &&
                //                       a.残業理由 >= 10 &&
                //                       a.日 <= maxDay)
                //                .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));
                //        }

                //        s = s / 60;
                //        oxlsSheet.Cells[5 + iR, 27] = s;

                //        iR += 2;

                //        if (iR == 20)
                //        {
                //            // 2018/04/17
                //            //// 合計欄に当月計画時間の稼働日数割りを表示する
                //            //if (zd.Any(a => a.sDay == maxDay))
                //            //{
                //            //    foreach (var t in zd.Where(a => a.sDay == maxDay))
                //            //    {
                //            //        oxlsSheet.Cells[27, 25] = t.sMonthPlan;
                //            //        break;
                //            //    }
                //            //}
                //            //else
                //            //{
                //            //    oxlsSheet.Cells[27, 25] = 0;
                //            //}

                //            break;
                //        }
                //    }

                //    // 合計欄に当月計画時間の稼働日数割りを表示する
                //    //if (zd.Any(a => a.sDay == maxDay))
                //    //{
                //    //    foreach (var t in zd.Where(a => a.sDay == maxDay))
                //    //    {
                //    //        oxlsSheet.Cells[27, 25] = t.sMonthPlan;
                //    //        break;
                //    //    }
                //    //}
                //    //else
                //    //{
                //    //    oxlsSheet.Cells[27, 25] = 0;
                //    //}

                //    // 当月計画時間初期化 2018/04/17
                //    double pp = 0;

                //    // 合計欄に当月計画時間の稼働日数割りを表示する 2018/04/17
                //    foreach (var t in zd.Where(a => a.sDay <= maxDay))
                //    {
                //        // ゼロ以外の最大日付の計画値
                //        if (t.sMonthPlan != 0)
                //        {
                //            pp = t.sMonthPlan;
                //        }
                //    }

                //    oxlsSheet.Cells[27, 25] = pp;

                //    // 当月残業合計
                //    oxlsSheet.Cells[27, 27] = setTotalZan(szCode, outZangyoFile);
                //}
                //else
                //{
                //    // 間接部門
                //    for (int i = 0; i < zRe.Length; i++)
                //    {
                //        // 班別理由別残業計画を取得 : 2018/04/11
                //        string[] dd = zRe[i].ToString().Split(',');

                //        if (dd.Length < 5)
                //        {
                //            continue;
                //        }

                //        // 班コードが一致しているか？
                //        if (dd[0].ToString() != szCode.ToString())
                //        {
                //            continue;
                //        }

                //        // 年月コードが一致しているか？
                //        if (dd[1].ToString() != yy.ToString() || dd[2].ToString() != mm.ToString())
                //        {
                //            continue;
                //        }

                //        oxlsSheet.Cells[5 + iR, 20] = dd[3];    // 理由コード

                //        // 残業理由名称を部署別残業理由配列より取得：2018/04/12
                //        for (int ii = 1; ii <= zReSeizou.GetLength(0); ii++)
                //        {
                //            // 有効な部署コードでないときはネグる：2018/04/12
                //            if (Utility.NulltoStr(zReSeizou[ii, 1]).Length < 5)
                //            {
                //                continue;
                //            }

                //            // 部署コード５桁が一致しているか？
                //            if (zReSeizou[ii, 1].ToString().Substring(0, 5) != szCode.ToString())
                //            {
                //                continue;
                //            }

                //            // 理由コードが一致しているか？
                //            if (zReSeizou[ii, 2].ToString() != dd[3])
                //            {
                //                continue;
                //            }

                //            oxlsSheet.Cells[5 + iR, 21] = zReSeizou[ii, 3];    // 残業理由
                //            break;
                //        }

                //        //oxlsSheet.Cells[5 + iR, 21] = zRe[i, 5];    // 残業理由
                //        //oxlsSheet.Cells[5 + iR, 25] = Utility.StrtoDouble(zRe[i, 6].ToString()) / (double)kDays * n;    // 現在の日付で稼働日割りした理由別計画値 2018/01/17
                //        //oxlsSheet.Cells[5 + iR, 25] = Utility.StrtoDouble(zRe[i, 6].ToString()) / (double)kDays * d;    // 現在の日付で稼働日割りした理由別計画値 2018/01/17
                //        oxlsSheet.Cells[5 + iR, 25] = Utility.StrtoDouble(dd[4].ToString()) / (double)kDays * d;    // 現在の日付で稼働日割りした理由別計画値 2018/04/11

                //        // 理由別残業時間を集計 : 最大日付範囲で取得 2017/11/22
                //        //double s = dts.残業集計1
                //        //    .Where(a => a.部署コード.Substring(0, 3) == szCode && a.残業理由 == (double)zRe[i, 4] && 
                //        //                a.日 <= maxDay)
                //        //    .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));

                //        /* 理由別残業時間を集計 : 最大日付範囲で取得 2017/11/22
                //                                 応援先を対象とする 2018/02/11 */
                //        double s = dts.残業集計3
                //            .Where(a => a.応援先 == szCode &&
                //                        a.残業理由 == Utility.StrtoDouble(dd[3]) &&
                //                        a.日 <= maxDay)
                //            .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));

                //        s = s / 60;
                //        oxlsSheet.Cells[5 + iR, 27] = s;

                //        iR += 2;
                //        //zRePlanTl += Utility.StrtoDouble(zRe[i, 6].ToString()); // 計画合計に加算
                //        //zRePlanTl += Utility.StrtoDouble(zRe[i, 6].ToString()) / (double)kDays * n;     // 計画合計に加算 2018/01/17
                //        //zRePlanTl += Utility.StrtoDouble(zRe[i, 6].ToString()) / (double)kDays * d;     // 計画合計に加算 2018/01/17
                //        zRePlanTl += Utility.StrtoDouble(dd[4].ToString()) / (double)kDays * d;     // 計画合計に加算 2018/04/11

                //        if (iR == 20)
                //        {
                //            //// 現在の日付で稼働日割りした理由別計画の合計
                //            //oxlsSheet.Cells[27, 25] = zRePlanTl;
                //            break;
                //        }
                //    }

                //    // 現在の日付で稼働日割りした理由別計画の合計
                //    oxlsSheet.Cells[27, 25] = zRePlanTl;

                //    // 当月残業合計
                //    oxlsSheet.Cells[27, 27] = setTotalZan(szCode, outZangyoFile);
                //}



                // ============= 製造部門も理由別残業計画を行うため製造部門と間接部門の処理を共通化 2018/06/23 ============= 
                
                for (int i = 0; i < zRe.Length; i++)
                {
                    // 班別理由別残業計画を取得 : 2018/04/11
                    string[] dd = zRe[i].ToString().Split(',');

                    if (dd.Length < 5)
                    {
                        continue;
                    }

                    // 班コードが一致しているか？
                    if (dd[0].ToString() != szCode.ToString())
                    {
                        continue;
                    }

                    // 年月コードが一致しているか？
                    if (dd[1].ToString() != yy.ToString() || dd[2].ToString() != mm.ToString())
                    {
                        continue;
                    }

                    oxlsSheet.Cells[5 + iR, 20] = dd[3];    // 理由コード

                    // 残業理由名称を部署別残業理由配列より取得：2018/04/12
                    for (int ii = 1; ii <= zReSeizou.GetLength(0); ii++)
                    {
                        // 有効な部署コードでないときはネグる：2018/04/12
                        if (Utility.NulltoStr(zReSeizou[ii, 1]).Length < 5)
                        {
                            continue;
                        }

                        // 部署コード５桁が一致しているか？
                        if (zReSeizou[ii, 1].ToString().Substring(0, 5) != szCode.ToString())
                        {
                            continue;
                        }

                        // 理由コードが一致しているか？
                        if (zReSeizou[ii, 2].ToString() != dd[3])
                        {
                            continue;
                        }

                        oxlsSheet.Cells[5 + iR, 21] = zReSeizou[ii, 3];    // 残業理由
                        break;
                    }

                    // 現在の日付で稼働日割りした理由別計画値 2018/04/11
                    oxlsSheet.Cells[5 + iR, 25] = Utility.StrtoDouble(dd[4].ToString()) / (double)kDays * d;
                    
                    /* 理由別残業時間を集計 : 最大日付範囲で取得 2017/11/22
                                             応援先を対象とする 2018/02/11 */
                    double s = dts.残業集計3
                        .Where(a => a.応援先 == szCode &&
                                    a.残業理由 == Utility.StrtoDouble(dd[3]) &&
                                    a.日 <= maxDay)
                        .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));

                    s = s / 60;
                    oxlsSheet.Cells[5 + iR, 27] = s;

                    iR += 2;
                    zRePlanTl += Utility.StrtoDouble(dd[4].ToString()) / (double)kDays * d;     // 計画合計に加算 2018/04/11

                    if (iR == 20)
                    {
                        //// 現在の日付で稼働日割りした理由別計画の合計
                        //oxlsSheet.Cells[27, 25] = zRePlanTl;
                        break;
                    }
                }

                // 現在の日付で稼働日割りした理由別計画の合計
                oxlsSheet.Cells[27, 25] = zRePlanTl;

                // 当月残業合計
                oxlsSheet.Cells[27, 27] = setTotalZan(szCode, outZangyoFile);

                //==================================  2018/06/23 ここまで =============================================================


                // 前月との比較要素のセット
                oxlsSheet.Cells[35, 23] = zSeisan;                      // 前月生産数
                oxlsSheet.Cells[37, 23] = zNin;                         // 前月人数
                oxlsSheet.Cells[41, 23] = zenZan / zNin / zenKaDays;    // 前月1人当り残業（実績 / 人数 / 稼働日数）  : 2018/01/20
                oxlsSheet.Cells[39, 23] = zenZan / zenKaDays;           // 前月一日当たり合計
                oxlsSheet.Cells[35, 25] = sSeisan;                      // 生産数
                oxlsSheet.Cells[37, 25] = sNin;                         // 人数

                //MessageBox.Show(msg);

                // シート名 2018/02/17
                string nm = "'" + oxlsSheet.Name + "'";

                // 2018/02/08 データ範囲(系列値）の可変表示
                Excel.ChartObject ss = (Excel.ChartObject)oxlsSheet.ChartObjects(1);

                string msg = string.Empty;
                for (int i = 1; i <= 4; i++)
                {
                    // =SERIES('151総務課'!$AD$4,'151総務課'!$AE$3:$BI$3,'151総務課'!$AE$4:$BI$4,1)
                    Excel.Series sr = ss.Chart.SeriesCollection(i);
                    msg += sr.Formula + Environment.NewLine;

                    // =SERIES('111第１組立課'!$AD$4,'111第１組立課'!day,'111第１組立課'!date,1)
                    string f = sr.Formula;
                    f = f.Replace("=SERIES(", "");
                    f = f.Replace(")", "");
                    string[] fa = f.Split(',');

                    // =SERIES('111第１組立課'!$AD$4,'111第１組立課'!day,'111第１組立課'!date,1)

                    // 系列名、系列値のシート名も直接セットする 2018/02/17
                    if (fa[3] == "1")
                    {
                        // 日付別
                        fa[0] = nm + "!$AD$4";      // 系列名
                        fa[1] = nm + "!day";        // 軸ラベルの範囲
                        fa[2] = nm + "!date";       // 系列値
                    }
                    else if (fa[3] == "2")
                    {
                        // 計画
                        fa[0] = nm + "!$AD$5";      // 系列名
                        fa[1] = nm + "!day";        // 軸ラベルの範囲
                        fa[2] = nm + "!plan";       // 系列値

                        // グラフの色を変更 2018/03/08
                        sr.Format.Line.ForeColor.RGB = System.Drawing.Color.DimGray.ToArgb();
                    }
                    else if (fa[3] == "3")
                    {
                        // 実績
                        fa[0] = nm + "!$AD$6";          // 系列名
                        fa[1] = nm + "!day";            // 軸ラベルの範囲
                        fa[2] = nm + "!Performance";    // 系列値
                    }
                    else if (fa[3] == "4")
                    {
                        // 目標
                        fa[0] = nm + "!$AD$7";      // 系列名
                        fa[1] = nm + "!day";        // 軸ラベルの範囲
                        fa[2] = nm + "!goal";       // 系列値
                    }

                    string newF = "=SERIES(" + fa[0] + "," + fa[1] + "," + fa[2] + "," + fa[3] + ")";
                    //msg += newF + Environment.NewLine;

                    sr.Formula = newF;
                }

                // シート全体が印刷対象とするためセルをアクティブとする 2018/02/09
                oxlsSheet.Range["A1"].Select();
            }
            catch (Exception ex)
            {
                MessageBox.Show(szCode + "," + ex.Message, "残業グラフエクセル出力", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
            }
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     エクセルシートへ出力：係別 2018/04/09 </summary>
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
        private void xlsOutPut_KAKARI(Excel.Application oXls, ref Excel.Workbook oXlsBook, ref Excel.Worksheet oxlsSheet, clsZanSum[] zd, string xlsFile, int yy, int mm, object[,] zReSeizou, object[] zRe, sqlControl.DataControl sdCon)
        {
            oxlsSheet.Select(Type.Missing);

            Excel.Range rng = null;

            int iX = 1;

            // 実稼働日数を取得
            //int kDays = zd.Where(a => a.sHoliday == 0 || a.sZangyo > 0).Count(); 2018/01/17
            int kDays = zd.Count(a => a.sHoliday == 0); // 2018/01/17 休日出勤は含まない

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
            string szCode = string.Empty;

            try
            {
                // ウィンドウを非表示にする
                oXls.Visible = false;
                oXls.DisplayAlerts = false;

                oxlsSheet.Cells[1, 2] = yy + "年" + mm.ToString().PadLeft(2, ' ') + "月 残業推移グラフ";

                // 部署コード、名称
                szCode = zd.First().sSzCode.ToString();
                oxlsSheet.Cells[1, 11] = szCode + "0";
                oxlsSheet.Cells[1, 13] = getDepartmentName(_dbName, szCode + "0", sdCon);

                // シート名に部署名をつける
                oxlsSheet.Name = szCode + oxlsSheet.Cells[1, 13].value;

                //// 実稼働日数を求める
                //int d = zd.Where(a => a.sHoliday == 0 || a.sZangyo > 0).Count();
                //oxlsSheet.Cells[1, 21] = d.ToString();

                // 稼働日数を求める:休日出勤を含まない集計日までの稼働日数 2018/01/08
                int d = zd.Where(a => a.sHoliday == 0 && a.sDay <= days).Count();
                //oxlsSheet.Cells[1, 21] = d.ToString(); 2018/01/17
                oxlsSheet.Cells[1, 28] = d.ToString();  // 2018/01/17

                // 集計期間
                int maxDay = days;
                oxlsSheet.Cells[1, 24] = mm.ToString().PadLeft(2, ' ') + "月" + maxDay.ToString().PadLeft(2, ' ') + "日まで";
                
                // 集計日数：該当月の勤務日数 2018/01/12
                //int n = zd.Count(a => a.sHoliday == 0); 2018/01/17
                //oxlsSheet.Cells[1, 28] = n;   2018/01/17
                oxlsSheet.Cells[1, 21] = kDays;

                // グラフ用データ一括書き込み
                rng = oxlsSheet.Range[oxlsSheet.Cells[3, 31], oxlsSheet.Cells[7, 61]];
                rng.Value2 = xlsArray;

                iX = 0;

                int iR = 0;
                double zRePlanTl = 0;

                // 製造部門と間接部門の処理共通化のため以下、コメント化 2018/06/26

                //// 理由別残業計画シートの内容を配列に取得
                //if (Utility.StrtoInt(szCode.Substring(1, 1)) <= global.flgOn) // 製造部門
                //{
                //    for (int i = 1; i <= zReSeizou.GetLength(0); i++)
                //    {
                //        // 有効な部署コードでないときはネグる：2018/04/12
                //        if (Utility.NulltoStr(zReSeizou[i, 1]).Length < 5)
                //        {
                //            continue;
                //        }

                //        // 部署コード頭4桁が一致しているか？
                //        if (zReSeizou[i, 1].ToString().Substring(0, 4) != szCode.ToString())
                //        {
                //            continue;
                //        }

                //        oxlsSheet.Cells[5 + iR, 20] = zReSeizou[i, 2];    // 理由コード
                //        oxlsSheet.Cells[5 + iR, 21] = zReSeizou[i, 3];    // 残業理由

                //        // 理由別残業時間を集計 : 最大日付範囲で取得 2017/11/22
                //        double s = 0;

                //        // 応援先で集計 2018/02/11
                //        if ((double)zReSeizou[i, 2] < 10)
                //        {
                //            s = dts.残業集計2
                //                .Where(a => a.応援先.Substring(0, 4) == szCode &&
                //                       a.残業理由 == (double)zReSeizou[i, 2] &&
                //                       a.日 <= maxDay)
                //                .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));
                //        }
                //        else
                //        {
                //            s = dts.残業集計2
                //                .Where(a => a.応援先.Substring(0, 4) == szCode &&
                //                       a.残業理由 >= 10 &&
                //                       a.日 <= maxDay)
                //                .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));
                //        }

                //        s = s / 60;
                //        oxlsSheet.Cells[5 + iR, 27] = s;

                //        iR += 2;

                //        if (iR == 20)
                //        {
                //            break;
                //        }
                //    }

                //    // 当月計画時間初期化 2018/04/17
                //    double pp = 0;

                //    // 合計欄に当月計画時間の稼働日数割りを表示する 2018/04/17
                //    foreach (var t in zd.Where(a => a.sDay <= maxDay))
                //    {
                //        // ゼロ以外の最大日付の計画値
                //        if (t.sMonthPlan != 0)
                //        {
                //            pp = t.sMonthPlan;
                //        }
                //    }

                //    oxlsSheet.Cells[27, 25] = pp;

                //    // 当月残業合計
                //    oxlsSheet.Cells[27, 27] = setTotalZan(szCode + "0", outZangyoFile);
                //}
                //else
                //{
                //    for (int i = 0; i < zRe.Length; i++)
                //    {
                //        // 係別理由別残業計画を取得 : 2018/04/11
                //        string [] dd = zRe[i].ToString().Split(',');

                //        if (dd.Length < 5)
                //        {
                //            continue;
                //        }

                //        // 係コードが一致しているか？
                //        if (dd[0].ToString() != szCode.ToString())
                //        {
                //            continue;
                //        }

                //        // 年月コードが一致しているか？
                //        if (dd[1].ToString() != yy.ToString() || dd[2].ToString() != mm.ToString())
                //        {
                //            continue;
                //        }

                //        oxlsSheet.Cells[5 + iR, 20] = dd[3];    // 理由コード

                //        // 残業理由名称を部署別残業理由配列より取得：2018/04/12
                //        for (int ii = 1; ii <= zReSeizou.GetLength(0); ii++)
                //        {
                //            // 有効な部署コードでないときはネグる：2018/04/12
                //            if (Utility.NulltoStr(zReSeizou[ii, 1]).Length < 5)
                //            {
                //                continue;
                //            }

                //            // 部署コード頭4桁が一致しているか？
                //            if (zReSeizou[ii, 1].ToString().Substring(0, 4) != szCode.ToString())
                //            {
                //                continue;
                //            }

                //            // 理由コードが一致しているか？
                //            if (zReSeizou[ii, 2].ToString() != dd[3])
                //            {
                //                continue;
                //            }

                //            oxlsSheet.Cells[5 + iR, 21] = zReSeizou[ii, 3];    // 残業理由
                //            break;
                //        }
                        
                //        //oxlsSheet.Cells[5 + iR, 21] = zRe[i, 5];    // 残業理由
                //        //oxlsSheet.Cells[5 + iR, 25] = Utility.StrtoDouble(zRe[i, 6].ToString()) / (double)kDays * n;    // 現在の日付で稼働日割りした理由別計画値 2018/01/17
                //        //oxlsSheet.Cells[5 + iR, 25] = Utility.StrtoDouble(zRe[i, 6].ToString()) / (double)kDays * d;    // 現在の日付で稼働日割りした理由別計画値 2018/01/17
                //        oxlsSheet.Cells[5 + iR, 25] = Utility.StrtoDouble(dd[4].ToString()) / (double)kDays * d;    // 現在の日付で稼働日割りした理由別計画値 2018/04/11

                //        // 理由別残業時間を集計 : 最大日付範囲で取得 2017/11/22
                //        //double s = dts.残業集計1
                //        //    .Where(a => a.部署コード.Substring(0, 3) == szCode && a.残業理由 == (double)zRe[i, 4] && 
                //        //                a.日 <= maxDay)
                //        //    .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));

                //        /* 理由別残業時間を集計 : 最大日付範囲で取得 2017/11/22
                //                                 応援先を対象とする 2018/02/11 */
                //        double s = dts.残業集計2
                //            .Where(a => a.応援先.Substring(0, 4) == szCode &&
                //                        a.残業理由 == Utility.StrtoDouble(dd[3]) &&
                //                        a.日 <= maxDay)
                //            .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));

                //        s = s / 60;
                //        oxlsSheet.Cells[5 + iR, 27] = s;

                //        iR += 2;
                //        //zRePlanTl += Utility.StrtoDouble(zRe[i, 6].ToString()); // 計画合計に加算
                //        //zRePlanTl += Utility.StrtoDouble(zRe[i, 6].ToString()) / (double)kDays * n;     // 計画合計に加算 2018/01/17
                //        //zRePlanTl += Utility.StrtoDouble(zRe[i, 6].ToString()) / (double)kDays * d;     // 計画合計に加算 2018/01/17
                //        zRePlanTl += Utility.StrtoDouble(dd[4].ToString()) / (double)kDays * d;     // 計画合計に加算 2018/04/11

                //        if (iR == 20)
                //        {
                //            //// 現在の日付で稼働日割りした理由別計画の合計
                //            //oxlsSheet.Cells[27, 25] = zRePlanTl;
                //            break;
                //        }
                //    }

                //    // 現在の日付で稼働日割りした理由別計画の合計
                //    oxlsSheet.Cells[27, 25] = zRePlanTl;

                //    // 当月残業合計
                //    oxlsSheet.Cells[27, 27] = setTotalZan(szCode + "0", outZangyoFile);
                //}


                // ============= 製造部門も理由別残業計画を行うため製造部門と間接部門の処理を共通化 2018/06/26 =============               


                for (int i = 0; i < zRe.Length; i++)
                {
                    // 係別理由別残業計画を取得 : 2018/04/11
                    string[] dd = zRe[i].ToString().Split(',');

                    if (dd.Length < 5)
                    {
                        continue;
                    }

                    // 係コードが一致しているか？
                    if (dd[0].ToString() != szCode.ToString())
                    {
                        continue;
                    }

                    // 年月コードが一致しているか？
                    if (dd[1].ToString() != yy.ToString() || dd[2].ToString() != mm.ToString())
                    {
                        continue;
                    }

                    oxlsSheet.Cells[5 + iR, 20] = dd[3];    // 理由コード

                    // 残業理由名称を部署別残業理由配列より取得：2018/04/12
                    for (int ii = 1; ii <= zReSeizou.GetLength(0); ii++)
                    {
                        // 有効な部署コードでないときはネグる：2018/04/12
                        if (Utility.NulltoStr(zReSeizou[ii, 1]).Length < 5)
                        {
                            continue;
                        }

                        // 部署コード頭4桁が一致しているか？
                        if (zReSeizou[ii, 1].ToString().Substring(0, 4) != szCode.ToString())
                        {
                            continue;
                        }

                        // 理由コードが一致しているか？
                        if (zReSeizou[ii, 2].ToString() != dd[3])
                        {
                            continue;
                        }

                        oxlsSheet.Cells[5 + iR, 21] = zReSeizou[ii, 3];    // 残業理由
                        break;
                    }

                    oxlsSheet.Cells[5 + iR, 25] = Utility.StrtoDouble(dd[4].ToString()) / (double)kDays * d;    // 現在の日付で稼働日割りした理由別計画値 2018/04/11
                    
                    /* 理由別残業時間を集計 : 最大日付範囲で取得 2017/11/22
                                             応援先を対象とする 2018/02/11 */
                    double s = dts.残業集計2
                        .Where(a => a.応援先.Substring(0, 4) == szCode &&
                                    a.残業理由 == Utility.StrtoDouble(dd[3]) &&
                                    a.日 <= maxDay)
                        .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));

                    s = s / 60;
                    oxlsSheet.Cells[5 + iR, 27] = s;

                    iR += 2;
                    zRePlanTl += Utility.StrtoDouble(dd[4].ToString()) / (double)kDays * d;     // 計画合計に加算 2018/04/11

                    if (iR == 20)
                    {
                        //// 現在の日付で稼働日割りした理由別計画の合計
                        //oxlsSheet.Cells[27, 25] = zRePlanTl;
                        break;
                    }
                }

                // 現在の日付で稼働日割りした理由別計画の合計
                oxlsSheet.Cells[27, 25] = zRePlanTl;

                // 当月残業合計
                oxlsSheet.Cells[27, 27] = setTotalZan(szCode + "0", outZangyoFile);

                //==================================  2018/06/26 ここまで =============================================================


                // 前月との比較要素のセット
                oxlsSheet.Cells[35, 23] = zSeisan;                      // 前月生産数
                oxlsSheet.Cells[37, 23] = zNin;                         // 前月人数
                oxlsSheet.Cells[41, 23] = zenZan / zNin / zenKaDays;    // 前月1人当り残業（実績 / 人数 / 稼働日数）  : 2018/01/20
                oxlsSheet.Cells[39, 23] = zenZan / zenKaDays;           // 前月一日当たり合計
                oxlsSheet.Cells[35, 25] = sSeisan;                      // 生産数
                oxlsSheet.Cells[37, 25] = sNin;                         // 人数

                //MessageBox.Show(msg);

                // シート名 2018/02/17
                string nm = "'" + oxlsSheet.Name + "'";

                // 2018/02/08 データ範囲(系列値）の可変表示
                Excel.ChartObject ss = (Excel.ChartObject)oxlsSheet.ChartObjects(1);

                string msg = string.Empty;
                for (int i = 1; i <= 4; i++)
                {
                    // =SERIES('151総務課'!$AD$4,'151総務課'!$AE$3:$BI$3,'151総務課'!$AE$4:$BI$4,1)
                    Excel.Series sr = ss.Chart.SeriesCollection(i);
                    msg += sr.Formula + Environment.NewLine;

                    // =SERIES('111第１組立課'!$AD$4,'111第１組立課'!day,'111第１組立課'!date,1)
                    string f = sr.Formula;
                    f = f.Replace("=SERIES(", "");
                    f = f.Replace(")", "");
                    string[] fa = f.Split(',');
                    
                    // =SERIES('111第１組立課'!$AD$4,'111第１組立課'!day,'111第１組立課'!date,1)

                    // 系列名、系列値のシート名も直接セットする 2018/02/17
                    if (fa[3] == "1")
                    {
                        // 日付別
                        fa[0] = nm + "!$AD$4";      // 系列名
                        fa[1] = nm + "!day";        // 軸ラベルの範囲
                        fa[2] = nm + "!date";       // 系列値
                    }
                    else if (fa[3] == "2")
                    {
                        // 計画
                        fa[0] = nm + "!$AD$5";      // 系列名
                        fa[1] = nm + "!day";        // 軸ラベルの範囲
                        fa[2] = nm + "!plan";       // 系列値

                        // グラフの色を変更 2018/03/08
                        sr.Format.Line.ForeColor.RGB = System.Drawing.Color.DimGray.ToArgb();
                    }
                    else if (fa[3] == "3")
                    {
                        // 実績
                        fa[0] = nm + "!$AD$6";          // 系列名
                        fa[1] = nm + "!day";            // 軸ラベルの範囲
                        fa[2] = nm + "!Performance";    // 系列値
                    }
                    else if (fa[3] == "4")
                    {
                        // 目標
                        fa[0] = nm + "!$AD$7";      // 系列名
                        fa[1] = nm + "!day";        // 軸ラベルの範囲
                        fa[2] = nm + "!goal";       // 系列値
                    }

                    string newF = "=SERIES(" + fa[0] + "," + fa[1] + "," + fa[2] + "," + fa[3] + ")";
                    //msg += newF + Environment.NewLine;

                    sr.Formula = newF;
                }

                // シート全体が印刷対象とするためセルをアクティブとする 2018/02/09
                oxlsSheet.Range["A1"].Select();
            }
            catch (Exception ex)
            {
                MessageBox.Show(szCode + "," + ex.Message, "残業グラフエクセル出力", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
            }
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     エクセルシートへ出力：課別 2017/10/11 </summary>
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
        private void xlsOutPut_KA(Excel.Application oXls, ref Excel.Workbook oXlsBook, ref Excel.Worksheet oxlsSheet, clsZanSum[] zd, string xlsFile, int yy, int mm, object[,] zReSeizou, object[] zRe, sqlControl.DataControl sdCon)
        {
            oxlsSheet.Select(Type.Missing);
            
            Excel.Range rng = null;

            int iX = 1;

            // 実稼働日数を取得
            //int kDays = zd.Where(a => a.sHoliday == 0 || a.sZangyo > 0).Count(); 2018/01/17
            int kDays = zd.Count(a => a.sHoliday == 0); // 2018/01/17 休日出勤は含まない

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
                oxlsSheet.Cells[1, 11] = szCode + "00";
                oxlsSheet.Cells[1, 13] = getDepartmentName(_dbName, szCode + "00", sdCon);

                // シート名に部署名をつける
                oxlsSheet.Name = szCode + oxlsSheet.Cells[1, 13].value;

                //// 実稼働日数を求める
                //int d = zd.Where(a => a.sHoliday == 0 || a.sZangyo > 0).Count();
                //oxlsSheet.Cells[1, 21] = d.ToString();

                // 稼働日数を求める:休日出勤を含まない集計日までの稼働日数 2018/01/08
                int d = zd.Where(a => a.sHoliday == 0 && a.sDay <= days).Count();
                //oxlsSheet.Cells[1, 21] = d.ToString(); 2018/01/17
                oxlsSheet.Cells[1, 28] = d.ToString();  // 2018/01/17

                // 集計期間
                int maxDay = days;
                oxlsSheet.Cells[1, 24] = mm.ToString().PadLeft(2, ' ') + "月" + maxDay.ToString().PadLeft(2, ' ') + "日まで";

                //if (dts.過去勤務票ヘッダ.Any(a => a.年 == yy && a.月 == mm && a.部署コード == szCode))
                //{
                //    maxDay = dts.過去勤務票ヘッダ.Where(a => a.年 == yy && a.月 == mm && a.部署コード == szCode).Max(a => a.日);
                //    oxlsSheet.Cells[1, 24] = mm.ToString().PadLeft(2, ' ') + "月" + maxDay.ToString().PadLeft(2, ' ') + "日まで";
                //}
                //else
                //{
                //    maxDay = 0;
                //    oxlsSheet.Cells[1, 24] = "勤怠データなし";
                //}

                //// 集計日数
                //int n = zd.Count(a => (a.sHoliday == 0 || a.sZangyo > 0) && a.sDay <= maxDay);
                //oxlsSheet.Cells[1, 28] = n;

                // 集計日数：該当月の勤務日数 2018/01/12
                //int n = zd.Count(a => a.sHoliday == 0); 2018/01/17
                //oxlsSheet.Cells[1, 28] = n;   2018/01/17
                oxlsSheet.Cells[1, 21] = kDays;

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


                // 製造部門と間接部門の処理共通化のため以下、コメント化 2018/06/26

                //// 理由別残業計画シートの内容を配列に取得
                //if (Utility.StrtoInt(szCode.Substring(1, 1)) <= global.flgOn) // 製造部門
                //{
                //    // 製造部門：理由別残業計画はなし
                //    for (int i = 1; i <= zReSeizou.GetLength(0); i++)
                //    {
                //        // 部署コード頭3桁が一致しているか？
                //        if (zReSeizou[i, 1].ToString().Substring(0, 3) != szCode.ToString())
                //        {
                //            continue;
                //        }

                //        oxlsSheet.Cells[5 + iR, 20] = zReSeizou[i, 2];    // 理由コード
                //        oxlsSheet.Cells[5 + iR, 21] = zReSeizou[i, 3];    // 残業理由
                        
                //        // 理由別残業時間を集計 : 最大日付範囲で取得 2017/11/22
                //        double s = 0;

                //        //if ((double)zReSeizou[i, 2] < 10)
                //        //{
                //        //    s = dts.残業集計1
                //        //        .Where(a => a.部署コード.Substring(0, 3) == szCode && 
                //        //               a.残業理由 == (double)zReSeizou[i, 2] && 
                //        //               a.日 <= maxDay)
                //        //        .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));
                //        //}
                //        //else
                //        //{
                //        //    s = dts.残業集計1
                //        //        .Where(a => a.部署コード.Substring(0, 3) == szCode && 
                //        //               a.残業理由 >= 10 && 
                //        //               a.日 <= maxDay)
                //        //        .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));
                //        //}

                //        // 応援先で集計 2018/02/11
                //        if ((double)zReSeizou[i, 2] < 10)
                //        {
                //            s = dts.残業集計1
                //                .Where(a => a.応援先.Substring(0, 3) == szCode &&
                //                       a.残業理由 == (double)zReSeizou[i, 2] &&
                //                       a.日 <= maxDay)
                //                .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));
                //        }
                //        else
                //        {
                //            s = dts.残業集計1
                //                .Where(a => a.応援先.Substring(0, 3) == szCode &&
                //                       a.残業理由 >= 10 &&
                //                       a.日 <= maxDay)
                //                .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));
                //        }

                //        s = s / 60;
                //        oxlsSheet.Cells[5 + iR, 27] = s;

                //        iR += 2;

                //        if (iR == 20)
                //        {
                //            break;
                //        }
                //    }

                //    // 当月計画時間初期化 2018/04/17
                //    double pp = 0;

                //    // 合計欄に当月計画時間の稼働日数割りを表示する 2018/04/17
                //    foreach (var t in zd.Where(a => a.sDay <= maxDay))
                //    {
                //        // ゼロ以外の最大日付の計画値
                //        if (t.sMonthPlan != 0)
                //        {
                //            pp = t.sMonthPlan;
                //        }
                //    }

                //    oxlsSheet.Cells[27, 25] = pp;

                //    // 当月残業合計
                //    oxlsSheet.Cells[27, 27] = setTotalZan(szCode + "00", outZangyoFile);
                //}
                //else
                //{
                //    // 間接部門
                //    for (int i = 0; i < zRe.Length; i++)
                //    {
                //        // 課別理由別残業計画を取得 : 2018/04/13
                //        string[] dd = zRe[i].ToString().Split(',');

                //        if (dd.Length < 5)
                //        {
                //            continue;
                //        }

                //        // 課コードが一致しているか？
                //        if (dd[0].ToString() != szCode.ToString())
                //        {
                //            continue;
                //        }

                //        // 年月コードが一致しているか？
                //        if (dd[1].ToString() != yy.ToString() || dd[2].ToString() != mm.ToString())
                //        {
                //            continue;
                //        }

                //        oxlsSheet.Cells[5 + iR, 20] = dd[3];    // 理由コード

                //        // 残業理由名称を部署別残業理由配列より取得：2018/04/13
                //        for (int ii = 1; ii <= zReSeizou.GetLength(0); ii++)
                //        {
                //            // 有効な部署コードでないときはネグる：2018/04/13
                //            if (Utility.NulltoStr(zReSeizou[ii, 1]).Length < 5)
                //            {
                //                continue;
                //            }

                //            // 部署コード頭3桁が一致しているか？
                //            if (zReSeizou[ii, 1].ToString().Substring(0, 3) != szCode.ToString())
                //            {
                //                continue;
                //            }

                //            // 理由コードが一致しているか？
                //            if (zReSeizou[ii, 2].ToString() != dd[3])
                //            {
                //                continue;
                //            }

                //            oxlsSheet.Cells[5 + iR, 21] = zReSeizou[ii, 3];    // 残業理由
                //            break;
                //        }
                        
                //        //oxlsSheet.Cells[5 + iR, 25] = Utility.StrtoDouble(zRe[i, 6].ToString()) / (double)kDays * n;  // 現在の日付で稼働日割りした理由別計画値 2018/01/17
                //        //oxlsSheet.Cells[5 + iR, 25] = Utility.StrtoDouble(zRe[i, 6].ToString()) / (double)kDays * d;  // 現在の日付で稼働日割りした理由別計画値 2018/01/17
                //        oxlsSheet.Cells[5 + iR, 25] = Utility.StrtoDouble(dd[4].ToString()) / (double)kDays * d;        // 現在の日付で稼働日割りした理由別計画値 2018/04/13

                //        // 理由別残業時間を集計 : 最大日付範囲で取得 2017/11/22
                //        //double s = dts.残業集計1
                //        //    .Where(a => a.部署コード.Substring(0, 3) == szCode && a.残業理由 == (double)zRe[i, 4] && 
                //        //                a.日 <= maxDay)
                //        //    .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));

                //        /* 理由別残業時間を集計 : 最大日付範囲で取得 2017/11/22
                //                                 応援先を対象とする 2018/02/11 */
                //        double s = dts.残業集計1
                //            .Where(a => a.応援先.Substring(0, 3) == szCode && 
                //                        a.残業理由 == Utility.StrtoDouble(dd[3]) &&
                //                        a.日 <= maxDay)
                //            .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));

                //        s = s / 60;
                //        oxlsSheet.Cells[5 + iR, 27] = s;

                //        iR += 2;
                //        //zRePlanTl += Utility.StrtoDouble(zRe[i, 6].ToString()); // 計画合計に加算
                //        //zRePlanTl += Utility.StrtoDouble(zRe[i, 6].ToString()) / (double)kDays * n;   // 計画合計に加算 2018/01/17
                //        //zRePlanTl += Utility.StrtoDouble(zRe[i, 6].ToString()) / (double)kDays * d;   // 計画合計に加算 2018/01/17
                //        zRePlanTl += Utility.StrtoDouble(dd[4].ToString()) / (double)kDays * d;         // 計画合計に加算 2018/04/13

                //        if (iR == 20)
                //        {
                //            //// 現在の日付で稼働日割りした理由別計画の合計
                //            //oxlsSheet.Cells[27, 25] = zRePlanTl;
                //            break;
                //        }
                //    }

                //    // 現在の日付で稼働日割りした理由別計画の合計
                //    oxlsSheet.Cells[27, 25] = zRePlanTl;

                //    // 当月残業合計
                //    oxlsSheet.Cells[27, 27] = setTotalZan(szCode + "00", outZangyoFile);
                //}




                // ============= 製造部門も理由別残業計画を行うため製造部門と間接部門の処理を共通化 2018/06/26 ============= 
                
                for (int i = 0; i < zRe.Length; i++)
                {
                    // 課別理由別残業計画を取得 : 2018/04/13
                    string[] dd = zRe[i].ToString().Split(',');

                    if (dd.Length < 5)
                    {
                        continue;
                    }

                    // 課コードが一致しているか？
                    if (dd[0].ToString() != szCode.ToString())
                    {
                        continue;
                    }

                    // 年月コードが一致しているか？
                    if (dd[1].ToString() != yy.ToString() || dd[2].ToString() != mm.ToString())
                    {
                        continue;
                    }

                    oxlsSheet.Cells[5 + iR, 20] = dd[3];    // 理由コード

                    // 残業理由名称を部署別残業理由配列より取得：2018/04/13
                    for (int ii = 1; ii <= zReSeizou.GetLength(0); ii++)
                    {
                        // 有効な部署コードでないときはネグる：2018/04/13
                        if (Utility.NulltoStr(zReSeizou[ii, 1]).Length < 5)
                        {
                            continue;
                        }

                        // 部署コード頭3桁が一致しているか？
                        if (zReSeizou[ii, 1].ToString().Substring(0, 3) != szCode.ToString())
                        {
                            continue;
                        }

                        // 理由コードが一致しているか？
                        if (zReSeizou[ii, 2].ToString() != dd[3])
                        {
                            continue;
                        }

                        oxlsSheet.Cells[5 + iR, 21] = zReSeizou[ii, 3];    // 残業理由
                        break;
                    }

                    oxlsSheet.Cells[5 + iR, 25] = Utility.StrtoDouble(dd[4].ToString()) / (double)kDays * d;        // 現在の日付で稼働日割りした理由別計画値 2018/04/13
                    
                    /* 理由別残業時間を集計 : 最大日付範囲で取得 2017/11/22
                                             応援先を対象とする 2018/02/11 */
                    double s = dts.残業集計1
                        .Where(a => a.応援先.Substring(0, 3) == szCode &&
                                    a.残業理由 == Utility.StrtoDouble(dd[3]) &&
                                    a.日 <= maxDay)
                        .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));

                    s = s / 60;
                    oxlsSheet.Cells[5 + iR, 27] = s;

                    iR += 2;
                    zRePlanTl += Utility.StrtoDouble(dd[4].ToString()) / (double)kDays * d;         // 計画合計に加算 2018/04/13

                    if (iR == 20)
                    {
                        //// 現在の日付で稼働日割りした理由別計画の合計
                        //oxlsSheet.Cells[27, 25] = zRePlanTl;
                        break;
                    }
                }

                // 現在の日付で稼働日割りした理由別計画の合計
                oxlsSheet.Cells[27, 25] = zRePlanTl;

                // 当月残業合計
                oxlsSheet.Cells[27, 27] = setTotalZan(szCode + "00", outZangyoFile);


                //==================================  2018/06/26 ここまで =============================================================

                                
                // 前月との比較要素のセット
                oxlsSheet.Cells[35, 23] = zSeisan;                      // 前月生産数
                oxlsSheet.Cells[37, 23] = zNin;                         // 前月人数
                oxlsSheet.Cells[41, 23] = zenZan / zNin / zenKaDays;    // 前月1人当り残業（実績 / 人数 / 稼働日数）  : 2018/01/20
                oxlsSheet.Cells[39, 23] = zenZan / zenKaDays;           // 前月一日当たり合計

                oxlsSheet.Cells[35, 25] = sSeisan;                      // 生産数
                oxlsSheet.Cells[37, 25] = sNin;                         // 人数

                //MessageBox.Show(msg);

                // シート名 2018/02/17
                string nm = "'" + oxlsSheet.Name + "'"; 

                // 2018/02/08 データ範囲(系列値）の可変表示
                Excel.ChartObject ss = (Excel.ChartObject)oxlsSheet.ChartObjects(1);

                string msg = string.Empty;
                for (int i = 1; i <= 4; i++)
                {
                    // =SERIES('151総務課'!$AD$4,'151総務課'!$AE$3:$BI$3,'151総務課'!$AE$4:$BI$4,1)
                    Excel.Series sr = ss.Chart.SeriesCollection(i);
                    msg += sr.Formula + Environment.NewLine;

                    // =SERIES('111第１組立課'!$AD$4,'111第１組立課'!day,'111第１組立課'!date,1)
                    string f = sr.Formula;
                    f = f.Replace("=SERIES(", "");
                    f = f.Replace(")", "");

                    string[] fa = f.Split(',');

                    //if (fa[3] == "1")
                    //{
                    //    // 日付別
                    //    string[] p = fa[1].Split('!');
                    //    fa[1] = p[0] + "!" + "day";

                    //    p = fa[2].Split('!');
                    //    fa[2] = p[0] + "!" + "date";
                    //}
                    //else if (fa[3] == "2")
                    //{
                    //    // 計画
                    //    string[] p = fa[1].Split('!');
                    //    fa[1] = p[0] + "!" + "day";

                    //    p = fa[2].Split('!');
                    //    fa[2] = p[0] + "!" + "plan";
                    //}
                    //else if (fa[3] == "3")
                    //{
                    //    // 実績
                    //    string[] p = fa[1].Split('!');
                    //    fa[1] = p[0] + "!" + "day";

                    //    p = fa[2].Split('!');
                    //    fa[2] = p[0] + "!" + "Performance";
                    //}
                    //else if (fa[3] == "4")
                    //{
                    //    // 目標
                    //    string[] p = fa[1].Split('!');
                    //    fa[1] = p[0] + "!" + "day";

                    //    p = fa[2].Split('!');
                    //    fa[2] = p[0] + "!" + "goal";
                    //}
                                        
                    // =SERIES('111第１組立課'!$AD$4,'111第１組立課'!day,'111第１組立課'!date,1)

                    // 系列名、系列値のシート名も直接セットする 2018/02/17
                    if (fa[3] == "1")
                    {
                        // 日付別
                        fa[0] = nm + "!$AD$4";      // 系列名
                        fa[1] = nm + "!day";        // 軸ラベルの範囲
                        fa[2] = nm + "!date";       // 系列値
                    }
                    else if (fa[3] == "2")
                    {
                        // 計画
                        fa[0] = nm + "!$AD$5";      // 系列名
                        fa[1] = nm + "!day";        // 軸ラベルの範囲
                        fa[2] = nm + "!plan";       // 系列値

                        // グラフの色を変更 2018/03/08
                        sr.Format.Line.ForeColor.RGB = System.Drawing.Color.DimGray.ToArgb();
                    }
                    else if (fa[3] == "3")
                    {
                        // 実績
                        fa[0] = nm + "!$AD$6";          // 系列名
                        fa[1] = nm + "!day";            // 軸ラベルの範囲
                        fa[2] = nm + "!Performance";    // 系列値
                    }
                    else if (fa[3] == "4")
                    {
                        // 目標
                        fa[0] = nm + "!$AD$7";      // 系列名
                        fa[1] = nm + "!day";        // 軸ラベルの範囲
                        fa[2] = nm + "!goal";       // 系列値
                    }

                    string newF = "=SERIES(" + fa[0] + "," + fa[1] + "," + fa[2] + "," + fa[3] + ")";
                    //msg += newF + Environment.NewLine;

                    sr.Formula = newF;
                }

                // シート全体が印刷対象とするためセルをアクティブとする 2018/02/09
                oxlsSheet.Range["A1"].Select();
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
        private void xlsOutPutBumon(Excel.Application oXls, ref Excel.Workbook oXlsBook, ref Excel.Worksheet oxlsSheet, clsZanSum[] zd, string xlsFile, int yy, int mm, object[,] zReSeizou, object[] zRe, string sBmn)
        {
            oxlsSheet.Select(Type.Missing);

            Excel.Range rng = null;

            int iX = 1;

            // 実稼働日数を取得
            //int kDays = zd.Where(a => a.sHoliday == 0 || a.sZangyo > 0).Count(); 2018/01/17
            int kDays = zd.Count(a => a.sHoliday == 0);

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
                //int d = zd.Where(a => a.sHoliday == 0 || a.sZangyo > 0).Count();
                //oxlsSheet.Cells[1, 21] = d.ToString();

                // 実稼働日数を求める:休日出勤を含まない集計日までの稼働日数 2018/01/08
                int d = zd.Where(a => a.sHoliday == 0 && a.sDay <= days).Count();
                //oxlsSheet.Cells[1, 21] = d.ToString();
                oxlsSheet.Cells[1, 28] = d.ToString();

                // 集計期間
                int maxDay = days; 
                oxlsSheet.Cells[1, 24] = mm.ToString().PadLeft(2, ' ') + "月" + maxDay.ToString().PadLeft(2, ' ') + "日まで";

                // 集計日数：該当月の勤務日数 2018/01/12
                //int n = zd.Count(a => a.sHoliday == 0); // 2018/01/17
                //oxlsSheet.Cells[1, 28] = n; // 2018/01/17
                oxlsSheet.Cells[1, 21] = kDays; // 2018/01/17
                
                // グラフ用データ一括書き込み　← 書き込むと数値が文字扱いとなりグラフが描画されない
                //rng = oxlsSheet.Range[oxlsSheet.Cells[4, 31], oxlsSheet.Cells[7, 61]];
                //rng.NumberFormatLocal = "0.0";
                rng = oxlsSheet.Range[oxlsSheet.Cells[3, 31], oxlsSheet.Cells[7, 61]];
                rng.Value2 = xlsArray;

                iX = 0;

                int iR = 0;
                double zRePlanTl = 0;
                double zOver10 = 0;
                double zzz = 0;
                

                // 製造部門と間接部門の処理共通化のため以下、コメント化 2018/07/04

                //// 理由別残業計画：全社または製造部門
                //if (szCode == "0" || szCode == "1")
                //{
                //    // 製造部門：理由別残業計画はなし                    
                //    oxlsSheet.Cells[5 + iR, 20] = "";   // 理由コード
                //    oxlsSheet.Cells[5 + iR, 21] = "";   // 残業理由

                //    // 理由別残業時間を集計
                //    IEnumerable<riyuZan> s = getRiyuZanSum(szCode);

                //    foreach (var t in s)
                //    {
                //        if (t.riyu >= 10)
                //        {
                //            iR = 23;
                //            zzz = t.zan / 60;
                //            zOver10 += zzz;
                //            oxlsSheet.Cells[iR, 27] = zOver10;
                //        }
                //        else
                //        {
                //            iR = (int)t.riyu * 2 + 3;
                //            zzz = t.zan / 60;
                //            oxlsSheet.Cells[iR, 27] = zzz;
                //        }

                //        oxlsSheet.Cells[iR, 20] = "";   // 理由コード
                //        oxlsSheet.Cells[iR, 21] = "";   // 残業理由
                //    }
                    
                //    // 当月計画時間初期化 2018/04/17
                //    double pp = 0;

                //    // 合計欄に当月計画時間の稼働日数割りを表示する 2018/04/17
                //    foreach (var t in zd.Where(a => a.sDay <= maxDay))
                //    {
                //        // ゼロ以外の最大日付の計画値
                //        if (t.sMonthPlan != 0)
                //        {
                //            pp = t.sMonthPlan;
                //        }
                //    }

                //    oxlsSheet.Cells[27, 25] = pp;

                //    // 当月残業合計
                //    oxlsSheet.Cells[27, 27] = setTotalZanBumon(szCode, outZangyoFile);
                //}
                //else if (Utility.StrtoInt(szCode) >= 2)
                //{
                //    // 間接部門
                //    //zRe = bs.getZanReasonPlan();

                //    double[] keikaku = new double[10];
                //    for (int ind = 0; ind < keikaku.Length; ind++)
                //    {
                //        keikaku[ind] = 0;
                //    }

                //    // 理由別計画集計値を配列にセット
                //    for (int i = 0; i < zRe.Length; i++)
                //    {
                //        // 課別理由別残業計画を取得 : 2018/04/13
                //        string[] dd = zRe[i].ToString().Split(',');

                //        // 部門（製造、間接）コードが一致しているか？
                //        if (dd[0].ToString().Substring(1, 1) == "0" ||  
                //            dd[0].ToString().Substring(1, 1) == "1")
                //        {
                //            continue;
                //        }

                //        // 年月コードが一致しているか？
                //        if (dd[1].ToString() != yy.ToString() || dd[2].ToString() != mm.ToString())
                //        {
                //            continue;
                //        }
                        
                //        int p = Utility.StrtoInt(dd[3].ToString()); // 理由コード

                //        if (p > 10)
                //        {
                //            p = 10;
                //        }

                //        keikaku[p - 1] += Utility.StrtoDouble(dd[4].ToString());   // 計画値
                //    }

                //    // 理由別計画集計値配列を順次読む
                //    for (int ind = 0; ind < keikaku.Length; ind++)
                //    {
                //        oxlsSheet.Cells[5 + iR, 20] = ind + 1;          // 理由コード
                //        //oxlsSheet.Cells[5 + iR, 21] = zRe[i, 5];      // 残業理由
                //        oxlsSheet.Cells[5 + iR, 21] = "";               // 残業理由
                //        //oxlsSheet.Cells[5 + iR, 25] = keikaku[ind] / (double)kDays * n;     // 現在の日付で稼働日割りした理由別計画値 2018/01/17
                //        oxlsSheet.Cells[5 + iR, 25] = keikaku[ind] / (double)kDays * d;     // 現在の日付で稼働日割りした理由別計画値 2018/01/17

                //        double s = 0;

                //        // 理由別残業時間を集計：最大日付範囲で取得 2017/11/22

                //        //// 社員所属で集計
                //        //if (ind < 9)
                //        //{
                //        //    s = dts.残業集計1
                //        //        .Where(a => a.部署コード.Substring(1, 1) != "1" && a.部署コード.Substring(1, 1) != "0" && 
                //        //        a.残業理由 == (double)ind + 1 && a.日 <= days)
                //        //        .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));
                //        //}
                //        //else
                //        //{
                //        //    s = dts.残業集計1
                //        //        .Where(a => a.部署コード.Substring(1, 1) != "1" && a.部署コード.Substring(1, 1) != "0" &&
                //        //            a.残業理由 >= 10 && a.日 <= days)
                //        //        .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));
                //        //}

                //        // 応援先で集計 2018/02/11
                //        if (ind < 9)
                //        {
                //            s = dts.残業集計1
                //                .Where(a => a.応援先.Substring(1, 1) != "1" && 
                //                            a.応援先.Substring(1, 1) != "0" &&
                //                            a.残業理由 == (double)ind + 1 && 
                //                            a.日 <= days)
                //                .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));
                //        }
                //        else
                //        {
                //            s = dts.残業集計1
                //                .Where(a => a.応援先.Substring(1, 1) != "1" && 
                //                            a.応援先.Substring(1, 1) != "0" &&
                //                            a.残業理由 >= 10 && 
                //                            a.日 <= days)
                //                .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));
                //        }

                //        s = s / 60;
                //        oxlsSheet.Cells[5 + iR, 27] = s;

                //        iR += 2;
                //        //zRePlanTl += keikaku[ind] / (double)kDays * n; // 計画合計に加算 2018/01/17
                //        zRePlanTl += keikaku[ind] / (double)kDays * d; // 計画合計に加算 2018/01/17

                //        if (iR == 20)
                //        {
                //            // 計画合計
                //            oxlsSheet.Cells[27, 25] = zRePlanTl;
                //        }
                //    }

                //    // 当月残業合計
                //    oxlsSheet.Cells[27, 27] = setTotalZanBumon(szCode, outZangyoFile);
                //}

                

                // ============= 製造部門も理由別残業計画を行うため製造部門と間接部門の処理を共通化 2018/07/04 =============               

                double[] keikaku = new double[10];
                for (int ind = 0; ind < keikaku.Length; ind++)
                {
                    keikaku[ind] = 0;
                }

                // 理由別計画集計値を配列にセット
                for (int i = 0; i < zRe.Length; i++)
                {
                    // 課別理由別残業計画を取得 : 2018/04/13
                    string[] dd = zRe[i].ToString().Split(',');

                    // 部門（製造、間接）コードが一致しているか？ 2018/07/04
                    if (Utility.StrtoInt(szCode) == 1)
                    {
                        // 製造部門のとき
                        if (Utility.StrtoInt(dd[0].ToString().Substring(1, 1)) > 1)
                        {
                            continue;
                        }
                    }
                    else if (Utility.StrtoInt(szCode) == 2)
                    {
                        // 間接部門のとき
                        if (Utility.StrtoInt(dd[0].ToString().Substring(1, 1)) < 2)
                        {
                            continue;
                        }
                    }

                    // 年月コードが一致しているか？
                    if (dd[1].ToString() != yy.ToString() || dd[2].ToString() != mm.ToString())
                    {
                        continue;
                    }

                    int p = Utility.StrtoInt(dd[3].ToString()); // 理由コード

                    if (p > 10)
                    {
                        p = 10;
                    }

                    keikaku[p - 1] += Utility.StrtoDouble(dd[4].ToString());   // 計画値
                }


                // 理由別計画集計値配列を順次読む
                for (int ind = 0; ind < keikaku.Length; ind++)
                {
                    oxlsSheet.Cells[5 + iR, 20] = ind + 1;          // 理由コード
                    oxlsSheet.Cells[5 + iR, 21] = "";               // 残業理由
                    oxlsSheet.Cells[5 + iR, 25] = keikaku[ind] / (double)kDays * d;     // 現在の日付で稼働日割りした理由別計画値 2018/01/17

                    // 間接部門
                    if (szCode == "2")
                    {
                        double s = 0;

                        // 応援先で集計 2018/02/11
                        if (ind < 9)
                        {
                            s = dts.残業集計1
                                .Where(a => a.応援先.Substring(1, 1) != "1" &&
                                            a.応援先.Substring(1, 1) != "0" &&
                                            a.残業理由 == (double)ind + 1 &&
                                            a.日 <= days)
                                .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));
                        }
                        else
                        {
                            s = dts.残業集計1
                                .Where(a => a.応援先.Substring(1, 1) != "1" &&
                                            a.応援先.Substring(1, 1) != "0" &&
                                            a.残業理由 >= 10 &&
                                            a.日 <= days)
                                .Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10));
                        }

                        s = s / 60;
                        oxlsSheet.Cells[5 + iR, 27] = s;

                        iR += 2;

                        zRePlanTl += keikaku[ind] / (double)kDays * d; // 計画合計に加算 2018/01/17

                        if (iR == 20)
                        {
                            // 計画合計
                            oxlsSheet.Cells[27, 25] = zRePlanTl;
                        }
                    }
                    else
                    {
                        iR += 2;

                        zRePlanTl += keikaku[ind] / (double)kDays * d; // 計画合計に加算 2018/01/17

                        if (iR == 20)
                        {
                            // 計画合計
                            oxlsSheet.Cells[27, 25] = zRePlanTl;
                        }
                    }
                }


                // 全社または製造部門
                if (szCode == "0" || szCode == "1")
                {
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
                }


                // 当月残業合計
                oxlsSheet.Cells[27, 27] = setTotalZanBumon(szCode, outZangyoFile);


                //==================================  2018/07/04 ここまで =============================================================
                


                // 前月との比較要素のセット
                oxlsSheet.Cells[35, 23] = zSeisan;                      // 前月生産数
                oxlsSheet.Cells[37, 23] = zNin;                         // 前月人数
                oxlsSheet.Cells[41, 23] = zenZan / zNin / zenKaDays;    // 前月1人当り残業（実績 / 人数 / 稼働日数）  : 2018/01/08
                oxlsSheet.Cells[39, 23] = zenZan / zenKaDays;           // 前月一日当たり合計

                oxlsSheet.Cells[35, 25] = sSeisan;              // 生産数
                oxlsSheet.Cells[37, 25] = sNin;                 // 人数

                oxlsSheet.Range["A1"].Select();

                // シート名 2018/02/17
                string nm = "'" + oxlsSheet.Name + "'"; 

                // 2018/02/08 データ範囲(系列値）の可変表示
                Excel.ChartObject ss = (Excel.ChartObject)oxlsSheet.ChartObjects(1);

                // コメント化 2018/03/23
                //// 2018/02/08 全社のときグラフ幅を拡張
                //if (szCode == "0")
                //{
                //    ss.Width = ss.Width * 1.538;
                //}

                // 2018/02/08 データ範囲(系列値）の可変表示
                string msg = string.Empty;
                for (int i = 1; i <= 4; i++)
                {
                    // 例：=SERIES('151総務課'!$AD$4,'151総務課'!$AE$3:$BI$3,'151総務課'!$AE$4:$BI$4,1)
                    Excel.Series sr = ss.Chart.SeriesCollection(i);
                    msg += sr.Formula + Environment.NewLine;

                    // 例：=SERIES('111第１組立課'!$AD$4,'111第１組立課'!day,'111第１組立課'!date,1)
                    string f = sr.Formula;
                    f = f.Replace("=SERIES(", "");
                    f = f.Replace(")", "");

                    string[] fa = f.Split(',');

                    // 以下、コメント化 2018/02/19
                    //if (fa[3] == "1")
                    //{
                    //    // 日付別
                    //    string[] p = fa[1].Split('!');
                    //    fa[1] = p[0] + "!" + "day";

                    //    p = fa[2].Split('!');
                    //    fa[2] = p[0] + "!" + "date";
                    //}
                    //else if (fa[3] == "2")
                    //{
                    //    // 計画
                    //    string[] p = fa[1].Split('!');
                    //    fa[1] = p[0] + "!" + "day";

                    //    p = fa[2].Split('!');
                    //    fa[2] = p[0] + "!" + "plan";
                    //}
                    //else if (fa[3] == "3")
                    //{
                    //    // 実績
                    //    string[] p = fa[1].Split('!');
                    //    fa[1] = p[0] + "!" + "day";

                    //    p = fa[2].Split('!');
                    //    fa[2] = p[0] + "!" + "Performance";
                    //}
                    //else if (fa[3] == "4")
                    //{
                    //    // 目標
                    //    string[] p = fa[1].Split('!');
                    //    fa[1] = p[0] + "!" + "day";

                    //    p = fa[2].Split('!');
                    //    fa[2] = p[0] + "!" + "goal";
                    //}

                    // 例：=SERIES('111第１組立課'!$AD$4,'111第１組立課'!day,'111第１組立課'!date,1)

                    // 系列名、系列値のシート名も直接セットする 2018/02/19
                    if (fa[3] == "1")
                    {
                        // 日付別
                        fa[0] = nm + "!$AD$4";      // 系列名
                        fa[1] = nm + "!day";        // 軸ラベルの範囲
                        fa[2] = nm + "!date";       // 系列値
                    }
                    else if (fa[3] == "2")
                    {
                        // 計画
                        fa[0] = nm + "!$AD$5";      // 系列名
                        fa[1] = nm + "!day";        // 軸ラベルの範囲
                        fa[2] = nm + "!plan";       // 系列値

                        // グラフの色を変更 2018/03/08
                        sr.Format.Line.ForeColor.RGB = System.Drawing.Color.DimGray.ToArgb();
                    }
                    else if (fa[3] == "3")
                    {
                        // 実績
                        fa[0] = nm + "!$AD$6";      // 系列名
                        fa[1] = nm + "!day";        // 軸ラベルの範囲
                        fa[2] = nm + "!Performance";   // 系列値
                    }
                    else if (fa[3] == "4")
                    {
                        // 目標
                        fa[0] = nm + "!$AD$7";      // 系列名
                        fa[1] = nm + "!day";        // 軸ラベルの範囲
                        fa[2] = nm + "!goal";       // 系列値
                    }

                    string newF = "=SERIES(" + fa[0] + "," + fa[1] + "," + fa[2] + "," + fa[3] + ")";
                    //msg += newF + Environment.NewLine;

                    sr.Formula = newF;
                }

                // シート全体を印刷対象とするためセルをアクティブとする 2018/02/09
                oxlsSheet.Range["A1"].Select();
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
            int i = 0;

            foreach (var t in zd.Where(a => (a.sHoliday == 0 || a.sZangyo > 0) && a.sDay <= sMaxDay).OrderBy(a => a.sDay))
            {
                t.sZissekibyDay = (v + t.sZangyo).ToString();
                
                //// 2017/10/25
                //zd[i].sZissekibyDay = (v + t.sZangyo).ToString();

                v += t.sZangyo;
                i++;
            }
        }

        ///------------------------------------------------------------------------------
        /// <summary>
        ///     残業計画時間の稼働日数割りと日々目標値を配列にセットする </summary>
        /// <param name="zd">
        ///     日別配列</param>
        /// <param name="zPlan">
        ///     月残業計画値</param>
        ///------------------------------------------------------------------------------
        private void setDaybyPlan(ref clsZanSum[] zd, double zPlan)
        {
            // 稼働日数を取得
            int kDays = zd.Count(a => a.sHoliday == 0);

            // 日々目標値
            double val = zPlan / (double)kDays;

            int i = 0;

            // 2018/01/17
            foreach (var t in zd.Where(a => a.sHoliday == 0 || a.sZangyo > 0).OrderBy(a => a.sDay))
            {
                if (t.sHoliday == 0)
                {
                    i++;
                }

                t.sMonthPlan = Utility.StrtoDouble((zPlan / (double)kDays * i).ToString("#,##0.0"));
                t.sPlanbyDay = val;

                // 2017/10/25
                //zd[i].sMonthPlan = Utility.StrtoDouble((zPlan / (double)kDays * (i + 1)).ToString("#,##0.0"));
                //zd[i].sPlanbyDay = val;

            }

            // 2018/01/17
            //foreach (var t in zd.Where(a => a.sHoliday == 0 || a.sZangyo > 0).OrderBy(a => a.sDay))
            //{
            //    t.sMonthPlan = Utility.StrtoDouble((zPlan / (double)kDays * (i + 1)).ToString("#,##0.0"));
            //    t.sPlanbyDay = val;

            //    // 2017/10/25
            //    //zd[i].sMonthPlan = Utility.StrtoDouble((zPlan / (double)kDays * (i + 1)).ToString("#,##0.0"));
            //    //zd[i].sPlanbyDay = val;
            //    i++;
            //}
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
        private void setDaybyZan(string bCode, ref clsZanSum[] z, ref decimal zTotal)
        {
            var context = new CsvContext();

            // CSVの情報を示すオブジェクトを構築
            var description = new CsvFileDescription
            {
                SeparatorChar = ',',
                FirstLineHasColumnNames = false,
                EnforceCsvColumnAttribute = true,
                TextEncoding = Encoding.GetEncoding(932)
            };
            
            var s = context.Read<Common.clsLinqCsv>(outZangyoFile, description)
                .Where(a => a.buCode.ToString() == bCode)
                .GroupBy(a => a.workDate.Day)
                .Select(a => new
                {
                    sday = a.Key,
                    zan = a.Sum(n => (n.zHayade + n.zFutsuu + n.zShinya + n.zKyuushutsu + n.zKyuShinya))
                });

            foreach (var t in s)
            {
                // 月間残業合計
                zTotal += t.zan;

                // 日別の残業時間を配列にセット
                for (int iZ = 0; iZ < z.Length; iZ++)
                {
                    if (z[iZ].sDay == t.sday)
                    {
                        z[iZ].sZangyo =  (double)t.zan; 
                        break;
                    }
                }
            }
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     任意の班・係・課の月間残業時間集計 ： 2018/04/13</summary>
        /// <param name="bCode">
        ///     班・係・課コード</param>
        /// <param name="zTotal">
        ///     残業時間月合計</param>
        ///----------------------------------------------------------------------
        private decimal setTotalZan(string bCode, string filePath)
        {
            decimal val = 0;

            var context = new CsvContext();

            // CSVの情報を示すオブジェクトを構築
            var description = new CsvFileDescription
            {
                SeparatorChar = ',',
                FirstLineHasColumnNames = false,
                EnforceCsvColumnAttribute = true,
                TextEncoding = Encoding.GetEncoding(932)
            };

            var s = context.Read<Common.clsLinqCsv>(filePath, description)
                //.Where(a => a.buCode.ToString() == bCode + "00")
                .Where(a => a.buCode.ToString() == bCode)
                .GroupBy(a => a.buCode)
                .Select(a => new
                {
                    sday = a.Key,
                    zan = a.Sum(n => (n.zHayade + n.zFutsuu + n.zShinya + n.zKyuushutsu + n.zKyuShinya))
                });

            foreach (var t in s)
            {
                val = t.zan;
            }

            return val;
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     任意の課の月間残業時間集計 </summary>
        /// <param name="bCode">
        ///     部門コード</param>
        /// <param name="zTotal">
        ///     残業時間月合計</param>
        ///----------------------------------------------------------------------
        private decimal setTotalZanBumon(string bCode, string filePath)
        {
            decimal val = 0;

            var context = new CsvContext();

            // CSVの情報を示すオブジェクトを構築
            var description = new CsvFileDescription
            {
                SeparatorChar = ',',
                FirstLineHasColumnNames = false,
                EnforceCsvColumnAttribute = true,
                TextEncoding = Encoding.GetEncoding(932)
            };

            var s = context.Read<Common.clsLinqCsv>(filePath, description)
                .GroupBy(a => a.buCode)
                .Select(a => new
                {
                    sday = a.Key,
                    zan = a.Sum(n => (n.zHayade + n.zFutsuu + n.zShinya + n.zKyuushutsu + n.zKyuShinya))
                });

            foreach (var t in s)
            {
                if (bCode == "1")
                {
                    // 製造部門
                    if (Utility.StrtoInt(t.sday.ToString().PadLeft(5, '0').Substring(1, 1)) == Utility.StrtoInt(bCode))
                    {
                        val += t.zan;
                    }
                }
                else if (bCode == "2")
                {
                    // 間接部門
                    if (Utility.StrtoInt(t.sday.ToString().PadLeft(5, '0').Substring(1, 1)) >= Utility.StrtoInt(bCode))
                    {
                        val += t.zan;
                    }
                }
                else if (bCode == "0")
                {
                    // 全社
                    if (Utility.StrtoInt(t.sday.ToString().PadLeft(5, '0').Substring(1, 1)) != global.flgOff)
                    {
                        val += t.zan;
                    }
                }
            }

            return val;
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
        private void setDaybyZanBumon(string bCode, ref clsZanSum[] z, ref decimal zTotal)
        {
            //IEnumerable<dayZan> d = getDaybyZanSum(bCode);

            var context = new CsvContext();

            // CSVの情報を示すオブジェクトを構築
            var description = new CsvFileDescription
            {
                SeparatorChar = ',',
                FirstLineHasColumnNames = false,
                EnforceCsvColumnAttribute = true,
                TextEncoding = Encoding.GetEncoding(932)
            };

            // 全社
            if (bCode == "0")
            {
                var s = context.Read<Common.clsLinqCsv>(outZangyoFile, description)
                    .Where(a => a.buCode.ToString().PadLeft(5, '0').Substring(1, 1) != "0")
                    .GroupBy(a => a.workDate.Day)
                    .Select(a => new
                    {
                        sday = a.Key,
                        zan = a.Sum(n => (n.zHayade + n.zFutsuu + n.zShinya + n.zKyuushutsu + n.zKyuShinya))
                    });

                foreach (var t in s)
                {
                    // 月間残業合計
                    zTotal += t.zan;

                    // 日別の残業時間を配列にセット
                    for (int iZ = 0; iZ < z.Length; iZ++)
                    {
                        if (z[iZ].sDay == t.sday)
                        {
                            z[iZ].sZangyo = (double)t.zan;
                            break;
                        }
                    }
                }
            }

            // 製造部門
            if (bCode == "1")
            {
                var s = context.Read<Common.clsLinqCsv>(outZangyoFile, description)
                    .Where(a => a.buCode.ToString().PadLeft(5, '0').Substring(1, 1) == bCode)
                    .GroupBy(a => a.workDate.Day)
                    .Select(a => new
                    {
                        sday = a.Key,
                        zan = a.Sum(n => (n.zHayade + n.zFutsuu + n.zShinya + n.zKyuushutsu + n.zKyuShinya))
                    });

                foreach (var t in s)
                {
                    // 月間残業合計
                    zTotal += t.zan;

                    // 日別の残業時間を配列にセット
                    for (int iZ = 0; iZ < z.Length; iZ++)
                    {
                        if (z[iZ].sDay == t.sday)
                        {
                            z[iZ].sZangyo = (double)t.zan;
                            break;
                        }
                    }
                }
            }

            // 間接部門
            if (bCode == "2")
            {
                var s = context.Read<Common.clsLinqCsv>(outZangyoFile, description)
                    .Where(a => a.buCode.ToString().PadLeft(5, '0').Substring(1, 1) != "0" && a.buCode.ToString().PadLeft(5, '0').Substring(1, 1) != "1")
                    .GroupBy(a => a.workDate.Day)
                    .Select(a => new
                    {
                        sday = a.Key,
                        zan = a.Sum(n => (n.zHayade + n.zFutsuu + n.zShinya + n.zKyuushutsu + n.zKyuShinya))
                    });

                foreach (var t in s)
                {
                    // 月間残業合計
                    zTotal += t.zan;

                    // 日別の残業時間を配列にセット
                    for (int iZ = 0; iZ < z.Length; iZ++)
                    {
                        if (z[iZ].sDay == t.sday)
                        {
                            z[iZ].sZangyo = (double)t.zan;
                            break;
                        }
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

            for (int iZ = 0; iZ < zz.GetLength(0); iZ++)
            {
                // 対象年月以外のときは対象外
                if (Utility.StrtoInt(zz[iZ, 1].ToString()) != yy ||
                    Utility.StrtoInt(zz[iZ, 2].ToString()) != mm)
                {
                    continue;
                }

                // 該当部署の当月計画値を取得する
                if (zz[iZ, 0].ToString() == bushoCode)
                {
                    zanPlan = Utility.StrtoDouble(Utility.NulltoStr(zz[iZ, 4]));
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

            for (int iZ = 0; iZ < zz.GetLength(0); iZ++)
            {
                // 対象年月以外のときは対象外
                if (Utility.StrtoInt(zz[iZ, 1].ToString()) != yy ||
                    Utility.StrtoInt(zz[iZ, 2].ToString()) != mm)
                {
                    continue;
                }

                // 部署コードの頭2桁目が部門（１：製造、2以上：間接）
                int bmn = Utility.StrtoInt(zz[iZ, 0].ToString().Substring(1, 1));

                // 全社または部門の当月計画値を取得する
                switch (bmnCode)
                {
                    case 0: // 全社
                        zanPlan += Utility.StrtoDouble(Utility.NulltoStr(zz[iZ, 4]));
                        break;

                    case 1: // 製造部門
                        if (bmn == bmnCode)
                        {
                            zanPlan += Utility.StrtoDouble(Utility.NulltoStr(zz[iZ, 4]));
                        }
                        break;

                    case 2: // 間接部門
                        if (bmn >= bmnCode)
                        {
                            zanPlan += Utility.StrtoDouble(Utility.NulltoStr(zz[iZ, 4]));
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
        /// <param name="keta">
        ///     集計桁数</param>
        ///-----------------------------------------------------------------------
        private void setZengetsuZan(ref string[,] zenArray, int keta)
        {
            var context = new CsvContext();

            // CSVの情報を示すオブジェクトを構築
            var description = new CsvFileDescription
            {
                SeparatorChar = ',',
                FirstLineHasColumnNames = false,
                EnforceCsvColumnAttribute = true,
                TextEncoding = Encoding.GetEncoding(932)
            };

            var s = context.Read<Common.clsLinqCsv>(outZengetsuFile, description)
                .GroupBy(a => a.buCode.ToString().PadLeft(5, '0').Substring(0, keta))
                .Select(a => new
                {
                    sKa = a.Key,
                    zan = a.Sum(n => (n.zHayade + n.zFutsuu + n.zShinya + n.zKyuushutsu + n.zKyuShinya))
                });

            // 前月残業実績配列を作成
            zenArray = new string[s.Count(), 2];

            int ix = 0;
            foreach (var t in s)
            {
                zenArray[ix, 0] = t.sKa;
                zenArray[ix, 1] = t.zan.ToString();
                ix++;
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
            for (int i = 0; i < zpArrayNew.GetLength(0); i++)
            {
                // 人数が0のときはネグる
                if (Utility.StrtoInt(zpArrayNew[i, 3].ToString()) == 0)
                {
                    continue;
                }

                if (bmnCode == "0")
                {
                    // 全社
                    if (Utility.StrtoInt(zpArrayNew[i, 1].ToString()) == yy &&
                        Utility.StrtoInt(zpArrayNew[i, 2].ToString()) == mm)
                    {
                        // 生産数、人数を取得
                        zNi += Utility.StrtoInt(zpArrayNew[i, 3].ToString());
                        zSei += Utility.StrtoInt(zpArrayNew[i, 5].ToString());
                    }
                }
                else if (bmnCode == "1")
                {
                    // 製造部門
                    if (zpArrayNew[i, 0].ToString().Substring(1, 1) == bmnCode &&
                        Utility.StrtoInt(zpArrayNew[i, 1].ToString()) == yy &&
                        Utility.StrtoInt(zpArrayNew[i, 2].ToString()) == mm)
                    {
                        // 生産数、人数を取得
                        zNi += Utility.StrtoInt(zpArrayNew[i, 3].ToString());
                        zSei += Utility.StrtoInt(zpArrayNew[i, 5].ToString());
                    }
                }
                else if (bmnCode == "2")
                {
                    // 間接部門
                    if (zpArrayNew[i, 0].ToString().Substring(1, 1) != "0" &&
                        zpArrayNew[i, 0].ToString().Substring(1, 1) != "1" &&
                        Utility.StrtoInt(zpArrayNew[i, 1].ToString()) == yy &&
                        Utility.StrtoInt(zpArrayNew[i, 2].ToString()) == mm)
                    {
                        // 生産数、人数を取得
                        zNi += Utility.StrtoInt(zpArrayNew[i, 3].ToString());
                        zSei += Utility.StrtoInt(zpArrayNew[i, 5].ToString());
                    }
                }
            }
        }

        ///--------------------------------------------------------------------
        /// <summary>
        ///     理由別残業時間 : 最大日付範囲で取得 2017/11/22</summary>
        /// <param name="bmnCode">
        ///     部門コード</param>
        /// <returns>
        ///     IEnumerable<riyuZan> </returns>
        ///--------------------------------------------------------------------
        private IEnumerable<riyuZan> getRiyuZanSum(string bmnCode)
        {
            EnumerableRowCollection<DataSet1.残業集計1Row> ss;

            // 集計
            if (bmnCode == "0")
            {
                // 最大日付範囲で取得 2017/11/22
                //ss = dts.残業集計1.Where(a => a.部署コード != "" && a.日 <= days);

                // 応援先部門を対象とする：2018/04/18
                ss = dts.残業集計1.Where(a => a.応援先 != "" && a.日 <= days);
            }
            else
            {
                // 最大日付範囲で取得 2017/11/22
                //ss = dts.残業集計1.Where(a => a.部署コード.Substring(1, 1) == bmnCode && a.日 <= days);

                // 応援先部門を対象とする：2018/04/18
                ss = dts.残業集計1.Where(a => a.応援先.Substring(1, 1) == bmnCode && a.日 <= days);
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
            EnumerableRowCollection<DataSet1.残業集計1Row> d;

            if (bmnCode == "0")
            {
                // 全社
                d = dts.残業集計1.Where(a => a.部署コード != ""); 
            }
            else
            {
                // 部門で絞込み
                d = dts.残業集計1.Where(a => a.部署コード.Substring(1, 1) == bmnCode);
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
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (label6.Text == string.Empty)
            {
                MessageBox.Show("前月勤務実績データを選択してください", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                button1.Focus();
                return;
            }

            if (label7.Text == string.Empty)
            {
                MessageBox.Show("当月勤務実績データを選択してください", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                button2.Focus();
                return;
            }

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

            // 前月取得
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

            // 当月勤務実績ファイル出力 2018/04/16
            if (rBtn5.Checked && chart_Status != global.CHART_HAN || 
                rBtn6.Checked && chart_Status != global.CHART_KAKARI || 
                rBtn4.Checked && chart_Status != global.CHART_KA || 
                rBtn2.Checked && chart_Status != global.CHART_KA || 
                rBtn3.Checked && chart_Status != global.CHART_KA)
            {
                // 当月勤務実績ファイル出力
                outWorkData(workArray, outZangyoFile, "当月", ref openCsvTou);

                // 残業振替データを当月勤務実績ファイルに追記 2018/02/10
                // 残業振替データを奉行からの残業実績データの対象日付範囲と同期を合わせる 2018/03/08

                // 班集計
                if (rBtn5.Checked)
                {
                    // 班：2018/04/09
                    putOuenFurikae_HAN(outZangyoFile, yy, mm, days);
                    chart_Status = global.CHART_HAN;        // 2018/04/16
                }

                // 係集計
                if (rBtn6.Checked)
                {
                    // 係：2018/04/09
                    putOuenFurikae_KAKARI(outZangyoFile, yy, mm, days);
                    chart_Status = global.CHART_KAKARI;     // 2018/04/16
                }
                
                // 課,部門,全社集計
                if (rBtn4.Checked || rBtn2.Checked || rBtn3.Checked)
                {
                    // 課：2018/04/09
                    putOuenFurikae(outZangyoFile, yy, mm, days);
                    chart_Status = global.CHART_KA;        // 2018/04/16
                }
            }

            //if (!openCsvTou)
            //{
            //    // 当月勤務実績ファイル出力
            //    outWorkData(workArray, outZangyoFile, "当月", ref openCsvTou);

            //    // 残業振替データを当月勤務実績ファイルに追記 2018/02/10
            //    // 残業振替データを奉行からの残業実績データの対象日付範囲と同期を合わせる 2018/03/08

            //    if (rBtn6.Checked)
            //    {
            //        // 係：2018/04/09
            //        putOuenFurikae_KAKARI(outZangyoFile, yy, mm, days);
            //    }
            //    else
            //    {
            //        putOuenFurikae(outZangyoFile, yy, mm, days);
            //    }
            //}

            // 前月勤務実績ファイル出力 2018/04/16
            if (rBtn5.Checked && chart_StatusZen != global.CHART_HAN ||
                rBtn6.Checked && chart_StatusZen != global.CHART_KAKARI ||
                rBtn4.Checked && chart_StatusZen != global.CHART_KA ||
                rBtn2.Checked && chart_StatusZen != global.CHART_KA ||
                rBtn3.Checked && chart_StatusZen != global.CHART_KA)
            {
                // 前月勤務実績ファイル出力
                outWorkData(zenArray, outZengetsuFile, "前月", ref openCsvZen);

                // 前月残業振替データを前月勤務実績ファイルに追記 2018/02/10
                // 前月残業振替データは1ヶ月分取得する 2018/03/08

                // 班集計
                if (rBtn5.Checked)
                {
                    // 班：2018/04/09
                    putOuenFurikae_HAN(outZengetsuFile, zYY, zMM, days);
                    chart_StatusZen = global.CHART_HAN;        // 2018/04/16
                }

                // 係集計
                if (rBtn6.Checked)
                {
                    // 係：2018/04/09
                    putOuenFurikae_KAKARI(outZengetsuFile, zYY, zMM, days);
                    chart_StatusZen = global.CHART_KAKARI;     // 2018/04/16
                }

                // 課,部門,全社集計
                if (rBtn4.Checked || rBtn2.Checked || rBtn3.Checked)
                {
                    // 班：2018/04/09
                    putOuenFurikae(outZengetsuFile, zYY, zMM, days);
                    chart_StatusZen = global.CHART_KA;        // 2018/04/16
                }
            }

            //if (!openCsvZen)
            //{
            //    // 前月勤務実績ファイル出力
            //    outWorkData(zenArray, outZengetsuFile, "前月", ref openCsvZen);

            //    // 前月残業振替データを前月勤務実績ファイルに追記 2018/02/10
            //    // 前月残業振替データは1ヶ月分取得する 2018/03/08
            //    if (rBtn6.Checked)
            //    {
            //        // 係：2018/04/09
            //        putOuenFurikae_KAKARI(outZangyoFile, yy, mm, days);
            //    }
            //    else
            //    {
            //        putOuenFurikae(outZangyoFile, yy, mm, days);
            //    }
            //}

            // 残業推移チャート作成
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

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "勤怠データ選択";
            openFileDialog1.FileName = string.Empty;
            openFileDialog1.Filter = "ＣＳＶファイル(*.csv)|*.csv|全てのファイル(*.*)|*.*";

            //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
            string fileName;
            DialogResult ret = openFileDialog1.ShowDialog();

            if (ret == System.Windows.Forms.DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
                label7.Text = openFileDialog1.FileName;

                // 勤怠データ配列読み込み
                workArray = System.IO.File.ReadAllLines(label7.Text, Encoding.Default);

                getWorkDateDays(out yy, out mm, out days);

                txtYear.Text = yy.ToString();
                txtMonth.Text = mm.ToString();

                openCsvTou = false;
                chart_Status = global.CHART_NOTHING; 
            }
            else
            {
                fileName = string.Empty;
            }
        }


        ///-----------------------------------------------------------------
        /// <summary>
        ///     CSVデータの対象年月と最大日付を取得する </summary>
        /// <param name="_yy">
        ///     年</param>
        /// <param name="_mm">
        ///     月</param>
        /// <param name="_days">
        ///     最大日付</param>
        /// <returns>
        ///     true, false</returns>
        ///-----------------------------------------------------------------
        private bool getWorkDateDays(out int _yy, out int _mm, out int _days)
        {
            bool rtn = false;
            _yy = 0;
            _mm = 0;
            _days = 0;

            string sNum = string.Empty;

            foreach (var item in workArray)
            {
                string[] t = item.Split(',');

                // 社員番号取得
                string strSnum = t[0].Replace("\"", "");

                // 1行目見出し行は読み飛ばす
                if (strSnum == "社員番号")
                {
                    continue;
                }

                if (sNum != string.Empty && sNum != strSnum)
                {
                    rtn = true;
                    break;
                }

                sNum = strSnum;

                // 日付情報を取得する
                string strDate = t[2].Replace("\"", "");
                //string f = "ggyy年MM月dd日";
                System.Globalization.CultureInfo ci = new System.Globalization.CultureInfo("ja-JP");
                DateTime iDt = DateTime.Parse(strDate, ci, System.Globalization.DateTimeStyles.AssumeLocal);

                _yy = iDt.Year;
                _mm = iDt.Month;
                _days = iDt.Day;
            }

            return rtn;
        }

        ///-----------------------------------------------------------------
        /// <summary>
        ///     所属コードを付加したCSVデータを出力する </summary>
        ///-----------------------------------------------------------------
        private void outWorkData(string [] inCsv, string outCsv, string msg, ref bool sts)
        {
            // 奉行データベース接続
            string sc = sqlControl.obcConnectSting.get(_dbName);
            sqlControl.DataControl sdCon = new Common.sqlControl.DataControl(sc);

            // progressBar
            int nMax = inCsv.Length;
            toolStripProgressBar1.Maximum = nMax;
            toolStripProgressBar1.Minimum = 0;
            toolStripProgressBar1.Visible = true;
            int iRow = 0;

            label1.Visible = true;
            label1.Text = msg + "勤務実績データの所属情報を取得中です...";
            this.Refresh(); // ← 追加

            try
            {
                Cursor = Cursors.WaitCursor;

                string[] sss = null;
                int sCnt = 0;

                foreach (var item in inCsv)
                {
                    string[] t = item.Split(',');

                    // 社員番号取得
                    string strSnum = t[0].Replace("\"", "");

                    // 1行目見出し行は読み飛ばす
                    if (strSnum == "社員番号")
                    {
                        continue;
                    }

                    // progressBar表示
                    iRow++;
                    //label1.Text = "勤務実績データに所属情報を取得中..." + iRow + "/" + nMax;
                    toolStripProgressBar1.Value = iRow;
                    //this.Refresh(); // ← 追加

                    string sNum = strSnum.PadLeft(10, '0');

                    // 日付情報を取得する
                    string strDate = t[2].Replace("\"", "");
                    System.Globalization.CultureInfo ci = new System.Globalization.CultureInfo("ja-JP");
                    DateTime iDt = DateTime.Parse(strDate, ci, System.Globalization.DateTimeStyles.AssumeLocal);

                    string dCode = getDepartmentCode(sdCon, sNum, iDt.ToShortDateString());

                    if (Utility.StrtoInt(dCode) == global.flgOff)
                    {
                        // 1111N等の所属コードへの対応 : 2017/10/25
                        dCode = dCode.Trim().PadLeft(15, '0');
                    }

                    if (dCode != string.Empty)
                    {
                        // 新たなCSVデータを出力
                        StringBuilder sb = new StringBuilder();

                        if (rBtn4.Checked)
                        {
                            // 課別
                            sb.Append(dCode.Substring(10, 3) + "00").Append(",");
                        }
                        else if (rBtn2.Checked)
                        {
                            // 製造部門・間接
                            sb.Append(dCode.Substring(10, 3) + "00").Append(",");
                        }
                        else if (rBtn3.Checked)
                        {
                            // 全社
                            sb.Append(dCode.Substring(10, 3) + "00").Append(",");
                        }
                        else if (rBtn6.Checked)
                        {
                            // 係 : 2018/04/09
                            sb.Append(dCode.Substring(10, 4) + "0").Append(",");
                        }
                        else if (rBtn5.Checked)
                        {
                            // 班 : 2018/04/16
                            sb.Append(dCode.Substring(10, 5)).Append(",");
                        }

                        sb.Append(strSnum).Append(",");
                        sb.Append(iDt.ToShortDateString()).Append(",");
                        sb.Append(Utility.StrtoDouble(t[33].Replace("\"", ""))).Append(",");     // 早出残業
                        sb.Append(Utility.StrtoDouble(t[36].Replace("\"", ""))).Append(",");     // 普通残業
                        sb.Append(Utility.StrtoDouble(t[39].Replace("\"", ""))).Append(",");     // 深夜残業
                        sb.Append(Utility.StrtoDouble(t[42].Replace("\"", ""))).Append(",");     // 休出残業
                        sb.Append(Utility.StrtoDouble(t[45].Replace("\"", "")));                 // 休出深残
                        sb.Append(",S"); // 実績データ記号[S] 2018/02/10

                        // 配列にデータを出力
                        sCnt++;
                        Array.Resize(ref sss, sCnt);
                        sss[sCnt - 1] = sb.ToString();
                    }
                }

                txtFileWrite(outCsv, sss);
                sts = true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                if (sdCon.Cn.State == ConnectionState.Open)
                {
                    sdCon.Close();
                }

                Cursor = Cursors.Default;

                // プログレスバーを非表示
                toolStripProgressBar1.Visible = false;
                label1.Visible = false;
            }
        }

        ///----------------------------------------------------------------------------
        /// <summary>
        ///     テキストファイルを出力する</summary>
        /// <param name="outFilePath">
        ///     出力するフォルダ</param>
        /// <param name="arrayData">
        ///     書き込む配列データ</param>
        ///----------------------------------------------------------------------------
        private void txtFileWrite(string sPath, string[] arrayData)
        {
            // 同名ファイルがあったら削除する
            if (System.IO.File.Exists(sPath))
            {
                System.IO.File.Delete(sPath);
            }

            // テキストファイル出力
            System.IO.File.WriteAllLines(sPath, arrayData, System.Text.Encoding.GetEncoding(932));
        }


        private string getDepartmentCode(sqlControl.DataControl sdCon, string sNum, string sDt)
        {
            SqlDataReader dR = null;
            string sDepartmentCode = string.Empty;

            try
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("select tbEmployeeBase.EmployeeID, a.DepartmentCode, a.DepartmentName ");
                sb.Append(" from tbEmployeeBase inner join ");
                sb.Append("(select d.EmployeeID, DepartmentCode, DepartmentName from tbDepartment inner join ");
                sb.Append("((select tbEmployeeMainDutyPersonnelChange.EmployeeID, BelongID ");
                sb.Append("from tbEmployeeMainDutyPersonnelChange inner join ");
                sb.Append("(select EmployeeID, max(AnnounceDate) as AnnounceDate ");
                sb.Append(" from tbEmployeeMainDutyPersonnelChange ");
                sb.Append("where AnnounceDate <= '" + sDt + "'");
                sb.Append("group by EmployeeID) as s ");
                sb.Append("on(tbEmployeeMainDutyPersonnelChange.EmployeeID = s.EmployeeID) and ");
                sb.Append("(tbEmployeeMainDutyPersonnelChange.AnnounceDate = s.AnnounceDate))) as d ");
                sb.Append("on tbDepartment.DepartmentID = d.BelongID) as a ");
                sb.Append("on tbEmployeeBase.EmployeeID = a.EmployeeID ");
                sb.Append("where EmployeeNo = '" + sNum + "'");

                dR = sdCon.free_dsReader(sb.ToString());

                while (dR.Read())
                {
                    sDepartmentCode = dR["DepartmentCode"].ToString().Trim();
                    break;
                }

                dR.Close();

                return sDepartmentCode;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return sDepartmentCode;
            }
            finally
            {
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            openFileDialog1.Title = "前月勤怠データ選択";
            openFileDialog1.FileName = string.Empty;
            openFileDialog1.Filter = "ＣＳＶファイル(*.csv)|*.csv|全てのファイル(*.*)|*.*";

            //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
            string fileName;
            DialogResult ret = openFileDialog1.ShowDialog();

            if (ret == System.Windows.Forms.DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
                label6.Text = openFileDialog1.FileName;

                // 前月勤怠データ配列読み込み
                zenArray = System.IO.File.ReadAllLines(label6.Text, Encoding.Default);

                openCsvZen = false;
                chart_StatusZen = global.CHART_NOTHING; 
            }
            else
            {
                fileName = string.Empty;
            }
        }

        private void rBtn4_CheckedChanged(object sender, EventArgs e)
        {
            if (rBtn4.Checked)
            {
                comboBox2.Enabled = true;
                comboBox2.Text = string.Empty;

                // 部署コンボボックスにデータソース・課名をセットする
                label2.Text = "課名";
                Utility.ComboBumon.loadKa(comboBox2, _dbName, global.CHART_KA);
            }
            else
            {
                comboBox2.Enabled = false;
            }
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     部署別残業計画.xlsxをCSV出力する </summary>
        ///----------------------------------------------------------
        private void kaZanPlanSet()
        {
            string[] sss = null;
            int x = 0;

            for (int i = 1; i <= bs.zpArray.GetLength(0); i++)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(Utility.NulltoStr(bs.zpArray[i, 1])).Append(",");
                sb.Append(Utility.StrtoInt(Utility.NulltoStr(bs.zpArray[i, 2])).ToString()).Append(",");
                sb.Append(Utility.StrtoInt(Utility.NulltoStr(bs.zpArray[i, 3])).ToString()).Append(",");
                sb.Append(Utility.StrtoInt(Utility.NulltoStr(bs.zpArray[i, 4])).ToString()).Append(",");
                sb.Append(Utility.StrtoDouble(Utility.NulltoStr(bs.zpArray[i, 5])).ToString()).Append(",");
                sb.Append(Utility.StrtoInt(Utility.NulltoStr(bs.zpArray[i, 6])));

                Array.Resize(ref sss, x + 1);
                sss[x] = sb.ToString();
                x++;
            }

            txtFileWrite(outKaZanPlanFile, sss);
        }
        
        ///----------------------------------------------------------
        /// <summary>
        ///     部署別理由別残業計画.xlsxをCSV出力する 
        ///     ：2018/04/11</summary>
        ///----------------------------------------------------------
        private void kaByreByZanPlanSet()
        {
            string[] sss = null;
            int x = 0;

            for (int i = 1; i <= bs.zrpArray.GetLength(0); i++)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(Utility.NulltoStr(bs.zrpArray[i, 1])).Append(",");
                sb.Append(Utility.NulltoStr(bs.zrpArray[i, 2])).Append(",");
                sb.Append(Utility.NulltoStr(bs.zrpArray[i, 3])).Append(",");
                sb.Append(Utility.NulltoStr(bs.zrpArray[i, 4])).Append(",");
                sb.Append(Utility.NulltoStr(bs.zrpArray[i, 5])).Append(",");
                sb.Append(Utility.StrtoDouble(Utility.NulltoStr(bs.zrpArray[i, 6])).ToString());

                Array.Resize(ref sss, x + 1);
                sss[x] = sb.ToString();
                x++;
            }

            txtFileWrite(outKaByReByZanPlanFile, sss);
        }

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     部署別理由別残業計画.csv(kaByreByZanPlan.csv)から
        ///     班、係、課毎に残業計画時間を集計した配列を作成する 
        ///     ：2018/04/11</summary>
        ///-----------------------------------------------------------------------
        private void kaByreByZanPlanSum(int keta)
        {
            var context = new CsvContext();

            // CSVの情報を示すオブジェクトを構築
            var description = new CsvFileDescription
            {
                SeparatorChar = ',',
                FirstLineHasColumnNames = false,
                EnforceCsvColumnAttribute = true,
                TextEncoding = Encoding.GetEncoding(932)
            };

            var s = context.Read<Common.clsLinqkaByreByZan>(outKaByReByZanPlanFile, description)
                .GroupBy(a => new { ka = a.buCode.ToString().Substring(0, keta), a.zYear, a.zMonth, a.zReason })
                .Select(n => new
                {
                    sKa = n.Key.ka,
                    sYear = n.Key.zYear,
                    sMonth = n.Key.zMonth,
                    sReason = n.Key.zReason,
                    sZan = n.Sum(k => Utility.StrtoDouble(Utility.NulltoStr(k.zZanPlan)))
                })
                .OrderBy(a => a.sYear).ThenBy(a => a.sMonth).ThenBy(a => a.sKa);

            int x = 0;

            szByreByzanPlan = new string[s.Count()];

            foreach (var t in s)
            {
                szByreByzanPlan[x] = t.sKa + "," + t.sYear + "," + t.sMonth + "," + t.sReason + "," + t.sZan;
                x++;
            }
        }

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     部署別残業計画から課別年月毎に人数・残業計画・生産数を集計した
        ///     2次元配列を作成する </summary>
        ///-----------------------------------------------------------------------
        private void kaZanPlanSum(int keta)
        {
            var context = new CsvContext();

            // CSVの情報を示すオブジェクトを構築
            var description = new CsvFileDescription
            {
                SeparatorChar = ',',
                FirstLineHasColumnNames = false,
                EnforceCsvColumnAttribute = true,
                TextEncoding = Encoding.GetEncoding(932)
            };

            var s = context.Read<Common.clsLinqZan>(outKaZanPlanFile, description)
                .GroupBy(a => new {ka = a.buCode.ToString().Substring(0, keta), a.zYear, a.zMonth})
                .Select(n => new
                {
                    sKa = n.Key.ka,
                    sYear = n.Key.zYear,
                    sMonth = n.Key.zMonth,
                    sNin = n.Sum(k => Utility.StrtoDouble(Utility.NulltoStr(k.zNin))),
                    sZan = n.Sum(k => Utility.StrtoDouble(Utility.NulltoStr(k.zZanPlan))),
                    sSeisan = n.Sum(k => Utility.StrtoDouble(Utility.NulltoStr(k.zSeisan)))
                })
                .OrderBy(a => a.sYear).ThenBy(a => a.sMonth).ThenBy(a => a.sKa);

            int x = 0;

            zpArrayNew = new string[s.Count(), 6];

            foreach (var t in s)
            {
                zpArrayNew[x, 0] = t.sKa;
                zpArrayNew[x, 1] = t.sYear.ToString();
                zpArrayNew[x, 2] = t.sMonth.ToString();
                zpArrayNew[x, 3] = t.sNin.ToString();
                zpArrayNew[x, 4] = t.sZan.ToString();
                zpArrayNew[x, 5] = t.sSeisan.ToString();
                x++;
            }
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     部門別残業振替データ作成・班 : 2018/04/16</summary>
        /// <param name="outPath">
        ///     当月残業ファイルパス</param>
        /// <param name="yy">
        ///     対象年</param>
        /// <param name="mm">
        ///     対象月</param>
        /// <param name="maxDay">
        ///     対象日付範囲</param>
        ///------------------------------------------------------------------
        private void putOuenFurikae_HAN(string outPath, int yy, int mm, int maxDay)
        {
            string[] csvArray = null;
            int iX = 0;

            Cursor = Cursors.WaitCursor;

            // 2018/02/10
            // 該当月データ取得
            DataSet1TableAdapters.残業集計3TableAdapter zAdp = new DataSet1TableAdapters.残業集計3TableAdapter();
            zAdp.Fill(dts.残業集計3, yy, mm, yy, mm, yy, mm, yy, mm, yy, mm, yy, mm);

            // nullに「０」をセット
            foreach (var item in dts.残業集計3)
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

            label1.Text = "応援先残業の振替データを作成中です...";

            try
            {
                // 奉行からの残業実績データの対象日付範囲と同期を合わせる　2018/03/08
                foreach (var t in dts.残業集計3.Where(a => a.日 <= maxDay && a.区分 == 2))
                {
                    // 応援元(振替元）残業時間マイナス用
                    StringBuilder sb = new StringBuilder();
                    sb.Append(t.部署コード).Append(",");
                    sb.Append(t.社員番号.ToString()).Append(",");
                    sb.Append(yy + "/" + mm + "/");
                    sb.Append(t.日.ToString()).Append(",");
                    sb.Append("0,");
                    sb.Append("-").Append(t.残業時 + "." + t.残業分).Append(",");
                    sb.Append("0,0,0,F");

                    Array.Resize(ref csvArray, iX + 1);
                    csvArray[iX] = sb.ToString();
                    iX++;

                    // 応援先(振替先）残業時間加算用
                    sb.Clear();
                    sb.Append(t.応援先).Append(",");
                    sb.Append(t.社員番号.ToString()).Append(",");
                    sb.Append(yy + "/" + mm + "/");
                    sb.Append(t.日.ToString()).Append(",");
                    sb.Append("0,");
                    sb.Append(t.残業時 + "." + t.残業分).Append(",");
                    sb.Append("0,0,0,F");

                    Array.Resize(ref csvArray, iX + 1);
                    csvArray[iX] = sb.ToString();
                    iX++;
                }

                if (csvArray != null)
                {
                    // 残業実績ファイルに追記する
                    System.IO.File.AppendAllLines(outPath, csvArray);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }
        

        ///------------------------------------------------------------------
        /// <summary>
        ///     部門別残業振替データ作成・係 : 2018/04/09</summary>
        /// <param name="outPath">
        ///     当月残業ファイルパス</param>
        /// <param name="yy">
        ///     対象年</param>
        /// <param name="mm">
        ///     対象月</param>
        /// <param name="maxDay">
        ///     対象日付範囲</param>
        ///------------------------------------------------------------------
        private void putOuenFurikae_KAKARI(string outPath, int yy, int mm, int maxDay)
        {
            string[] csvArray = null;
            int iX = 0;

            Cursor = Cursors.WaitCursor;

            // 2018/02/10
            // 該当月データ取得
            DataSet1TableAdapters.残業集計2TableAdapter zAdp = new DataSet1TableAdapters.残業集計2TableAdapter();
            zAdp.Fill(dts.残業集計2, yy, mm, yy, mm, yy, mm, yy, mm, yy, mm, yy, mm);

            // nullに「０」をセット
            foreach (var item in dts.残業集計2)
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

            label1.Text = "応援先残業の振替データを作成中です...";

            try
            {
                // 奉行からの残業実績データの対象日付範囲と同期を合わせる　2018/03/08
                foreach (var t in dts.残業集計2.Where(a => a.日 <= maxDay && a.区分 == 2))
                {
                    // 応援元(振替元）残業時間マイナス用
                    StringBuilder sb = new StringBuilder();
                    sb.Append(t.部署コード).Append(",");
                    sb.Append(t.社員番号.ToString()).Append(",");
                    sb.Append(yy + "/" + mm + "/");
                    sb.Append(t.日.ToString()).Append(",");
                    sb.Append("0,");
                    sb.Append("-").Append(t.残業時 + "." + t.残業分).Append(",");
                    sb.Append("0,0,0,F");

                    Array.Resize(ref csvArray, iX + 1);
                    csvArray[iX] = sb.ToString();
                    iX++;

                    // 応援先(振替先）残業時間加算用
                    sb.Clear();
                    sb.Append(t.応援先).Append(",");
                    sb.Append(t.社員番号.ToString()).Append(",");
                    sb.Append(yy + "/" + mm + "/");
                    sb.Append(t.日.ToString()).Append(",");
                    sb.Append("0,");
                    sb.Append(t.残業時 + "." + t.残業分).Append(",");
                    sb.Append("0,0,0,F");

                    Array.Resize(ref csvArray, iX + 1);
                    csvArray[iX] = sb.ToString();
                    iX++;
                }

                if (csvArray != null)
                {
                    // 残業実績ファイルに追記する
                    System.IO.File.AppendAllLines(outPath, csvArray);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     部門別残業振替データ作成 : 2018/02/10</summary>
        /// <param name="outPath">
        ///     当月残業ファイルパス</param>
        /// <param name="yy">
        ///     対象年</param>
        /// <param name="mm">
        ///     対象月</param>
        /// <param name="maxDay">
        ///     対象日付範囲</param>
        ///------------------------------------------------------------------
        private void putOuenFurikae(string outPath, int yy, int mm, int maxDay)
        {
            string[] csvArray = null;
            int iX = 0;

            Cursor = Cursors.WaitCursor;

            // 2018/02/10
            // 該当月データ取得
            DataSet1TableAdapters.残業集計1TableAdapter zAdp = new DataSet1TableAdapters.残業集計1TableAdapter();
            zAdp.Fill(dts.残業集計1, yy, mm, yy, mm, yy, mm, yy, mm, yy, mm, yy, mm);

            // nullに「０」をセット
            foreach (var item in dts.残業集計1)
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

            label1.Text = "応援先残業の振替データを作成中です...";

            try
            {
                //foreach (var t in dts.残業集計1.Where(a => a.区分 == 2)) 2018/03/08
                // 奉行からの残業実績データの対象日付範囲と同期を合わせる　2018/03/08
                foreach (var t in dts.残業集計1.Where(a => a.日 <= maxDay && a.区分 == 2))
                {
                    // 応援元(振替元）残業時間マイナス用
                    StringBuilder sb = new StringBuilder();
                    sb.Append(t.部署コード).Append(",");
                    sb.Append(t.社員番号.ToString()).Append(",");
                    sb.Append(yy + "/" + mm + "/");
                    sb.Append(t.日.ToString()).Append(",");
                    sb.Append("0,");
                    sb.Append("-").Append(t.残業時 + "." + t.残業分).Append(",");
                    sb.Append("0,0,0,F");
                    
                    Array.Resize(ref csvArray, iX + 1);
                    csvArray[iX] = sb.ToString();
                    iX++;

                    // 応援先(振替先）残業時間加算用
                    sb.Clear();
                    sb.Append(t.応援先).Append(",");
                    sb.Append(t.社員番号.ToString()).Append(",");
                    sb.Append(yy + "/" + mm + "/");
                    sb.Append(t.日.ToString()).Append(",");
                    sb.Append("0,");
                    sb.Append(t.残業時 + "." + t.残業分).Append(",");
                    sb.Append("0,0,0,F");

                    Array.Resize(ref csvArray, iX + 1);
                    csvArray[iX] = sb.ToString();
                    iX++;
                }

                if (csvArray != null)
                {
                    // 残業実績ファイルに追記する
                    System.IO.File.AppendAllLines(outPath, csvArray);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void rBtn5_CheckedChanged(object sender, EventArgs e)
        {
            if (rBtn5.Checked)
            {
                comboBox2.Enabled = true;
                comboBox2.Text = string.Empty;

                // 部署コンボボックスにデータソース・班名をセットする
                label2.Text = "班名";
                Utility.ComboBumon.loadKa(comboBox2, _dbName, global.CHART_HAN);
            }
            else
            {
                comboBox2.Enabled = false;
            }
        }

        private void rBtn6_CheckedChanged(object sender, EventArgs e)
        {
            if (rBtn6.Checked)
            {
                comboBox2.Enabled = true;
                comboBox2.Text = string.Empty;

                // 部署コンボボックスにデータソース・係名をセットする
                label2.Text = "係名";
                Utility.ComboBumon.loadKa(comboBox2, _dbName, global.CHART_KAKARI);
            }
            else
            {
                comboBox2.Enabled = false;
            }
        }

        private void rBtn2_CheckedChanged(object sender, EventArgs e)
        {
            if (rBtn2.Checked)
            {
                comboBox2.Enabled = false;
                comboBox2.Text = string.Empty;
            }
            else
            {
                //comboBox2.Enabled = false;
            }
        }

        private void rBtn3_CheckedChanged(object sender, EventArgs e)
        {
            if (rBtn3.Checked)
            {
                comboBox2.Enabled = false;
                comboBox2.Text = string.Empty;
            }
            else
            {
                //comboBox2.Enabled = false;
            }
        }
    }
}
