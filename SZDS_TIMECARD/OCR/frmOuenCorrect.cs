using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using SZDS_TIMECARD.Common;
using SZDS_TIMECARD.OCR;
using GrapeCity.Win.MultiRow;
using Excel = Microsoft.Office.Interop.Excel;

namespace SZDS_TIMECARD.OCR
{
    public partial class frmOuenCorrect : Form
    {
        public frmOuenCorrect(string dbName, string comName, string sID, bool eMode)
        {
            InitializeComponent();

            _dbName = dbName;       // データベース名
            _comName = comName;     // 会社名
            dID = sID;              // 処理モード
            _eMode = eMode;         // 処理モード２　勤怠データＩ／Ｐ票画面からの呼出し            

            // 処理モードの判断
            if (dID == string.Empty)    // データ編集モード
            {
                //// 画像パス取得
                //global.pblImagePath = Properties.Settings.Default.dataPathOuen;
            }

            // テーブルアダプターマネージャーに応援移動票ヘッダ、応援移動票明細テーブルアダプターを割り付ける
            adpMn.応援移動票ヘッダTableAdapter = hAdp;
            adpMn.応援移動票明細TableAdapter = iAdp;

            // 休日テーブル読み込み
            kAdp.Fill(dts.休日);
        }

        // データアダプターオブジェクト
        DataSet1TableAdapters.TableAdapterManager adpMn = new DataSet1TableAdapters.TableAdapterManager();
        DataSet1TableAdapters.応援移動票ヘッダTableAdapter hAdp = new DataSet1TableAdapters.応援移動票ヘッダTableAdapter();
        DataSet1TableAdapters.応援移動票明細TableAdapter iAdp = new DataSet1TableAdapters.応援移動票明細TableAdapter();
        DataSet1TableAdapters.休日TableAdapter kAdp = new DataSet1TableAdapters.休日TableAdapter();

        DataSet1TableAdapters.勤務票ヘッダTableAdapter iphAdp = new DataSet1TableAdapters.勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.勤務票明細TableAdapter ipmAdp = new DataSet1TableAdapters.勤務票明細TableAdapter();

        // データセットオブジェクト
        DataSet1 dts = new DataSet1();

        #region 終了ステータス定数
        const string END_BUTTON = "btn";
        const string END_MAKEDATA = "data";
        const string END_CONTOROL = "close";
        const string END_NODATA = "non Data";
        #endregion

        string dID = string.Empty;                  // 表示する過去データのID
        string sDBNM = string.Empty;                // データベース名

        string _dbName = string.Empty;           // 会社領域データベース識別番号
        string _comNo = string.Empty;            // 会社番号
        string _comName = string.Empty;          // 会社名
        bool _eMode = true;                      // 勤怠データＩ／Ｐ票画面からの呼出し

        // dataGridView1_CellEnterステータス
        bool gridViewCellEnterStatus = true;

        // 部署別勤務体系配列クラス
        xlsData bs;

        // ライン・部門・製品群コード配列取得 
        string[] hArray = null;

        // カレントデータRowsインデックス
        string[] cID = null;
        int cI = 0;

        // グローバルクラス
        global gl = new global();

        private void frmOuenCorrectcs_Load(object sender, EventArgs e)
        {
            this.pictureBox1.Image = new Bitmap(pictureBox1.Width, pictureBox1.Height);

            // フォーム最大値
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小値
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // Tabキーの既定のショートカットキーを解除する。
            gcMultiRow1.ShortcutKeyManager.Unregister(Keys.Tab);
            gcMultiRow2.ShortcutKeyManager.Unregister(Keys.Tab);
            gcMultiRow3.ShortcutKeyManager.Unregister(Keys.Tab);
            gcMultiRow1.ShortcutKeyManager.Unregister(Keys.Enter);
            gcMultiRow2.ShortcutKeyManager.Unregister(Keys.Enter);
            gcMultiRow3.ShortcutKeyManager.Unregister(Keys.Enter);

            // Tabキーのショートカットキーにユーザー定義のショートカットキーを割り当てる。
            gcMultiRow1.ShortcutKeyManager.Register(new clsKeyTab.CustomMoveToNextContorol(), Keys.Tab);
            gcMultiRow2.ShortcutKeyManager.Register(new clsKeyTab.CustomMoveToNextContorol(), Keys.Tab);
            gcMultiRow3.ShortcutKeyManager.Register(new clsKeyTab.CustomMoveToNextContorol(), Keys.Tab);
            gcMultiRow1.ShortcutKeyManager.Register(new clsKeyTab.CustomMoveToNextContorol(), Keys.Enter);
            gcMultiRow2.ShortcutKeyManager.Register(new clsKeyTab.CustomMoveToNextContorol(), Keys.Enter);
            gcMultiRow3.ShortcutKeyManager.Register(new clsKeyTab.CustomMoveToNextContorol(), Keys.Enter);
            
            // 勤務データ登録
            if (dID == string.Empty)
            {
                // CSVデータをMDBへ読み込みます
                GetCsvDataToMDB();

                // データセットへデータを読み込みます
                getDataSet();
                getOuenDataSet();

                // データテーブル件数カウント
                if (dts.応援移動票ヘッダ.Count == 0)
                {
                    MessageBox.Show("応援移動票データがありません", "応援移動票データ登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    //終了処理
                    Environment.Exit(0);
                }

                // キー配列作成
                keyArrayCreate();
            }

            // 部署別残業理由シートデータ配列取得
            bs = new xlsData();
            bs.zArray = bs.getShiftCode(string.Empty);

            // 部署別残業理由シートデータ配列取得
            bs.rArray = bs.getZanReason();

            // ライン・部門・製品群コード配列取得 
            hArray = Utility.getCategoryArray(_dbName);

            // キャプション
            this.Text = "応援移動票データ表示";

            // GCMultiRow初期化
            gcMrSetting();

            // 編集作業、過去データ表示の判断
            if (dID == string.Empty) // パラメータのヘッダIDがないときは編集作業
            {
                // 最初のレコードを表示
                cI = 0;
                showOcrData(cI);
            }

            // tagを初期化
            this.Tag = string.Empty;

            // 現在の表示倍率を初期化
            gl.miMdlZoomRate = 0f;
        }

        ///----------------------------------------------------------------------------
        /// <summary>
        ///     CSVデータをMDBへインサートする</summary>
        ///----------------------------------------------------------------------------
        private void GetCsvDataToMDB()
        {
            // CSVファイル数をカウント
            string[] inCsv = System.IO.Directory.GetFiles(Properties.Settings.Default.dataPathOuen, "*.csv");

            // CSVファイルがなければ終了
            if (inCsv.Length == 0) return;

            // オーナーフォームを無効にする
            this.Enabled = false;

            //プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            // OCRのCSVデータをMDBへ取り込む
            OCRData ocr = new OCRData(_dbName, bs);
            ocr.csvToMdbOuen(Properties.Settings.Default.dataPathOuen, frmP, _dbName);

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     キー配列作成 </summary>
        ///-------------------------------------------------------------
        private void keyArrayCreate()
        {
            int iX = 0;
            foreach (var t in dts.応援移動票ヘッダ.OrderBy(a => a.ID))
            {
                Array.Resize(ref cID, iX + 1);
                cID[iX] = t.ID;
                iX++;
            }
        }

        private void gcMrSetting()
        {
            // 年月・部署コード
            gcMultiRow1.EditMode = EditMode.EditProgrammatically;

            this.gcMultiRow1.AllowUserToAddRows = false;                    // 手動による行追加を禁止する
            this.gcMultiRow1.AllowUserToDeleteRows = false;                 // 手動による行削除を禁止する
            this.gcMultiRow1.Rows.Clear();                                  // 行数をクリア
            this.gcMultiRow1.RowCount = 1;                                  // 行数を設定
            this.gcMultiRow1.HideSelection = true;                          // GcMultiRow コントロールがフォーカスを失ったとき、セルの選択状態を非表示にする


            // 日中応援
            gcMultiRow2.EditMode = EditMode.EditProgrammatically;

            this.gcMultiRow2.AllowUserToAddRows = false;                    // 手動による行追加を禁止する
            this.gcMultiRow2.AllowUserToDeleteRows = false;                 // 手動による行削除を禁止する
            this.gcMultiRow2.Rows.Clear();                                  // 行数をクリア
            this.gcMultiRow2.RowCount = 5;                                  // 行数を設定
            this.gcMultiRow2.HideSelection = true;                          // GcMultiRow コントロールがフォーカスを失ったとき、セルの選択状態を非表示にする


            // 残業応援
            gcMultiRow3.EditMode = EditMode.EditProgrammatically;

            this.gcMultiRow3.AllowUserToAddRows = false;                    // 手動による行追加を禁止する
            this.gcMultiRow3.AllowUserToDeleteRows = false;                 // 手動による行削除を禁止する
            this.gcMultiRow3.Rows.Clear();                                  // 行数をクリア
            this.gcMultiRow3.RowCount = 5;                                  // 行数を設定
            this.gcMultiRow3.HideSelection = true;                          // GcMultiRow コントロールがフォーカスを失ったとき、セルの選択状態を非表示にする


            //// 部署残業理由コード一覧
            //gcMultiRow4.EditMode = EditMode.EditProgrammatically;

            //this.gcMultiRow4.AllowUserToAddRows = false;                    // 手動による行追加を禁止する
            //this.gcMultiRow4.AllowUserToDeleteRows = false;                 // 手動による行削除を禁止する
            //this.gcMultiRow4.Rows.Clear();                                  // 行数をクリア
            //this.gcMultiRow4.RowCount = 11;                                 // 行数を設定
        }

        private void gcMultiRow2_CellValueChanged(object sender, CellEventArgs e)
        {
            if (!gl.ChangeValueStatus) return;

            if (e.RowIndex < 0) return;

            // 過去データ表示のときは終了
            if (dID != string.Empty) return;

            // 社員番号のとき社員名を表示します
            if (e.CellName == "txtShainNum")
            {
                // ChangeValueイベントを発生させない
                gl.ChangeValueStatus = false;

                // 氏名を初期化
                gcMultiRow2[e.RowIndex, "lblName"].Value = string.Empty;

                // 奉行データベースより社員名を取得して表示します
                if (Utility.NulltoStr(gcMultiRow2[e.RowIndex, "txtShainNum"].Value) != string.Empty)
                {
                    string bCode = gcMultiRow2[e.RowIndex, "txtShainNum"].Value.ToString().PadLeft(10, '0');

                    gcMultiRow2[e.RowIndex, "lblName"].Value = getShainName(bCode);

                    // 2018/03/18 コメント化
                    //// ChangeValueイベントステータスをtrueに戻す
                    //gl.ChangeValueStatus = true;
                }

                // 2018/03/18
                // ChangeValueイベントステータスをtrueに戻す
                gl.ChangeValueStatus = true;
            }

            // 取消チェックのとき
            if (e.CellName == "chkTorikeshi")
            {
                if (gcMultiRow2[e.RowIndex, "chkTorikeshi"].Value.ToString() == "True")
                {
                    gcMultiRow2.Rows[e.RowIndex].BackColor = SystemColors.Control;
                    gcMultiRow2[e.RowIndex, "txtShainNum"].ReadOnly = true;
                    gcMultiRow2[e.RowIndex, "txtLineNum"].ReadOnly = true;
                    gcMultiRow2[e.RowIndex, "txtBmn"].ReadOnly = true;
                    gcMultiRow2[e.RowIndex, "txtHin"].ReadOnly = true;
                    gcMultiRow2[e.RowIndex, "txtOh"].ReadOnly = true;
                    gcMultiRow2[e.RowIndex, "txtOm"].ReadOnly = true;

                    gcMultiRow2[e.RowIndex, "lblName"].Style.ForeColor = Color.LightGray;
                    gcMultiRow2[e.RowIndex, "txtShainNum"].Style.ForeColor = Color.LightGray;
                    gcMultiRow2[e.RowIndex, "txtLineNum"].Style.ForeColor = Color.LightGray;
                    gcMultiRow2[e.RowIndex, "txtBmn"].Style.ForeColor = Color.LightGray;
                    gcMultiRow2[e.RowIndex, "txtHin"].Style.ForeColor = Color.LightGray;
                    gcMultiRow2[e.RowIndex, "txtOh"].Style.ForeColor = Color.LightGray;
                    gcMultiRow2[e.RowIndex, "txtOm"].Style.ForeColor = Color.LightGray;
                    gcMultiRow2[e.RowIndex, "labelCell7"].Style.ForeColor = Color.LightGray;
                    gcMultiRow2[e.RowIndex, "labelCell8"].Style.ForeColor = Color.LightGray;
                }
                else
                {
                    gcMultiRow2.Rows[e.RowIndex].BackColor = Color.Empty;
                    gcMultiRow2[e.RowIndex, "txtShainNum"].ReadOnly = false;
                    gcMultiRow2[e.RowIndex, "txtLineNum"].ReadOnly = false;
                    gcMultiRow2[e.RowIndex, "txtBmn"].ReadOnly = false;
                    gcMultiRow2[e.RowIndex, "txtHin"].ReadOnly = false;
                    gcMultiRow2[e.RowIndex, "txtOh"].ReadOnly = false;
                    gcMultiRow2[e.RowIndex, "txtOm"].ReadOnly = false;

                    gcMultiRow2[e.RowIndex, "lblName"].Style.ForeColor = Color.Blue;
                    gcMultiRow2[e.RowIndex, "txtShainNum"].Style.ForeColor = Color.Empty;
                    gcMultiRow2[e.RowIndex, "txtLineNum"].Style.ForeColor = Color.Empty;
                    gcMultiRow2[e.RowIndex, "txtBmn"].Style.ForeColor = Color.Empty;
                    gcMultiRow2[e.RowIndex, "txtHin"].Style.ForeColor = Color.Empty;
                    gcMultiRow2[e.RowIndex, "txtOh"].Style.ForeColor = Color.Empty;
                    gcMultiRow2[e.RowIndex, "txtOm"].Style.ForeColor = Color.Empty;
                    gcMultiRow2[e.RowIndex, "labelCell7"].Style.ForeColor = Color.Empty;
                    gcMultiRow2[e.RowIndex, "labelCell8"].Style.ForeColor = Color.Empty;
                }
            }
        }

        ///--------------------------------------------------------------------
        /// <summary>
        ///     奉行より社員名を取得する </summary>
        /// <param name="bCode">
        ///     検索用社員番号文字列 </param>
        /// <returns>
        ///     社員名</returns>
        ///--------------------------------------------------------------------
        private string getShainName(string bCode)
        {
            string sName = string.Empty;

            // 接続文字列取得
            string sc = sqlControl.obcConnectSting.get(_dbName);
            sqlControl.DataControl sdCon = new Common.sqlControl.DataControl(sc);

            SqlDataReader dR = sdCon.free_dsReader(Utility.getEmployee(bCode));

            while (dR.Read())
            {
                // 社員名取得
                sName = dR["Name"].ToString().Trim();
                break;
            }

            dR.Close();
            sdCon.Close();

            return sName;
        }


        private void gcMultiRow3_CellValueChanged(object sender, CellEventArgs e)
        {
            if (!gl.ChangeValueStatus) return;

            if (e.RowIndex < 0) return;

            // 過去データ表示のときは終了
            if (dID != string.Empty) return;

            // 社員番号のとき社員名を表示します
            if (e.CellName == "txtShainNum")
            {
                // ChangeValueイベントを発生させない
                gl.ChangeValueStatus = false;

                // 氏名を初期化
                gcMultiRow3[e.RowIndex, "lblName"].Value = string.Empty;
                
                // 奉行データベースより社員名を取得して表示します
                if (Utility.NulltoStr(gcMultiRow3[e.RowIndex, "txtShainNum"].Value) != string.Empty)
                {
                    string bCode = gcMultiRow3[e.RowIndex, "txtShainNum"].Value.ToString().PadLeft(10, '0');

                    gcMultiRow3[e.RowIndex, "lblName"].Value = getShainName(bCode);

                    // 2018/03/18 コメント化
                    //// ChangeValueイベントステータスをtrueに戻す
                    //gl.ChangeValueStatus = true;
                }

                // 2018/03/18
                // ChangeValueイベントステータスをtrueに戻す
                gl.ChangeValueStatus = true;
            }

            // 取消チェックのとき
            if (e.CellName == "chkTorikeshi")
            {
                if (gcMultiRow3[e.RowIndex, "chkTorikeshi"].Value.ToString() == "True")
                {
                    gcMultiRow3.Rows[e.RowIndex].BackColor = SystemColors.Control;
                    gcMultiRow3[e.RowIndex, "txtShainNum"].ReadOnly = true;
                    gcMultiRow3[e.RowIndex, "txtLineNum"].ReadOnly = true;
                    gcMultiRow3[e.RowIndex, "txtBmn"].ReadOnly = true;
                    gcMultiRow3[e.RowIndex, "txtHin"].ReadOnly = true;
                    gcMultiRow3[e.RowIndex, "txtZanRe1"].ReadOnly = true;
                    gcMultiRow3[e.RowIndex, "txtZanH1"].ReadOnly = true;
                    gcMultiRow3[e.RowIndex, "txtZanM1"].ReadOnly = true;
                    gcMultiRow3[e.RowIndex, "txtZanRe2"].ReadOnly = true;
                    gcMultiRow3[e.RowIndex, "txtZanH2"].ReadOnly = true;
                    gcMultiRow3[e.RowIndex, "txtZanM2"].ReadOnly = true;

                    gcMultiRow3[e.RowIndex, "lblName"].Style.ForeColor = Color.LightGray;
                    gcMultiRow3[e.RowIndex, "txtShainNum"].Style.ForeColor = Color.LightGray;
                    gcMultiRow3[e.RowIndex, "txtLineNum"].Style.ForeColor = Color.LightGray;
                    gcMultiRow3[e.RowIndex, "txtBmn"].Style.ForeColor = Color.LightGray;
                    gcMultiRow3[e.RowIndex, "txtHin"].Style.ForeColor = Color.LightGray;
                    gcMultiRow3[e.RowIndex, "txtZanRe1"].Style.ForeColor = Color.LightGray;
                    gcMultiRow3[e.RowIndex, "txtZanH1"].Style.ForeColor = Color.LightGray;
                    gcMultiRow3[e.RowIndex, "txtZanM1"].Style.ForeColor = Color.LightGray;
                    gcMultiRow3[e.RowIndex, "txtZanRe2"].Style.ForeColor = Color.LightGray;
                    gcMultiRow3[e.RowIndex, "txtZanH2"].Style.ForeColor = Color.LightGray;
                    gcMultiRow3[e.RowIndex, "txtZanM2"].Style.ForeColor = Color.LightGray;
                    gcMultiRow3[e.RowIndex, "labelCell10"].Style.ForeColor = Color.LightGray;
                    gcMultiRow3[e.RowIndex, "labelCell11"].Style.ForeColor = Color.LightGray;
                    gcMultiRow3[e.RowIndex, "labelCell16"].Style.ForeColor = Color.LightGray;
                    gcMultiRow3[e.RowIndex, "labelCell17"].Style.ForeColor = Color.LightGray;
                }
                else
                {
                    gcMultiRow3.Rows[e.RowIndex].BackColor = Color.Empty;
                    gcMultiRow3[e.RowIndex, "txtShainNum"].ReadOnly = false;
                    gcMultiRow3[e.RowIndex, "txtLineNum"].ReadOnly = false;
                    gcMultiRow3[e.RowIndex, "txtBmn"].ReadOnly = false;
                    gcMultiRow3[e.RowIndex, "txtHin"].ReadOnly = false;
                    gcMultiRow3[e.RowIndex, "txtZanRe1"].ReadOnly = false;
                    gcMultiRow3[e.RowIndex, "txtZanH1"].ReadOnly = false;
                    gcMultiRow3[e.RowIndex, "txtZanM1"].ReadOnly = false;
                    gcMultiRow3[e.RowIndex, "txtZanRe2"].ReadOnly = false;
                    gcMultiRow3[e.RowIndex, "txtZanH2"].ReadOnly = false;
                    gcMultiRow3[e.RowIndex, "txtZanM2"].ReadOnly = false;

                    gcMultiRow3[e.RowIndex, "lblName"].Style.ForeColor = Color.Blue;
                    gcMultiRow3[e.RowIndex, "txtShainNum"].Style.ForeColor = Color.Empty;
                    gcMultiRow3[e.RowIndex, "txtLineNum"].Style.ForeColor = Color.Empty;
                    gcMultiRow3[e.RowIndex, "txtBmn"].Style.ForeColor = Color.Empty;
                    gcMultiRow3[e.RowIndex, "txtHin"].Style.ForeColor = Color.Empty;
                    gcMultiRow3[e.RowIndex, "txtZanRe1"].Style.ForeColor = Color.Empty;
                    gcMultiRow3[e.RowIndex, "txtZanH1"].Style.ForeColor = Color.Empty;
                    gcMultiRow3[e.RowIndex, "txtZanM1"].Style.ForeColor = Color.Empty;
                    gcMultiRow3[e.RowIndex, "txtZanRe2"].Style.ForeColor = Color.Empty;
                    gcMultiRow3[e.RowIndex, "txtZanH2"].Style.ForeColor = Color.Empty;
                    gcMultiRow3[e.RowIndex, "txtZanM2"].Style.ForeColor = Color.Empty;
                    gcMultiRow3[e.RowIndex, "labelCell10"].Style.ForeColor = Color.Empty;
                    gcMultiRow3[e.RowIndex, "labelCell11"].Style.ForeColor = Color.Empty;
                    gcMultiRow3[e.RowIndex, "labelCell16"].Style.ForeColor = Color.Empty;
                    gcMultiRow3[e.RowIndex, "labelCell17"].Style.ForeColor = Color.Empty;
                }
            }
        }

        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     伝票画像表示 </summary>
        /// <param name="iX">
        ///     現在の伝票</param>
        /// <param name="tempImgName">
        ///     画像名</param>
        /// ------------------------------------------------------------------------------
        public void ShowImage(string tempImgName)
        {
            //修正画面へ組み入れた画像フォームの表示    
            //画像の出力が無い場合は、画像表示をしない。
            if (tempImgName == string.Empty)
            {
                leadImg.Visible = false;
                lblNoImage.Visible = false;
                //global.pblImagePath = string.Empty;
                return;
            }

            //画像ファイルがあるとき表示
            if (File.Exists(tempImgName))
            {
                lblNoImage.Visible = false;
                leadImg.Visible = true;

                // 画像操作ボタン
                btnPlus.Enabled = true;
                btnMinus.Enabled = true;

                //画像ロード
                Leadtools.Codecs.RasterCodecs.Startup();
                Leadtools.Codecs.RasterCodecs cs = new Leadtools.Codecs.RasterCodecs();

                // 描画時に使用される速度、品質、およびスタイルを制御します。 
                Leadtools.RasterPaintProperties prop = new Leadtools.RasterPaintProperties();
                prop = Leadtools.RasterPaintProperties.Default;
                prop.PaintDisplayMode = Leadtools.RasterPaintDisplayModeFlags.Resample;
                leadImg.PaintProperties = prop;

                leadImg.Image = cs.Load(tempImgName, 0, Leadtools.Codecs.CodecsLoadByteOrder.BgrOrGray, 1, 1);

                //画像表示倍率設定
                if (gl.miMdlZoomRate == 0f)
                {
                    leadImg.ScaleFactor *= (gl.ZOOM_RATE + gl.ZOOM_STEP);
                }
                else
                {
                    leadImg.ScaleFactor *= gl.miMdlZoomRate;
                }

                //画像のマウスによる移動を可能とする
                leadImg.InteractiveMode = Leadtools.WinForms.RasterViewerInteractiveMode.Pan;

                // グレースケールに変換
                Leadtools.ImageProcessing.GrayscaleCommand grayScaleCommand = new Leadtools.ImageProcessing.GrayscaleCommand();
                grayScaleCommand.BitsPerPixel = 8;
                grayScaleCommand.Run(leadImg.Image);
                leadImg.Refresh();

                cs.Dispose();
                Leadtools.Codecs.RasterCodecs.Shutdown();
                //global.pblImagePath = tempImgName;
            }
            else
            {
                //画像ファイルがないとき
                lblNoImage.Visible = true;

                // 画像操作ボタン
                btnPlus.Enabled = false;
                btnMinus.Enabled = false;

                leadImg.Visible = false;
                //global.pblImagePath = string.Empty;
            }
        }

        private void leadImg_MouseLeave(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Default;
        }

        private void leadImg_MouseMove(object sender, MouseEventArgs e)
        {
            this.Cursor = Cursors.Hand;
        }

        private void btnPlus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor < gl.ZOOM_MAX)
            {
                leadImg.ScaleFactor += gl.ZOOM_STEP;
            }

            gl.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }

        private void btnMinus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor > gl.ZOOM_MIN)
            {
                leadImg.ScaleFactor -= gl.ZOOM_STEP;
            }

            gl.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }

        private void gcMultiRow1_CellEnter(object sender, CellEventArgs e)
        {
            if (gcMultiRow1.EditMode == EditMode.EditProgrammatically)
            {
                gcMultiRow1.BeginEdit(true);
            }
        }

        private void gcMultiRow1_EditingControlShowing(object sender, EditingControlShowingEventArgs e)
        {
            if (e.Control is TextBoxEditingControl)
            {
                //イベントハンドラが複数回追加されてしまうので最初に削除する
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress2);

                // 数字のみ入力可能とする
                if (gcMultiRow1.CurrentCell.Name == "txtYear" || gcMultiRow1.CurrentCell.Name == "txtMonth" ||
                    gcMultiRow1.CurrentCell.Name == "txtDay")
                {
                    //イベントハンドラを追加する
                    e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
                }
            }
        }

        void Control_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b' && e.KeyChar != '\t')
                e.Handled = true;
        }

        void Control_KeyPress2(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= '0' && e.KeyChar <= '9') || (e.KeyChar >= 'a' && e.KeyChar <= 'z') ||
                e.KeyChar == '\b' || e.KeyChar == '\t')
                e.Handled = false;
            else e.Handled = true;
        }

        void Control_KeyPress3(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar != '0' && e.KeyChar != '5') && e.KeyChar != '\b' && e.KeyChar != '\t')
                e.Handled = true;
        }

        private void gcMultiRow1_CellValueChanged(object sender, CellEventArgs e)
        {
            if (!gl.ChangeValueStatus) return;

            if (e.RowIndex < 0) return;

            // 過去データ表示のときは終了
            if (dID != string.Empty) return;

            // 部署コードのとき部署名を表示します
            if (e.CellName == "txtBushoCode")
            {
                // ChangeValueイベントを発生させない
                gl.ChangeValueStatus = false;

                // 部署名を初期化
                gcMultiRow1[e.RowIndex, "lblShozoku"].Value = string.Empty;

                // 奉行データベースより部署名を取得して表示します
                if (Utility.NulltoStr(gcMultiRow1[e.RowIndex, "txtBushoCode"].Value) != string.Empty)
                {
                    string dName = string.Empty;
                    if (getDepartMentName(out dName, gcMultiRow1[e.RowIndex, "txtBushoCode"].Value.ToString(), e.RowIndex))
                    {
                        gcMultiRow1[e.RowIndex, "lblShozoku"].Value = dName;
                    }

                    // ChangeValueイベントステータスをtrueに戻す
                    gl.ChangeValueStatus = true;
                }
            }
            
            // 年月日
            if (e.CellName == "txtYear" || e.CellName == "txtMonth" || e.CellName == "txtDay")
            {
                // 曜日
                DateTime eDate;
                int tYY = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[e.RowIndex, "txtYear"].Value));
                string sDate = tYY.ToString() + "/" + Utility.NulltoStr(gcMultiRow1[e.RowIndex, "txtMonth"].Value) + "/" +
                        Utility.NulltoStr(gcMultiRow1[e.RowIndex, "txtDay"].Value);

                // 存在する日付と認識された場合、曜日を表示する
                if (DateTime.TryParse(sDate, out eDate))
                {
                    gcMultiRow1[e.RowIndex, "lblWeek"].Value = ("日月火水木金土").Substring(int.Parse(eDate.DayOfWeek.ToString("d")), 1);
                }
                else
                {
                    gcMultiRow1[e.RowIndex, "lblWeek"].Value = string.Empty;
                }
            }
        }

        ///-------------------------------------------------------------------------
        /// <summary>
        ///     奉行シリーズ部署名取得 </summary>
        /// <param name="dName">
        ///     取得する部署名</param>
        /// <param name="dCode">
        ///     部署コード</param>
        /// <param name="r">
        ///     MultiRowRowIndex</param>
        /// <returns>
        ///     true:該当あり, false:該当なし</returns>
        ///-------------------------------------------------------------------------
        private bool getDepartMentName(out string dName, string dCode, int r)
        {
            bool rtn = false;
            int c = 0;

            // 部署名を初期化
            dName = string.Empty;

            // 奉行データベースより部署名を取得して表示します
            if (Utility.NulltoStr(gcMultiRow1[r, "txtBushoCode"].Value) != string.Empty)
            {
                string b = string.Empty;

                // 検索用部署コード
                if (Utility.StrtoInt(gcMultiRow1[r, "txtBushoCode"].Value.ToString()) != global.flgOff)
                {
                    b = gcMultiRow1[r, "txtBushoCode"].Value.ToString().Trim().PadLeft(15, '0');
                }
                else
                {
                    b = gcMultiRow1[r, "txtBushoCode"].Value.ToString().Trim().PadRight(15, ' ');
                }

                // 接続文字列取得
                string sc = sqlControl.obcConnectSting.get(_dbName);
                sqlControl.DataControl sdCon = new Common.sqlControl.DataControl(sc);

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
                    c++;
                }

                dR.Close();
                sdCon.Close();

                if (c > 0)
                {
                    rtn = true;
                }
            }

            return rtn;
        }

        private void gcMultiRow2_CellEnter(object sender, CellEventArgs e)
        {
            if (gcMultiRow2.EditMode == EditMode.EditProgrammatically)
            {
                gcMultiRow2.BeginEdit(true);
            }
        }

        private void gcMultiRow3_CellEnter(object sender, CellEventArgs e)
        {
            if (gcMultiRow3.EditMode == EditMode.EditProgrammatically)
            {
                gcMultiRow3.BeginEdit(true);
            }
        }

        private void gcMultiRow2_EditingControlShowing(object sender, EditingControlShowingEventArgs e)
        {
            if (e.Control is TextBoxEditingControl)
            {
                //イベントハンドラが複数回追加されてしまうので最初に削除する
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress2);
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress3);

                // 数字のみ入力可能とする
                if (gcMultiRow2.CurrentCell.Name == "txtShainNum" ||
                    gcMultiRow2.CurrentCell.Name == "txtBmn" || gcMultiRow2.CurrentCell.Name == "txtOh")
                {
                    //イベントハンドラを追加する
                    e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
                }

                // 残業時は「０」「５」のみ入力可能とする
                if (gcMultiRow2.CurrentCell.Name == "txtOm")
                {
                    //イベントハンドラを追加する
                    e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress3);
                }
            }
        }

        private void gcMultiRow3_EditingControlShowing(object sender, EditingControlShowingEventArgs e)
        {
            if (e.Control is TextBoxEditingControl)
            {
                //イベントハンドラが複数回追加されてしまうので最初に削除する
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress2);
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress3);

                // 数字のみ入力可能とする
                if (gcMultiRow3.CurrentCell.Name == "txtShainNum" || gcMultiRow3.CurrentCell.Name == "txtBmn" || 
                    gcMultiRow3.CurrentCell.Name == "txtZanRe1" || gcMultiRow3.CurrentCell.Name == "txtZanH1" ||  
                    gcMultiRow3.CurrentCell.Name == "txtZanRe2" || gcMultiRow3.CurrentCell.Name == "txtZanH2") 
                {
                    //イベントハンドラを追加する
                    e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
                }

                // 残業時は「０」「５」のみ入力可能とする
                if (gcMultiRow3.CurrentCell.Name == "txtZanM1" || gcMultiRow3.CurrentCell.Name == "txtZanM2")
                {
                    //イベントハンドラを追加する
                    e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress3);
                }
            }
        }

        ///-----------------------------------------------------------------------------------
        /// <summary>
        ///     カレントデータを更新する</summary>
        /// <param name="iX">
        ///     カレントレコードのインデックス</param>
        ///-----------------------------------------------------------------------------------
        private void CurDataUpDate(string iX)
        {
            // エラーメッセージ
            string errMsg = "応援移動票データ更新";

            try
            {
                // 応援移動票ヘッダ行を取得
                DataSet1.応援移動票ヘッダRow r = dts.応援移動票ヘッダ.Single(a => a.ID == iX);

                // 応援移動票ヘッダテーブルセット更新
                r.年 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[0, "txtYear"].Value));
                r.月 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[0, "txtMonth"].Value));
                r.日 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[0, "txtDay"].Value));
                r.部署コード = Utility.NulltoStr(gcMultiRow1[0, "txtBushoCode"].Value);
                r.データ領域名 = _dbName;
                r.更新年月日 = DateTime.Now;

                if (checkBox1.Checked)
                {
                    r.確認 = global.flgOn;
                }
                else
                {
                    r.確認 = global.flgOff;
                }

                // 応援移動票明細・日中応援データ更新
                for (int i = 0; i < gcMultiRow2.RowCount; i++)
                {
                    int sID = int.Parse(gcMultiRow2[i, "txtID"].Value.ToString());
                    DataSet1.応援移動票明細Row m = (DataSet1.応援移動票明細Row)dts.応援移動票明細.FindByID(sID);

                    // 社員番号：先頭ゼロは除去
                    string sN = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow2[i, "txtShainNum"].Value)).ToString();
                    if (sN != global.FLGOFF)
                    {
                        m.社員番号 = sN;
                    }
                    else
                    {
                        m.社員番号 = string.Empty;
                    }

                    m.ライン = Utility.NulltoStr(gcMultiRow2[i, "txtLineNum"].Value);
                    m.部門 = Utility.NulltoStr(gcMultiRow2[i, "txtBmn"].Value);
                    m.製品群 = Utility.NulltoStr(gcMultiRow2[i, "txtHin"].Value);
                    m.応援時 = Utility.NulltoStr(gcMultiRow2[i, "txtOh"].Value);
                    m.応援分 = Utility.NulltoStr(gcMultiRow2[i, "txtOm"].Value);
                    m.残業理由1 = string.Empty;
                    m.残業時1 = string.Empty;
                    m.残業分1 = string.Empty;
                    m.残業理由2 = string.Empty;
                    m.残業時2 = string.Empty;
                    m.残業分2 = string.Empty;

                    if (gcMultiRow2[i, "chkTorikeshi"].Value.ToString() == "True")
                    {
                        m.取消 = global.FLGON;
                    }
                    else
                    {
                        m.取消 = global.FLGOFF;
                    }

                    m.更新年月日 = DateTime.Now;
                    m.社員名 = Utility.NulltoStr(gcMultiRow2[i, "lblName"].Value);
                }
                
                // 応援移動票明細・残業応援データ更新
                for (int i = 0; i < gcMultiRow3.RowCount; i++)
                {
                    int sID = int.Parse(gcMultiRow3[i, "txtID"].Value.ToString());
                    DataSet1.応援移動票明細Row m = (DataSet1.応援移動票明細Row)dts.応援移動票明細.FindByID(sID);

                    // 社員番号：先頭ゼロは除去
                    string sN = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[i, "txtShainNum"].Value)).ToString();
                    if (sN != global.FLGOFF)
                    {
                        m.社員番号 = sN;
                    }
                    else
                    {
                        m.社員番号 = string.Empty;
                    }

                    m.ライン = Utility.NulltoStr(gcMultiRow3[i, "txtLineNum"].Value);
                    m.部門 = Utility.NulltoStr(gcMultiRow3[i, "txtBmn"].Value);
                    m.製品群 = Utility.NulltoStr(gcMultiRow3[i, "txtHin"].Value);
                    m.応援時 = string.Empty;
                    m.応援分 = string.Empty;

                    // 残業理由１：先頭ゼロは除去
                    sN = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[i, "txtZanRe1"].Value)).ToString();
                    if (sN != global.FLGOFF)
                    {
                        // 残業理由記入あり
                        m.残業理由1 = sN;
                        m.残業時1 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[i, "txtZanH1"].Value)).ToString();
                        m.残業分1 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[i, "txtZanM1"].Value)).ToString();
                    }
                    else
                    {
                        // 残業理由記入なし
                        m.残業理由1 = string.Empty;
                        m.残業時1 = string.Empty;
                        m.残業分1 = string.Empty;
                    }

                    // 残業理由２：先頭ゼロは除去
                    sN = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[i, "txtZanRe2"].Value)).ToString();
                    if (sN != global.FLGOFF)
                    {
                        // 残業理由記入あり
                        m.残業理由2 = sN;
                        m.残業時2 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[i, "txtZanH2"].Value)).ToString();
                        m.残業分2 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow3[i, "txtZanM2"].Value)).ToString();
                    }
                    else
                    {
                        // 残業理由記入なし
                        m.残業理由2 = string.Empty;
                        m.残業時2 = string.Empty;
                        m.残業分2 = string.Empty;
                    }                    

                    if (gcMultiRow3[i, "chkTorikeshi"].Value.ToString() == "True")
                    {
                        m.取消 = global.FLGON;
                    }
                    else
                    {
                        m.取消 = global.FLGOFF;
                    }

                    m.更新年月日 = DateTime.Now;
                    m.社員名 = Utility.NulltoStr(gcMultiRow3[i, "lblName"].Value);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, errMsg, MessageBoxButtons.OK);
            }
            finally
            {
            }
        }

        private void btnFirst_Click(object sender, EventArgs e)
        {
            gcMultiRow1.CurrentCell = null;
            gcMultiRow2.CurrentCell = null;
            gcMultiRow3.CurrentCell = null;

            //カレントデータの更新
            CurDataUpDate(cID[cI]);

            //レコードの移動
            cI = 0;
            showOcrData(cI);
        }

        private void btnBefore_Click(object sender, EventArgs e)
        {
            gcMultiRow1.CurrentCell = null;
            gcMultiRow2.CurrentCell = null;
            gcMultiRow3.CurrentCell = null;

            //カレントデータの更新
            CurDataUpDate(cID[cI]);

            //レコードの移動
            if (cI > 0)
            {
                cI--;
                showOcrData(cI);
            }
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            gcMultiRow1.CurrentCell = null;
            gcMultiRow2.CurrentCell = null;
            gcMultiRow3.CurrentCell = null;

            //カレントデータの更新
            CurDataUpDate(cID[cI]);

            //レコードの移動
            if (cI + 1 < dts.応援移動票ヘッダ.Rows.Count)
            {
                cI++;
                showOcrData(cI);
            }
        }

        private void btnEnd_Click(object sender, EventArgs e)
        {
            gcMultiRow1.CurrentCell = null;
            gcMultiRow2.CurrentCell = null;
            gcMultiRow3.CurrentCell = null;

            //カレントデータの更新
            CurDataUpDate(cID[cI]);

            //レコードの移動
            cI = dts.応援移動票ヘッダ.Rows.Count - 1;
            showOcrData(cI);
        }

        private void hScrollBar1_Scroll(object sender, ScrollEventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(cID[cI]);

            //レコードの移動
            cI = hScrollBar1.Value;
            showOcrData(cI);
        }

        private void frmOuenCorrect_Shown(object sender, EventArgs e)
        {
            if (dID != string.Empty)
            {
                lnkRtn.Focus();
            }
        }

        private void frmOuenCorrect_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.Tag != null)
            {
                //「受入データ作成終了」「勤務票データなし」以外での終了のとき
                if (this.Tag.ToString() != END_MAKEDATA && this.Tag.ToString() != END_NODATA)
                {
                    //if (MessageBox.Show("終了します。よろしいですか", "終了確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    //{
                    //    e.Cancel = true;
                    //    return;
                    //}

                    // カレントデータ更新
                    if (dID == string.Empty) CurDataUpDate(cID[cI]);
                }

                // データベース更新
                adpMn.UpdateAll(dts);
            }

            // 解放する
            this.Dispose();
        }
        
        private void btnErrCheck_Click(object sender, EventArgs e)
        {
        }
        
        /// -----------------------------------------------------------------------------------
        /// <summary>
        ///     エラーチェックを実行する</summary>
        /// <param name="cIdx">
        ///     現在表示中の応援移動票ヘッダインデックス</param>
        /// <param name="ocr">
        ///     OCRDATAクラスインスタンス</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        /// -----------------------------------------------------------------------------------
        private bool getErrData(int cIdx, OCRData ocr)
        {
            // カレントレコード更新
            CurDataUpDate(cID[cIdx]);

            // エラー番号初期化
            ocr._errNumber = ocr.eNothing;

            // エラーメッセージクリーン
            ocr._errMsg = string.Empty;

            // 勤怠データＩ／Ｐ票データ読み込み
            getDataSet();

            // エラーチェック実行①:カレントレコードから最終レコードまで
            if (!ocr.errCheckMainOuen(cIdx, (dts.応援移動票ヘッダ.Rows.Count - 1), this, dts, hArray, cID))
            {
                return false;
            }

            // エラーチェック実行②:最初のレコードからカレントレコードの前のレコードまで
            if (cIdx > 0)
            {
                if (!ocr.errCheckMainOuen(0, (cIdx - 1), this, dts, hArray, cID))
                {
                    return false;
                }
            }

            // エラーなし
            lblErrMsg.Text = string.Empty;

            return true;
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (dts.勤務票ヘッダ.Count == 0)
            {
                MessageBox.Show("勤怠データＩ／Ｐ票データがありません", "勤怠データＩ／Ｐ票データ登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            
            // 勤怠データＩ／Ｐ票
            frmCorrect frmC = new frmCorrect(_dbName, _comName, string.Empty, false);
            frmC.ShowDialog();
            
            // 勤怠データＩ／Ｐ票データ読み込み
            getDataSet();
        }

        private void gcMultiRow2_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            string colName = gcMultiRow2.CurrentCell.Name;

            if (colName == "chkTorikeshi")
            {
                if (gcMultiRow2.IsCurrentCellDirty)
                {
                    gcMultiRow2.CommitEdit(DataErrorContexts.Commit);
                    gcMultiRow2.Refresh();
                }
            }
        }

        private void gcMultiRow3_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            string colName = gcMultiRow3.CurrentCell.Name;

            if (colName == "chkTorikeshi")
            {
                if (gcMultiRow3.IsCurrentCellDirty)
                {
                    gcMultiRow3.CommitEdit(DataErrorContexts.Commit);
                    gcMultiRow3.Refresh();
                }
            }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
        }

        ///-------------------------------------------------------------------------------
        /// <summary>
        ///     １．指定した応援移動票ヘッダデータと応援移動票明細データを削除する　
        ///     ２．該当する画像データを削除する</summary>
        /// <param name="i">
        ///     応援移動票ヘッダRow インデックス</param>
        ///-------------------------------------------------------------------------------
        private void DataDelete(int i)
        {
            string sImgNm = string.Empty;
            string errMsg = string.Empty;

            // 応援移動票データ削除
            try
            {
                // ヘッダIDを取得します
                DataSet1.応援移動票ヘッダRow r = dts.応援移動票ヘッダ.Single(a => a.ID == cID[i]);

                // 画像ファイル名を取得します
                sImgNm = r.画像名;

                // データテーブルからヘッダIDが一致する応援移動票明細データを削除します。
                errMsg = "応援移動票明細データ";
                foreach (DataSet1.応援移動票明細Row item in dts.応援移動票明細.Rows)
                {
                    if (item.RowState != DataRowState.Deleted && item.ヘッダID == r.ID)
                    {
                        item.Delete();
                    }
                }

                // データテーブルから勤務票ヘッダデータを削除します
                errMsg = "応援移動票ヘッダデータ";
                r.Delete();

                // データベース更新
                adpMn.UpdateAll(dts);

                // 画像ファイルを削除します
                errMsg = "応援移動票画像";
                if (sImgNm != string.Empty)
                {
                    if (System.IO.File.Exists(Properties.Settings.Default.dataPathOuen + sImgNm))
                    {
                        System.IO.File.Delete(Properties.Settings.Default.dataPathOuen + sImgNm);
                    }
                }

                // キー配列再構築
                keyArrayCreate();
            }
            catch (Exception ex)
            {
                MessageBox.Show(errMsg + "の削除に失敗しました" + Environment.NewLine + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            frmOuenIndex frm = new frmOuenIndex(_dbName, dts, hAdp, iAdp);
            frm.ShowDialog();
            string hID = frm.hdID;
            frm.Dispose();

            if (hID != string.Empty)
            {
                //カレントデータの更新
                CurDataUpDate(cID[cI]);

                // レコード検索
                for (int i = 0; i < cID.Length; i++)
                {
                    if (cID[i] == hID)
                    {
                        cI = i;
                        showOcrData(cI);
                        break;
                    }
                }
            }
        }

        private void lnkErrCheck_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // OCRDataクラス生成
            OCRData ocr = new OCRData(_dbName, bs);

            // エラーチェックを実行
            if (getErrData(cI, ocr))
            {
                MessageBox.Show("エラーはありませんでした", "エラーチェック", MessageBoxButtons.OK, MessageBoxIcon.Information);
                gcMultiRow1.CurrentCell = null;
                gcMultiRow2.CurrentCell = null;
                gcMultiRow3.CurrentCell = null;

                // データ表示
                showOcrData(cI);
            }
            else
            {
                // カレントインデックスをエラーありインデックスで更新
                cI = ocr._errHeaderIndex;

                // データ表示
                showOcrData(cI);

                // エラー表示
                ErrShow(ocr);
            }
        }

        private void lnkDel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("表示中の応援移動票データを削除します。よろしいですか", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            // レコードと画像ファイルを削除する
            DataDelete(cI);

            // 勤務票ヘッダテーブル件数カウント
            if (dts.応援移動票ヘッダ.Count() > 0)
            {
                // カレントレコードインデックスを再設定
                if (dts.応援移動票ヘッダ.Count() - 1 < cI) cI = dts.応援移動票ヘッダ.Count() - 1;

                // データ画面表示
                showOcrData(cI);
            }
            else
            {
                // ゼロならばプログラム終了
                MessageBox.Show("全ての応援移動票データが削除されました。処理を終了します。", "応援移動票削除", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                //終了処理
                this.Tag = END_NODATA;
                this.Close();
            }
        }

        private void lnkRtn_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // フォームを閉じる
            this.Tag = END_BUTTON;
            this.Close();
        }

        private void gcMultiRow2_CellLeave(object sender, CellEventArgs e)
        {
            if (gcMultiRow2.EditMode == EditMode.EditProgrammatically)
            {
                gcMultiRow2.EndEdit();
            }
        }

        private void gcMultiRow3_CellLeave(object sender, CellEventArgs e)
        {
            if (gcMultiRow3.EditMode == EditMode.EditProgrammatically)
            {
                gcMultiRow3.EndEdit();
            }
        }
    }
}
