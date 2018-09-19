using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Data.SqlClient;
using SZDS_TIMECARD.Common;
using SZDS_TIMECARD.OCR;
using GrapeCity.Win.MultiRow;
using Excel = Microsoft.Office.Interop.Excel;

namespace SZDS_TIMECARD.OCR
{
    public partial class frmCorrectPast : Form
    {
        /// ------------------------------------------------------------
        /// <summary>
        ///     コンストラクタ </summary>
        /// <param name="dbName">
        ///     会社領域データベース名</param>
        /// <param name="comName">
        ///     会社名</param>
        /// <param name="sID">
        ///     処理モード</param>
        /// <param name="eMode">
        ///     true:通常処理, false:応援移動票画面からの呼出し
        ///     （エラーチェック、汎用データ作成機能なし）</param>
        /// ------------------------------------------------------------
        public frmCorrectPast(string dbName, string comName, string sID, bool eMode)
        {
            InitializeComponent();

            _dbName = dbName;       // データベース名
            _comName = comName;     // 会社名
            dID = sID;              // 処理モード
            _eMode = eMode;         // 処理モード2

            /* テーブルアダプターマネージャーに過去勤務票ヘッダ、過去明細テーブルアダプターを割り付ける */
            pAdpMn.過去勤務票ヘッダTableAdapter = phAdp;
            pAdpMn.過去勤務票明細TableAdapter = piAdp;

            /* テーブルアダプターマネージャーに過去応援移動票ヘッダ、過去応援移動票明細テーブルアダプターを割り付ける */ 
            ouenAdp.過去応援移動票ヘッダTableAdapter = pohAdp;
            ouenAdp.過去応援移動票明細TableAdapter = pomAdp;

            // 休日テーブル読み込み
            kAdp.Fill(dts.休日);

            // 帰宅後勤務読み込み
            ktkAdp.Fill(dts.帰宅後勤務);
        }

        // データアダプターオブジェクト
        DataSet1TableAdapters.TableAdapterManager pAdpMn = new DataSet1TableAdapters.TableAdapterManager();
        DataSet1TableAdapters.過去勤務票ヘッダTableAdapter phAdp = new DataSet1TableAdapters.過去勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.過去勤務票明細TableAdapter piAdp = new DataSet1TableAdapters.過去勤務票明細TableAdapter();
        
        DataSet1TableAdapters.休日TableAdapter kAdp = new DataSet1TableAdapters.休日TableAdapter();

        DataSet1TableAdapters.TableAdapterManager ouenAdp = new DataSet1TableAdapters.TableAdapterManager();
        DataSet1TableAdapters.過去応援移動票ヘッダTableAdapter pohAdp = new DataSet1TableAdapters.過去応援移動票ヘッダTableAdapter();
        DataSet1TableAdapters.過去応援移動票明細TableAdapter pomAdp = new DataSet1TableAdapters.過去応援移動票明細TableAdapter();

        DataSet1TableAdapters.帰宅後勤務TableAdapter ktkAdp = new DataSet1TableAdapters.帰宅後勤務TableAdapter();

        // データセットオブジェクト
        DataSet1 dts = new DataSet1();

        // セル値
        private string cellName = string.Empty;         // セル名
        private string cellBeforeValue = string.Empty;  // 編集前
        private string cellAfterValue = string.Empty;   // 編集後

        #region 編集ログ・項目名 2015/09/08
        private const string LOG_YEAR = "年";
        private const string LOG_MONTH = "月";
        private const string LOG_DAY = "日";
        private const string LOG_TAIKEICD = "体系コード";
        private const string CELL_TORIKESHI = "取消";
        private const string CELL_NUMBER = "社員番号";
        private const string CELL_KIGOU = "記号";
        private const string CELL_FUTSU = "普通残業・時";
        private const string CELL_FUTSU_M = "普通残業・分";
        private const string CELL_SHINYA = "深夜残業・時";
        private const string CELL_SHINYA_M = "深夜残業・分";
        private const string CELL_SHIGYO = "始業時刻・時";
        private const string CELL_SHIGYO_M = "始業時刻・分";
        private const string CELL_SHUUGYO = "終業時刻・時";
        private const string CELL_SHUUGYO_M = "終業時刻・分";
        #endregion 編集ログ・項目名

        // カレント社員情報
        //SCCSDataSet.社員所属Row cSR = null;
        
        // 社員マスターより取得した所属コード
        string mSzCode = string.Empty;

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

        bool _eMode = true;

        // dataGridView1_CellEnterステータス
        bool gridViewCellEnterStatus = true;

        // 編集ログ書き込み状態
        bool editLogStatus = false;

        // 部署別勤務体系配列クラス
        xlsData bs;

        // ライン・部門・製品群コード配列取得 
        string[] hArray = null;

        // カレントデータRowsインデックス
        string [] cID = null;
        int cI = 0;

        // グローバルクラス
        global gl = new global();

        // プリントイメージ
        string _img = string.Empty;

        private void frmCorrect_Load(object sender, EventArgs e)
        {
            this.pictureBox1.Image = new Bitmap(pictureBox1.Width, pictureBox1.Height);

            // フォーム最大値
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小値
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // Tabキーの既定のショートカットキーを解除する。
            gcMultiRow1.ShortcutKeyManager.Unregister(Keys.Tab);
            gcMultiRow2.ShortcutKeyManager.Unregister(Keys.Tab);
            gcMultiRow1.ShortcutKeyManager.Unregister(Keys.Enter);
            gcMultiRow2.ShortcutKeyManager.Unregister(Keys.Enter);

            // Tabキーのショートカットキーにユーザー定義のショートカットキーを割り当てる。
            gcMultiRow1.ShortcutKeyManager.Register(new clsKeyTab.CustomMoveToNextContorol(), Keys.Tab);
            gcMultiRow2.ShortcutKeyManager.Register(new clsKeyTab.CustomMoveToNextContorol(), Keys.Tab);
            gcMultiRow1.ShortcutKeyManager.Register(new clsKeyTab.CustomMoveToNextContorol(), Keys.Enter);
            gcMultiRow2.ShortcutKeyManager.Register(new clsKeyTab.CustomMoveToNextContorol(), Keys.Enter);
            
            // データセットへデータを読み込みます
            getDataSet();   // 出勤簿
            getOuenDataSet();   // 応援移動票

            // 部署別残業理由シートデータ配列取得
            bs = new xlsData();
            bs.zArray = bs.getShiftCode(string.Empty);

            // 部署別残業理由シートデータ配列取得
            bs.rArray = bs.getZanReason();

            // ライン・部門・製品群コード配列取得 
            hArray = getCategoryArray();

            // キャプション
            this.Text = "過去勤怠データＩ／Ｐ票表示";

            // GCMultiRow初期化
            gcMrSetting();

            // データ表示
            showOcrData(dID);

            // tagを初期化
            this.Tag = string.Empty;

            // 現在の表示倍率を初期化
            gl.miMdlZoomRate = 0f;
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     キー配列作成 </summary>
        ///-------------------------------------------------------------
        private void keyArrayCreate()
        {
            int iX = 0;
            foreach (var t in dts.過去勤務票ヘッダ.OrderBy(a => a.ID))
            {
                Array.Resize(ref cID, iX + 1);
                cID[iX] = t.ID;
                iX++;
            }
        }

        #region データグリッドビューカラム定義
        private static string cCheck = "col1";      // 取消
        private static string cShainNum = "col2";   // 社員番号
        private static string cName = "col3";       // 氏名
        private static string cKinmu = "col4";      // 勤務記号
        private static string cZH = "col5";         // 残業時
        private static string cZE = "col6";         // :
        private static string cZM = "col7";         // 残業分
        private static string cSIH = "col8";        // 深夜時
        private static string cSIE = "col9";        // :
        private static string cSIM = "col10";       // 深夜分
        private static string cSH = "col11";        // 開始時
        private static string cSE = "col12";        // :
        private static string cSM = "col13";        // 開始分
        private static string cEH = "col14";        // 終了時
        private static string cEE = "col15";        // :
        private static string cEM = "col16";        // 終了分
        //private static string cID = "colID";        // ID
        private static string cSzCode = "colSzCode";  // 所属コード
        private static string cSzName = "colSzName";  // 所属名

        #endregion

        private void gcMrSetting()
        {
            //multirow編集モード
            gcMultiRow2.EditMode = EditMode.EditProgrammatically;

            this.gcMultiRow2.AllowUserToAddRows = false;                    // 手動による行追加を禁止する
            this.gcMultiRow2.AllowUserToDeleteRows = false;                 // 手動による行削除を禁止する
            this.gcMultiRow2.Rows.Clear();                                  // 行数をクリア
            this.gcMultiRow2.RowCount = 1;                                  // 行数を設定
            this.gcMultiRow2.HideSelection = true;                          // GcMultiRow コントロールがフォーカスを失ったとき、セルの選択状態を非表示にする

            //multirow編集モード
            gcMultiRow1.EditMode = EditMode.EditProgrammatically;

            this.gcMultiRow1.AllowUserToAddRows = false;                    // 手動による行追加を禁止する
            this.gcMultiRow1.AllowUserToDeleteRows = false;                 // 手動による行削除を禁止する
            this.gcMultiRow1.Rows.Clear();                                  // 行数をクリア
            this.gcMultiRow1.RowCount = global.MAX_GYO;                     // 行数を設定
            this.gcMultiRow1.HideSelection = true;                          // GcMultiRow コントロールがフォーカスを失ったとき、セルの選択状態を非表示にする
        }

        ///----------------------------------------------------------------------------
        /// <summary>
        ///     CSVデータをMDBへインサートする</summary>
        ///----------------------------------------------------------------------------
        private void GetCsvDataToMDB()
        {
            // CSVファイル数をカウント
            string[] inCsv = System.IO.Directory.GetFiles(Properties.Settings.Default.dataPathIP, "*.csv");

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
            ocr.CsvToMdb(Properties.Settings.Default.dataPathIP, frmP, _dbName);

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;
        }

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            //if (e.Control is DataGridViewTextBoxEditingControl)
            //{
            //    // 数字のみ入力可能とする
            //    if (dGV.CurrentCell.ColumnIndex != 0 && dGV.CurrentCell.ColumnIndex != 2)
            //    {
            //        //イベントハンドラが複数回追加されてしまうので最初に削除する
            //        e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
            //        e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress2);

            //        //イベントハンドラを追加する
            //        e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
            //    }
            //}
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
            if (e.KeyChar != '0' && e.KeyChar != '5' && e.KeyChar != '\b' && e.KeyChar != '\t')
                e.Handled = true;
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void frmCorrect_Shown(object sender, EventArgs e)
        {
            if (dID != string.Empty) lnkRtn.Focus();
        }

        private void dataGridView3_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is DataGridViewTextBoxEditingControl)
            {
                //イベントハンドラが複数回追加されてしまうので最初に削除する
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                //イベントハンドラを追加する
                e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
            }
        }

        private void dataGridView4_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is DataGridViewTextBoxEditingControl)
            {
                //イベントハンドラが複数回追加されてしまうので最初に削除する
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                //イベントハンドラを追加する
                e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
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
            string errMsg = "出勤簿テーブル更新";

            try
            {
                // 過去勤務票ヘッダテーブル行を取得
                DataSet1.過去勤務票ヘッダRow r = dts.過去勤務票ヘッダ.Single(a => a.ID == iX);

                // 過去勤務票ヘッダテーブルセット更新
                r.年 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow2[0, "txtYear"].Value));
                r.月 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow2[0, "txtMonth"].Value));
                r.日 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow2[0, "txtDay"].Value));
                r.部署コード = Utility.NulltoStr(gcMultiRow2[0, "txtBushoCode"].Value);
                r.シフトコード = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow2[0, "txtSftCode"].Value));
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

                // 過去勤務票明細テーブルセット更新
                for (int i = 0; i < gcMultiRow1.RowCount; i++)
                {
                    int sID = int.Parse(gcMultiRow1[i, "txtID"].Value.ToString());
                    DataSet1.過去勤務票明細Row m = (DataSet1.過去勤務票明細Row)dts.過去勤務票明細.FindByID(sID);

                    if (gcMultiRow1[i, "chkOuen"].Value.ToString() == "True")
                    {
                        m.応援 = global.FLGON;
                    }
                    else
                    {
                        m.応援 = global.FLGOFF;
                    }

                    if (gcMultiRow1[i, "chkSft"].Value.ToString() == "True")
                    {
                        m.シフト通り = global.FLGON;
                    }
                    else
                    {
                        m.シフト通り = global.FLGOFF;
                    }

                    // 社員番号：先頭ゼロは除去
                    string sN = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[i, "txtShainNum"].Value)).ToString();
                    if (sN != global.FLGOFF)
                    {
                        m.社員番号 = sN;
                    }
                    else
                    {
                        m.社員番号 = string.Empty;
                    }

                    m.事由1 = Utility.NulltoStr(gcMultiRow1[i, "txtJiyu1"].Value);
                    m.事由2 = Utility.NulltoStr(gcMultiRow1[i, "txtJiyu2"].Value);
                    m.事由3 = Utility.NulltoStr(gcMultiRow1[i, "txtJiyu3"].Value);
                    m.シフトコード = Utility.NulltoStr(gcMultiRow1[i, "txtSftCode"].Value);
                    m.出勤時 = Utility.NulltoStr(gcMultiRow1[i, "txtSh"].Value);
                    m.出勤分 = Utility.NulltoStr(gcMultiRow1[i, "txtSm"].Value);
                    m.退勤時 = Utility.NulltoStr(gcMultiRow1[i, "txtEh"].Value);
                    m.退勤分 = Utility.NulltoStr(gcMultiRow1[i, "txtEm"].Value);

                    // 残業理由１：先頭ゼロは除去
                    sN = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[i, "txtZanRe1"].Value)).ToString();
                    if (sN != global.FLGOFF)
                    {
                        // 残業理由記入あり
                        m.残業理由1 = sN;
                        m.残業時1 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[i, "txtZanH1"].Value)).ToString();
                        m.残業分1 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[i, "txtZanM1"].Value)).ToString();
                    }
                    else
                    {
                        // 残業理由記入なし
                        m.残業理由1 = string.Empty;
                        m.残業時1 = string.Empty;
                        m.残業分1 = string.Empty;
                    }
                    
                    // 残業理由２：先頭ゼロは除去
                    sN = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[i, "txtZanRe2"].Value)).ToString();
                    if (sN != global.FLGOFF)
                    {
                        // 残業理由記入あり
                        m.残業理由2 = sN;
                        m.残業時2 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[i, "txtZanH2"].Value)).ToString();
                        m.残業分2 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[i, "txtZanM2"].Value)).ToString();
                    }
                    else
                    {
                        // 残業理由記入なし
                        m.残業理由2 = string.Empty;
                        m.残業時2 = string.Empty;
                        m.残業分2 = string.Empty;
                    }

                    if (gcMultiRow1[i, "chkTorikeshi"].Value.ToString() == "True")
                    {
                        m.取消 = global.FLGON;
                    }
                    else
                    {
                        m.取消 = global.FLGOFF;
                    }

                    m.社員名 = Utility.NulltoStr(gcMultiRow1[i, "lblName"].Value);
                    m.更新年月日 = DateTime.Now;
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

        /// ----------------------------------------------------------------------------------------------------
        /// <summary>
        ///     空白以外のとき、指定された文字数になるまで左側に０を埋めこみ、右寄せした文字列を返す
        /// </summary>
        /// <param name="tm">
        ///     文字列</param>
        /// <param name="len">
        ///     文字列の長さ</param>
        /// <returns>
        ///     文字列</returns>
        /// ----------------------------------------------------------------------------------------------------
        private string timeVal(object tm, int len)
        {
            string t = Utility.NulltoStr(tm);
            if (t != string.Empty) return t.PadLeft(len, '0');
            else return t;
        }

        /// ----------------------------------------------------------------------------------------------------
        /// <summary>
        ///     空白以外のとき、先頭文字が０のとき先頭文字を削除した文字列を返す　
        ///     先頭文字が０以外のときはそのまま返す
        /// </summary>
        /// <param name="tm">
        ///     文字列</param>
        /// <returns>
        ///     文字列</returns>
        /// ----------------------------------------------------------------------------------------------------
        private string timeValH(object tm)
        {
            string t = Utility.NulltoStr(tm);

            if (t != string.Empty)
            {
                t = t.PadLeft(2, '0');
                if (t.Substring(0, 1) == "0")
                {
                    t = t.Substring(1, 1);
                }
            }

            return t;
        }

        /// ------------------------------------------------------------------------------------
        /// <summary>
        ///     Bool値を数値に変換する </summary>
        /// <param name="b">
        ///     True or False</param>
        /// <returns>
        ///     true:1, false:0</returns>
        /// ------------------------------------------------------------------------------------
        private int booltoFlg(string b)
        {
            if (b == "True") return global.flgOn;
            else return global.flgOff;
        }
        
        ///-----------------------------------------------------------------
        /// <summary>
        ///     エラーチェックボタン </summary>
        /// <param name="sender">
        ///     </param>
        /// <param name="e">
        ///     </param>
        ///-----------------------------------------------------------------
        private void btnErrCheck_Click(object sender, EventArgs e)
        {
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
        }

        private void frmCorrect_FormClosing(object sender, FormClosingEventArgs e)
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
                CurDataUpDate(dID);
            }

            // データベース更新
            pAdpMn.UpdateAll(dts);

            // 解放する
            this.Dispose();
        }

        private void btnDataMake_Click(object sender, EventArgs e)
        {
        }

        /// -----------------------------------------------------------------------
        /// <summary>
        ///     就業奉行・受入CSVデータ出力 </summary>
        /// -----------------------------------------------------------------------
        private void textDataMake()
        {
            if (MessageBox.Show("就業奉行受け渡しデータを作成します。よろしいですか", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;

            // OCRDataクラス生成
            OCRData ocr = new OCRData(_dbName, bs);

            // エラーチェックを実行する
            if (getErrData(dID, ocr)) // エラーがなかったとき
            {
                // OCROutputクラス インスタンス生成
                OCROutputPast kd = new OCROutputPast(this, dts);

                // 汎用データ作成
                kd.SaveData(_dbName, dID);          

                //// 画像ファイル退避（出勤簿・応援移動票）
                //tifFileMove();

                //// 過去出勤簿データ作成
                //saveLastData();

                //// 過去応援移動票データ作成
                //saveLastOuenData();

                //// 設定月数分経過した過去画像と過去の出勤簿・応援移動票データを削除する
                //deleteArchived();

                //// 勤務票データ削除
                //deleteDataAll();

                //// 応援移動票データ削除
                //deleteOuenDataAll();

                // MDBファイル最適化
                mdbCompact();

                //終了
                MessageBox.Show("終了しました。就業奉行VERPで勤務データ受け入れを行ってください。", "就業奉行受け入れデータ作成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Tag = END_MAKEDATA;
                this.Close();
            }
            else
            {
                // カレントインデックスをエラーありインデックスで更新
                cI = ocr._errHeaderIndex;

                // エラーあり
                showOcrData(dID);    // データ表示
                ErrShow(ocr);   // エラー表示
            }
        }

        /// -----------------------------------------------------------------------------------
        /// <summary>
        ///     エラーチェックを実行する</summary>
        /// <param name="cIdx">
        ///     現在表示中の勤務票ヘッダデータインデックス</param>
        /// <param name="ocr">
        ///     OCRDATAクラスインスタンス</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        /// -----------------------------------------------------------------------------------
        private bool getErrData(string dID, OCRData ocr)
        {
            // カレントレコード更新
            CurDataUpDate(dID);

            // エラー番号初期化
            ocr._errNumber = ocr.eNothing;

            // エラーメッセージクリーン
            ocr._errMsg = string.Empty;

            // 過去応援移動票データ作成読み込み
            getOuenDataSet();

            // エラーチェック実行
            if (!ocr.errCheckMain(dID, dts))
            {
                return false;
            }

            //　応援移動票側から勤怠データＩ／Ｐ票データのチェックは撤廃 2017/04/07
            //// 応援移動票データに対応する勤怠データＩ／Ｐ票データが存在するか
            //string eMsg = string.Empty;
            //if (!chkIpOuenMatch(dts, out eMsg))
            //{
            //    ocr._errNumber = ocr.eIpOuen;
            //    ocr._errMsg = eMsg;
            //    ocr._errRow = 0;
            //    return false;
            //}

            // エラーなし
            lblErrMsg.Text = string.Empty;

            return true;
        }

        ///-------------------------------------------------------------------------------------
        /// <summary>
        ///     応援移動票データに対応する勤怠データＩ／Ｐ票明細データがあるか </summary>
        /// <param name="dts">
        ///     DataSet1 </param>
        /// <returns>
        ///     true:あり、false:なし</returns>
        ///-------------------------------------------------------------------------------------
        private bool chkIpOuenMatch(DataSet1 dts, out string msg)
        {
            msg = string.Empty;
            bool rtn = true;
            int eCnt = 0;

            foreach (var m in dts.過去応援移動票明細.Where(a => a.社員番号 != string.Empty && a.取消 == global.FLGOFF))
            {
                if (!dts.過去勤務票明細.Any(a => a.過去勤務票ヘッダRow.年 == m.過去応援移動票ヘッダRow.年 &&
                                           a.過去勤務票ヘッダRow.月 == m.過去応援移動票ヘッダRow.月 &&
                                           a.過去勤務票ヘッダRow.日 == m.過去応援移動票ヘッダRow.日 &&
                                           a.社員番号.PadLeft(6, '0') == m.社員番号.PadLeft(6, '0')))
                {
                    eCnt++;
                }
            }

            if (eCnt > 0)
            {
                //msg = m.応援移動票ヘッダRow.年.ToString() + "/" + m.応援移動票ヘッダRow.月.ToString() + "/" + m.応援移動票ヘッダRow.日.ToString() + " の勤怠データＩ／Ｐ票が存在しません";
                msg = "勤怠データＩ／Ｐ票が存在しない応援移動票データが" + eCnt.ToString() + "件あります";
                rtn = false;
            }

            return rtn;
        }

        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     画像ファイル退避処理 勤務票・応援移動票</summary>
        ///----------------------------------------------------------------------------------
        private void tifFileMove()
        {
            // 移動先フォルダがあるか？なければ作成する（TIFフォルダ）
            if (!System.IO.Directory.Exists(Properties.Settings.Default.tifPath))
                System.IO.Directory.CreateDirectory(Properties.Settings.Default.tifPath);

            // 出勤簿ヘッダデータを取得する
            var s = dts.過去勤務票ヘッダ.OrderBy(a => a.ID);

            foreach (var t in s)
            {
                // 画像ファイルパスを取得する
                string fromImg = Properties.Settings.Default.dataPathIP + t.画像名;

                // 移動先ファイルパス
                string toImg = Properties.Settings.Default.tifPath + t.画像名;

                // 同名ファイルが既に登録済みのときは削除する
                if (System.IO.File.Exists(toImg)) System.IO.File.Delete(toImg);

                // ファイルを移動する
                if (System.IO.File.Exists(fromImg)) System.IO.File.Move(fromImg, toImg);
            }
            
            // 応援移動票データを取得する
            var e = dts.応援移動票ヘッダ.OrderBy(a => a.ID);

            foreach (var t in e)
            {
                // 画像ファイルパスを取得する
                string fromImg = Properties.Settings.Default.dataPathOuen + t.画像名;

                // 移動先ファイルパス
                string toImg = Properties.Settings.Default.tifPath + t.画像名;

                // 同名ファイルが既に登録済みのときは削除する
                if (System.IO.File.Exists(toImg)) System.IO.File.Delete(toImg);

                // ファイルを移動する
                if (System.IO.File.Exists(fromImg)) System.IO.File.Move(fromImg, toImg);
            }
        }

        /// ---------------------------------------------------------------------
        /// <summary>
        ///     MDBファイルを最適化する </summary>
        /// ---------------------------------------------------------------------
        private void mdbCompact()
        {
            try
            {
                JRO.JetEngine jro = new JRO.JetEngine();
                string OldDb = Properties.Settings.Default.mdbOlePath;
                string NewDb = Properties.Settings.Default.mdbPathTemp;

                jro.CompactDatabase(OldDb, NewDb);

                //今までのバックアップファイルを削除する
                System.IO.File.Delete(Properties.Settings.Default.mdbPath + global.MDBBACK);

                //今までのファイルをバックアップとする
                System.IO.File.Move(Properties.Settings.Default.mdbPath + global.MDBFILE, Properties.Settings.Default.mdbPath + global.MDBBACK);

                //一時ファイルをMDBファイルとする
                System.IO.File.Move(Properties.Settings.Default.mdbPath + global.MDBTEMP, Properties.Settings.Default.mdbPath + global.MDBFILE);
            }
            catch (Exception e)
            {
                MessageBox.Show("MDB最適化中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
            }
        }

        private void btnPlus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor < gl.ZOOM_MAX)
            {
                leadImg.ScaleFactor += gl.ZOOM_STEP;
            }
            gl.miMdlZoomRate = (float)leadImg.ScaleFactor;

            //if (dGV.RowCount == global.NIPPOU_TATE)
            //{
            //    global.miMdlZoomRate_TATE = (float)leadImg.ScaleFactor;
            //}
            //else if (dGV.RowCount == global.NIPPOU_YOKO)
            //{
            //    global.miMdlZoomRate_YOKO = (float)leadImg.ScaleFactor;
            //}
        }

        private void btnMinus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor > gl.ZOOM_MIN)
            {
                leadImg.ScaleFactor -= gl.ZOOM_STEP;
            }
            gl.miMdlZoomRate = (float)leadImg.ScaleFactor;

            //if (dGV.RowCount == global.NIPPOU_TATE)
            //{
            //    global.miMdlZoomRate_TATE = (float)leadImg.ScaleFactor;
            //}
            //else if (dGV.RowCount == global.NIPPOU_YOKO)
            //{
            //    global.miMdlZoomRate_YOKO = (float)leadImg.ScaleFactor;
            //}
        }

        /// ---------------------------------------------------------------------------------
        /// <summary>
        ///     設定月数分経過した過去画像と過去勤務データ、過去応援移動票データを削除する </summary> 
        /// ---------------------------------------------------------------------------------
        private void deleteArchived()
        {
            // 削除月設定が0のとき、「過去画像削除しない」とみなし終了する
            if (Properties.Settings.Default.dataDelSpan == global.flgOff) return;

            try
            {
                // 削除年月の取得
                DateTime dt = DateTime.Parse(DateTime.Today.Year.ToString() + "/" + DateTime.Today.Month.ToString() + "/01");
                DateTime delDate = dt.AddMonths(Properties.Settings.Default.dataDelSpan * (-1));
                int _dYY = delDate.Year;            //基準年
                int _dMM = delDate.Month;           //基準月
                int _dYYMM = _dYY * 100 + _dMM;     //基準年月
                int _waYYMM = (delDate.Year - Properties.Settings.Default.rekiHosei) * 100 + _dMM;   //基準年月(和暦）

                // 設定月数分経過した過去画像・過去勤務票データを削除する
                deleteLastDataArchived(_dYYMM);

                // 設定月数分経過した過去画像・過去応援移動票データを削除する
                deleteLastOuenDataArchived(_dYYMM);
            }
            catch (Exception e)
            {
                MessageBox.Show("過去画像・過去勤務票データ削除中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
                return;
            }
            finally
            {
                //if (ocr.sCom.Connection.State == ConnectionState.Open) ocr.sCom.Connection.Close();
            }
        }

        /// ---------------------------------------------------------------------------
        /// <summary>
        ///     過去勤務票データ削除～登録 </summary>
        /// ---------------------------------------------------------------------------
        private void saveLastData()
        {
            try
            {
                // データベース更新
                //adpMn.UpdateAll(dts);
                pAdpMn.UpdateAll(dts);

                //  過去勤務票ヘッダデータとその明細データを削除します
                //deleteLastData();
                delPastData();

                // データセットへデータを再読み込みします
                getDataSet();

                // 過去勤務票ヘッダデータと過去勤務票明細データを作成します
                addLastdata();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "過去勤務票データ作成エラー", MessageBoxButtons.OK);
            }
            finally
            {
            }
        }


        ///------------------------------------------------------
        /// <summary>
        ///     過去勤務票データ削除 </summary>
        ///------------------------------------------------------
        private void delPastData()
        {
            // 過去勤務票ヘッダデータ削除
            foreach (var t in dts.過去勤務票ヘッダ)
            {
                string sBusho = t.部署コード;
                int sYY = t.年;
                int sMM = t.月;
                int sDD = t.日;

                // 過去勤務票ヘッダ削除
                delPastHeader(sBusho, sYY, sMM, sDD);
            }

            // 過去勤務票明細データ削除
            delPastItem();
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     過去勤務票ヘッダデータ削除 </summary>
        /// <param name="bCode">
        ///     部署コード</param>
        /// <param name="syy">
        ///     対象年</param>
        /// <param name="smm">
        ///     対象月</param>
        /// <param name="sdd">
        ///     対象日</param>
        ///----------------------------------------------------------------
        private void delPastHeader(string bCode, int syy, int smm, int sdd)
        {
            OleDbCommand sCom = new OleDbCommand();
            mdbControl mdb = new mdbControl();
            mdb.dbConnect(sCom);

            try
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();
                sb.Append("delete from 過去勤務票ヘッダ ");
                sb.Append("where 部署コード = ? and 年 = ? and 月 = ? and 日 = ?");

                sCom.CommandText = sb.ToString();
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@b", bCode);
                sCom.Parameters.AddWithValue("@y", syy);
                sCom.Parameters.AddWithValue("@m", smm);
                sCom.Parameters.AddWithValue("@d", sdd);

                sCom.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw;
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open)
                {
                    sCom.Connection.Close();
                }
            }
        }

        ///--------------------------------------------------------
        /// <summary>
        ///     過去勤務票明細データ削除 </summary>
        ///--------------------------------------------------------
        private void delPastItem()
        {
            OleDbCommand sCom = new OleDbCommand();
            mdbControl mdb = new mdbControl();
            mdb.dbConnect(sCom);

            try
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();
                sb.Append("delete a.ヘッダID from  過去勤務票明細 as a ");
                sb.Append("where not EXISTS (select * from 過去勤務票ヘッダ ");
                sb.Append("WHERE 過去勤務票ヘッダ.ID = a.ヘッダID)");
                
                sCom.CommandText = sb.ToString();
                sCom.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw;
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open)
                {
                    sCom.Connection.Close();
                }
            }
        }

        ///------------------------------------------------------
        /// <summary>
        ///     過去応援移動票データ削除 </summary>
        ///------------------------------------------------------
        private void delPastOuenData()
        {
            // 過去応援移動票ヘッダデータ削除
            foreach (var t in dts.応援移動票ヘッダ)
            {
                string sBusho = t.部署コード;
                int sYY = t.年;
                int sMM = t.月;
                int sDD = t.日;

                // 過去応援移動票ヘッダ削除
                delPastOuenHeader(sBusho, sYY, sMM, sDD);
            }

            // 過去応援移動票明細データ削除
            delPastOuenItem();
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     過去応援移動票ヘッダデータ削除 </summary>
        /// <param name="bCode">
        ///     部署コード</param>
        /// <param name="syy">
        ///     対象年</param>
        /// <param name="smm">
        ///     対象月</param>
        /// <param name="sdd">
        ///     対象日</param>
        ///----------------------------------------------------------------
        private void delPastOuenHeader(string bCode, int syy, int smm, int sdd)
        {
            OleDbCommand sCom = new OleDbCommand();
            mdbControl mdb = new mdbControl();
            mdb.dbConnect(sCom);

            try
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();
                sb.Append("delete from 過去応援移動票ヘッダ ");
                sb.Append("where 部署コード = ? and 年 = ? and 月 = ? and 日 = ?");

                sCom.CommandText = sb.ToString();
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@b", bCode);
                sCom.Parameters.AddWithValue("@y", syy);
                sCom.Parameters.AddWithValue("@m", smm);
                sCom.Parameters.AddWithValue("@d", sdd);

                sCom.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw;
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open)
                {
                    sCom.Connection.Close();
                }
            }
        }

        ///--------------------------------------------------------
        /// <summary>
        ///     過去応援移動票明細データ削除 </summary>
        ///--------------------------------------------------------
        private void delPastOuenItem()
        {
            OleDbCommand sCom = new OleDbCommand();
            mdbControl mdb = new mdbControl();
            mdb.dbConnect(sCom);

            try
            {
                StringBuilder sb = new StringBuilder();

                sb.Clear();
                sb.Append("delete a.ヘッダID from  過去応援移動票明細 as a ");
                sb.Append("where not EXISTS (select * from 過去応援移動票ヘッダ ");
                sb.Append("WHERE 過去応援移動票ヘッダ.ID = a.ヘッダID)");

                sCom.CommandText = sb.ToString();
                sCom.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw;
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open)
                {
                    sCom.Connection.Close();
                }
            }
        }

        /// ---------------------------------------------------------------------------
        /// <summary>
        ///     過去応援移動票データ登録 </summary>
        /// ---------------------------------------------------------------------------
        private void saveLastOuenData()
        {
            try
            {
                // データベース更新
                ouenAdp.UpdateAll(dts);

                // 過去応援移動票ヘッダデータとその明細データを削除します
                //deleteLastOuenData();
                delPastOuenData();

                // データセットへデータを再読み込みします
                getOuenDataSet();

                // 過去勤務票ヘッダデータと過去勤務票明細データを作成します
                addLastOuendata();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "過去勤務票データ作成エラー", MessageBoxButtons.OK);
            }
            finally
            {
            }
        }

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     過去勤務票ヘッダデータとその明細データを削除します</summary>    
        ///     
        /// -------------------------------------------------------------------------
        private void deleteLastData()
        {
            OleDbCommand sCom = new OleDbCommand();
            OleDbCommand sCom2 = new OleDbCommand();
            OleDbCommand sCom3 = new OleDbCommand();

            mdbControl mdb = new mdbControl();
            mdb.dbConnect(sCom);
            mdb.dbConnect(sCom2);
            mdb.dbConnect(sCom3);

            OleDbDataReader dR = null;
            OleDbDataReader dR2 = null;

            StringBuilder sb = new StringBuilder();
            StringBuilder sbd = new StringBuilder();

            try
            {
                // 対象データ : 取消は対象外とする
                sb.Clear();
                sb.Append("Select 勤務票明細.ヘッダID, 勤務票明細.ID,");
                sb.Append("勤務票ヘッダ.年, 勤務票ヘッダ.月, 勤務票ヘッダ.日,");
                sb.Append("勤務票明細.社員番号 from 勤務票ヘッダ inner join 勤務票明細 ");
                sb.Append("on 勤務票ヘッダ.ID = 勤務票明細.ヘッダID ");
                sb.Append("where 勤務票明細.取消 = '").Append(global.FLGOFF).Append("'");
                sb.Append("order by 勤務票明細.ヘッダID, 勤務票明細.ID");

                sCom.CommandText = sb.ToString();
                dR = sCom.ExecuteReader();

                while (dR.Read())
                {
                    // ヘッダID
                    string hdID = string.Empty;

                    // 日付と社員番号で過去データを抽出（該当するのは1件）
                    sb.Clear();
                    sb.Append("Select 過去勤務票明細.ヘッダID,過去勤務票明細.ID,");
                    sb.Append("過去勤務票ヘッダ.年, 過去勤務票ヘッダ.月, 過去勤務票ヘッダ.日,");
                    sb.Append("過去勤務票明細.社員番号 from 過去勤務票ヘッダ inner join 過去勤務票明細 ");
                    sb.Append("on 過去勤務票ヘッダ.ID = 過去勤務票明細.ヘッダID ");
                    sb.Append("where ");
                    sb.Append("過去勤務票ヘッダ.年 = ? and ");
                    sb.Append("過去勤務票ヘッダ.月 = ? and ");
                    sb.Append("過去勤務票ヘッダ.日 = ? and ");
                    sb.Append("過去勤務票ヘッダ.データ領域名 = ? and ");
                    sb.Append("過去勤務票明細.社員番号 = ?");

                    sCom2.CommandText = sb.ToString();
                    sCom2.Parameters.Clear();
                    sCom2.Parameters.AddWithValue("@yy", dR["年"].ToString());
                    sCom2.Parameters.AddWithValue("@mm", dR["月"].ToString());
                    sCom2.Parameters.AddWithValue("@dd", dR["日"].ToString());
                    sCom2.Parameters.AddWithValue("@db", _dbName);
                    sCom2.Parameters.AddWithValue("@n", dR["社員番号"].ToString());

                    dR2 = sCom2.ExecuteReader();

                    while (dR2.Read())
                    {
                        //// ヘッダIDを取得
                        //if (hdID == string.Empty)
                        //{
                        //    hdID = dR2["ヘッダID"].ToString();
                        //}

                        // 過去勤務票明細レコード削除
                        sbd.Clear();
                        sbd.Append("delete from 過去勤務票明細 ");
                        sbd.Append("where ID = ?");

                        sCom3.CommandText = sbd.ToString();
                        sCom3.Parameters.Clear();
                        sCom3.Parameters.AddWithValue("@id", dR2["ID"].ToString());

                        sCom3.ExecuteNonQuery();
                    }

                    dR2.Close();
                }

                dR.Close();

                // データベース接続解除
                if (sCom.Connection.State == ConnectionState.Open)
                {
                    sCom.Connection.Close();
                }

                if (sCom2.Connection.State == ConnectionState.Open)
                {
                    sCom2.Connection.Close();
                }

                if (sCom3.Connection.State == ConnectionState.Open)
                {
                    sCom3.Connection.Close();
                }

                // データベース再接続
                mdb.dbConnect(sCom);
                mdb.dbConnect(sCom2);

                // 明細データのない過去勤務票ヘッダデータを抽出
                sb.Clear();
                sb.Append("Select 過去勤務票ヘッダ.ID,過去勤務票明細.ヘッダID ");
                sb.Append("from 過去勤務票ヘッダ left join 過去勤務票明細 ");
                sb.Append("on 過去勤務票ヘッダ.ID = 過去勤務票明細.ヘッダID ");
                sb.Append("where ");
                sb.Append("過去勤務票明細.ヘッダID is null");
                sCom.CommandText = sb.ToString();
                dR = sCom.ExecuteReader();

                while (dR.Read())
                {
                    // 過去勤務票ヘッダレコード削除
                    sbd.Clear();

                    sbd.Append("delete from 過去勤務票ヘッダ ");
                    sbd.Append("where ID = ?");

                    sCom2.CommandText = sbd.ToString();
                    sCom2.Parameters.Clear();
                    sCom2.Parameters.AddWithValue("@id", dR["ID"].ToString());

                    sCom2.ExecuteNonQuery();
                }

                dR.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open)
                {
                    sCom.Connection.Close();
                }

                if (sCom2.Connection.State == ConnectionState.Open)
                {
                    sCom2.Connection.Close();
                }

                if (sCom3.Connection.State == ConnectionState.Open)
                {
                    sCom3.Connection.Close();
                }
            }
        }


        /// -------------------------------------------------------------------------
        /// <summary>
        ///     過去応援移動票ヘッダデータとその明細データを削除します</summary>
        /// -------------------------------------------------------------------------
        private void deleteLastOuenData()
        {
            OleDbCommand sCom = new OleDbCommand();
            OleDbCommand sCom2 = new OleDbCommand();
            OleDbCommand sCom3 = new OleDbCommand();

            mdbControl mdb = new mdbControl();
            mdb.dbConnect(sCom);
            mdb.dbConnect(sCom2);
            mdb.dbConnect(sCom3);

            OleDbDataReader dR = null;
            OleDbDataReader dR2 = null;

            StringBuilder sb = new StringBuilder();
            StringBuilder sbd = new StringBuilder();

            try
            {
                // 対象データ : 取消は対象外とする
                sb.Clear();
                sb.Append("Select 応援移動票明細.ヘッダID, 応援移動票明細.ID,");
                sb.Append("応援移動票ヘッダ.年, 応援移動票ヘッダ.月, 応援移動票ヘッダ.日,");
                sb.Append("応援移動票明細.社員番号 from 応援移動票ヘッダ inner join 応援移動票明細 ");
                sb.Append("on 応援移動票ヘッダ.ID = 応援移動票明細.ヘッダID ");
                sb.Append("where 応援移動票明細.取消 = '").Append(global.FLGOFF).Append("'");
                sb.Append("order by 応援移動票明細.ヘッダID, 応援移動票明細.ID");

                sCom.CommandText = sb.ToString();
                dR = sCom.ExecuteReader();

                while (dR.Read())
                {
                    // ヘッダID
                    string hdID = string.Empty;

                    // 日付と社員番号で過去データを抽出（該当するのは1件）
                    sb.Clear();
                    sb.Append("Select 過去応援移動票明細.ヘッダID,過去応援移動票明細.ID,");
                    sb.Append("過去応援移動票ヘッダ.年, 過去応援移動票ヘッダ.月, 過去応援移動票ヘッダ.日,");
                    sb.Append("過去応援移動票明細.社員番号 from 過去応援移動票ヘッダ inner join 過去応援移動票明細 ");
                    sb.Append("on 過去応援移動票ヘッダ.ID = 過去応援移動票明細.ヘッダID ");
                    sb.Append("where ");
                    sb.Append("過去応援移動票ヘッダ.年 = ? and ");
                    sb.Append("過去応援移動票ヘッダ.月 = ? and ");
                    sb.Append("過去応援移動票ヘッダ.日 = ? and ");
                    sb.Append("過去応援移動票ヘッダ.データ領域名 = ? and ");
                    sb.Append("過去応援移動票明細.社員番号 = ?");

                    sCom2.CommandText = sb.ToString();
                    sCom2.Parameters.Clear();
                    sCom2.Parameters.AddWithValue("@yy", dR["年"].ToString());
                    sCom2.Parameters.AddWithValue("@mm", dR["月"].ToString());
                    sCom2.Parameters.AddWithValue("@dd", dR["日"].ToString());
                    sCom2.Parameters.AddWithValue("@db", _dbName);
                    sCom2.Parameters.AddWithValue("@n", dR["社員番号"].ToString());

                    dR2 = sCom2.ExecuteReader();

                    while (dR2.Read())
                    {
                        //// ヘッダIDを取得
                        //if (hdID == string.Empty)
                        //{
                        //    hdID = dR2["ヘッダID"].ToString();
                        //}

                        // 過去応援移動票明細レコード削除
                        sbd.Clear();
                        sbd.Append("delete from 過去応援移動票明細 ");
                        sbd.Append("where ID = ?");

                        sCom3.CommandText = sbd.ToString();
                        sCom3.Parameters.Clear();
                        sCom3.Parameters.AddWithValue("@id", dR2["ID"].ToString());

                        sCom3.ExecuteNonQuery();
                    }

                    dR2.Close();
                }

                dR.Close();

                // データベース接続解除
                if (sCom.Connection.State == ConnectionState.Open)
                {
                    sCom.Connection.Close();
                }

                if (sCom2.Connection.State == ConnectionState.Open)
                {
                    sCom2.Connection.Close();
                }

                if (sCom3.Connection.State == ConnectionState.Open)
                {
                    sCom3.Connection.Close();
                }

                // データベース再接続
                mdb.dbConnect(sCom);
                mdb.dbConnect(sCom2);

                // 明細データのない過去応援移動票ヘッダデータを抽出
                sb.Clear();
                sb.Append("Select 過去応援移動票ヘッダ.ID,過去応援移動票明細.ヘッダID ");
                sb.Append("from 過去応援移動票ヘッダ left join 過去応援移動票明細 ");
                sb.Append("on 過去応援移動票ヘッダ.ID = 過去応援移動票明細.ヘッダID ");
                sb.Append("where ");
                sb.Append("過去応援移動票明細.ヘッダID is null");
                sCom.CommandText = sb.ToString();
                dR = sCom.ExecuteReader();

                while (dR.Read())
                {
                    // 過去応援移動票ヘッダレコード削除
                    sbd.Clear();

                    sbd.Append("delete from 過去応援移動票ヘッダ ");
                    sbd.Append("where ID = ?");

                    sCom2.CommandText = sbd.ToString();
                    sCom2.Parameters.Clear();
                    sCom2.Parameters.AddWithValue("@id", dR["ID"].ToString());

                    sCom2.ExecuteNonQuery();
                }

                dR.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open)
                {
                    sCom.Connection.Close();
                }

                if (sCom2.Connection.State == ConnectionState.Open)
                {
                    sCom2.Connection.Close();
                }

                if (sCom3.Connection.State == ConnectionState.Open)
                {
                    sCom3.Connection.Close();
                }
            }
        }

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     過去勤務票ヘッダデータと過去勤務票明細データを作成します</summary>
        ///     
        /// -------------------------------------------------------------------------
        private void addLastdata()
        {
            for (int i = 0; i < dts.過去勤務票ヘッダ.Rows.Count; i++)
            {
                // -------------------------------------------------------------------------
                //      過去勤務票ヘッダレコードを作成します
                // -------------------------------------------------------------------------
                DataSet1.勤務票ヘッダRow hr = (DataSet1.勤務票ヘッダRow)dts.過去勤務票ヘッダ.Rows[i];
                DataSet1.過去勤務票ヘッダRow nr = dts.過去勤務票ヘッダ.New過去勤務票ヘッダRow();

                #region テーブルカラム名比較～データコピー

                // 勤務票ヘッダのカラムを順番に読む
                for (int j = 0; j < dts.過去勤務票ヘッダ.Columns.Count; j++)
                {
                    // 過去勤務票ヘッダのカラムを順番に読む
                    for (int k = 0; k < dts.過去勤務票ヘッダ.Columns.Count; k++)
                    {
                        // フィールド名が同じであること
                        if (dts.過去勤務票ヘッダ.Columns[j].ColumnName == dts.過去勤務票ヘッダ.Columns[k].ColumnName)
                        {
                            if (dts.過去勤務票ヘッダ.Columns[k].ColumnName == "更新年月日")
                            {
                                nr[k] = DateTime.Now;   // 更新年月日はこの時点のタイムスタンプを登録
                            }
                            else
                            {
                                nr[k] = hr[j];          // データをコピー
                            }
                            break;
                        }
                    }
                }
                #endregion

                // 過去勤務票ヘッダデータテーブルに追加
                dts.過去勤務票ヘッダ.Add過去勤務票ヘッダRow(nr);

                // -------------------------------------------------------------------------
                //      過去勤務票明細レコードを作成します
                // -------------------------------------------------------------------------
                var mm = dts.過去勤務票明細
                    .Where(a => a.RowState != DataRowState.Deleted && a.RowState != DataRowState.Detached &&
                           a.ヘッダID == hr.ID)
                    .OrderBy(a => a.ID);

                foreach (var item in mm)
                {
                    DataSet1.勤務票明細Row m = (DataSet1.勤務票明細Row)dts.過去勤務票明細.Rows.Find(item.ID);
                    DataSet1.過去勤務票明細Row nm = dts.過去勤務票明細.New過去勤務票明細Row();

                    // 取消は対象外：2015/10/01
                    if (m.取消 == global.FLGON) continue;

                    // 社員番号が空白のレコードは対象外とします
                    if (m.社員番号 == string.Empty) continue;

                    #region  テーブルカラム名比較～データコピー

                    // 勤務票明細のカラムを順番に読む
                    for (int j = 0; j < dts.過去勤務票明細.Columns.Count; j++)
                    {
                        // IDはオートナンバーのため値はコピーしない
                        if (dts.過去勤務票明細.Columns[j].ColumnName != "ID")
                        {
                            // 過去勤務票ヘッダのカラムを順番に読む
                            for (int k = 0; k < dts.過去勤務票明細.Columns.Count; k++)
                            {
                                // フィールド名が同じであること
                                if (dts.過去勤務票明細.Columns[j].ColumnName == dts.過去勤務票明細.Columns[k].ColumnName)
                                {
                                    if (dts.過去勤務票明細.Columns[k].ColumnName == "更新年月日")
                                    {
                                        nm[k] = DateTime.Now;   // 更新年月日はこの時点のタイムスタンプを登録
                                    }
                                    else
                                    {
                                        nm[k] = m[j];          // データをコピー
                                    }
                                    break;
                                }
                            }
                        }
                    }
                    #endregion

                    // 過去勤務票明細データテーブルに追加
                    dts.過去勤務票明細.Add過去勤務票明細Row(nm);
                }
            }

            // データベース更新
            pAdpMn.UpdateAll(dts);
        }

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     過去応援移動票ヘッダデータと過去応援移動票明細データを作成します</summary>
        ///     
        /// -------------------------------------------------------------------------
        private void addLastOuendata()
        {
            for (int i = 0; i < dts.応援移動票ヘッダ.Rows.Count; i++)
            {
                // -------------------------------------------------------------------------
                //      過去応援移動票ヘッダレコードを作成します
                // -------------------------------------------------------------------------
                DataSet1.応援移動票ヘッダRow hr = (DataSet1.応援移動票ヘッダRow)dts.応援移動票ヘッダ.Rows[i];
                DataSet1.過去応援移動票ヘッダRow nr = dts.過去応援移動票ヘッダ.New過去応援移動票ヘッダRow();

                #region テーブルカラム名比較～データコピー

                // 応援移動票ヘッダのカラムを順番に読む
                for (int j = 0; j < dts.応援移動票ヘッダ.Columns.Count; j++)
                {
                    // 過去応援移動票ヘッダのカラムを順番に読む
                    for (int k = 0; k < dts.過去応援移動票ヘッダ.Columns.Count; k++)
                    {
                        // フィールド名が同じであること
                        if (dts.応援移動票ヘッダ.Columns[j].ColumnName == dts.過去応援移動票ヘッダ.Columns[k].ColumnName)
                        {
                            if (dts.過去応援移動票ヘッダ.Columns[k].ColumnName == "更新年月日")
                            {
                                nr[k] = DateTime.Now;   // 更新年月日はこの時点のタイムスタンプを登録
                            }
                            else
                            {
                                nr[k] = hr[j];          // データをコピー
                            }

                            break;
                        }
                    }
                }
                #endregion

                // 過去応援移動票ヘッダデータテーブルに追加
                dts.過去応援移動票ヘッダ.Add過去応援移動票ヘッダRow(nr);

                // -------------------------------------------------------------------------
                //      過去応援移動票明細レコードを作成します
                // -------------------------------------------------------------------------
                var mm = dts.応援移動票明細
                    .Where(a => a.RowState != DataRowState.Deleted && a.RowState != DataRowState.Detached &&
                           a.ヘッダID == hr.ID)
                    .OrderBy(a => a.ID);

                foreach (var item in mm)
                {
                    DataSet1.応援移動票明細Row m = (DataSet1.応援移動票明細Row)dts.応援移動票明細.Rows.Find(item.ID);
                    DataSet1.過去応援移動票明細Row nm = dts.過去応援移動票明細.New過去応援移動票明細Row();

                    // 取消は対象外：2015/10/01
                    if (m.取消 == global.FLGON) continue;

                    // 社員番号が空白のレコードは対象外とします
                    if (m.社員番号 == string.Empty) continue;

                    #region  テーブルカラム名比較～データコピー

                    // 応援移動票明細のカラムを順番に読む
                    for (int j = 0; j < dts.応援移動票明細.Columns.Count; j++)
                    {
                        // IDはオートナンバーのため値はコピーしない
                        if (dts.応援移動票明細.Columns[j].ColumnName != "ID")
                        {
                            // 過去応援移動票ヘッダのカラムを順番に読む
                            for (int k = 0; k < dts.過去応援移動票明細.Columns.Count; k++)
                            {
                                // フィールド名が同じであること
                                if (dts.応援移動票明細.Columns[j].ColumnName == dts.過去応援移動票明細.Columns[k].ColumnName)
                                {
                                    if (dts.過去応援移動票明細.Columns[k].ColumnName == "更新年月日")
                                    {
                                        nm[k] = DateTime.Now;   // 更新年月日はこの時点のタイムスタンプを登録
                                    }
                                    else
                                    {
                                        nm[k] = m[j];          // データをコピー
                                    }
                                    break;
                                }
                            }
                        }
                    }
                    #endregion

                    // 過去応援移動票明細データテーブルに追加
                    dts.過去応援移動票明細.Add過去応援移動票明細Row(nm);
                }
            }

            // データベース更新
            ouenAdp.UpdateAll(dts);
        }

        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
        //    //if (e.RowIndex < 0) return;

        //    string colName = dGV.Columns[e.ColumnIndex].Name;

        //    if (colName == cSH || colName == cSE || colName == cEH || colName == cEE ||
        //        colName == cZH || colName == cZE || colName == cSIH || colName == cSIE)
        //    {
        //        e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;
        //    }
        }

        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            //string colName = dGV.Columns[dGV.CurrentCell.ColumnIndex].Name;
            ////if (colName == cKyuka || colName == cCheck)
            ////{
            ////    if (dGV.IsCurrentCellDirty)
            ////    {
            ////        dGV.CommitEdit(DataGridViewDataErrorContexts.Commit);
            ////        dGV.RefreshEdit();
            ////    }
            ////}
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void dataGridView1_CellEnter_1(object sender, DataGridViewCellEventArgs e)
        {
            //// 時が入力済みで分が未入力のとき分に"00"を表示します
            //if (dGV[ColH, dGV.CurrentRow.Index].Value != null)
            //{
            //    if (dGV[ColH, dGV.CurrentRow.Index].Value.ToString().Trim() != string.Empty)
            //    {
            //        if (dGV[ColM, dGV.CurrentRow.Index].Value == null)
            //        {
            //            dGV[ColM, dGV.CurrentRow.Index].Value = "00";
            //        }
            //        else if (dGV[ColM, dGV.CurrentRow.Index].Value.ToString().Trim() == string.Empty)
            //        {
            //            dGV[ColM, dGV.CurrentRow.Index].Value = "00";
            //        }
            //    }
            //}
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
                    leadImg.ScaleFactor *= gl.ZOOM_RATE;
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

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     基準年月以前の過去勤務票ヘッダデータとその明細データを削除します</summary>
        /// <param name="sYYMM">
        ///     基準年月</param>     
        /// -------------------------------------------------------------------------
        private void deleteLastDataArchived(int sYYMM)
        {
            // データ読み込み
            getDataSet();

            // 基準年月以前の過去勤務票ヘッダデータを取得します
            var h = dts.過去勤務票ヘッダ
                    .Where(a => a.RowState != DataRowState.Deleted && a.RowState != DataRowState.Detached &&
                                a.年 * 100 + a.月 < sYYMM);

            // foreach用の配列を作成
            var hLst = h.ToList();

            foreach (var lh in hLst)
            {
                // ヘッダIDが一致する過去勤務票明細を取得します
                var m = dts.過去勤務票明細
                    .Where(a => a.RowState != DataRowState.Deleted && a.RowState != DataRowState.Detached &&
                                a.ヘッダID == lh.ID);

                // foreach用の配列を作成
                var list = m.ToList();

                // 該当過去勤務票明細を削除します
                foreach (var lm in list)
                {
                    DataSet1.過去勤務票明細Row lRow = (DataSet1.過去勤務票明細Row)dts.過去勤務票明細.Rows.Find(lm.ID);
                    lRow.Delete();
                }

                // 画像ファイルを削除します
                string imgPath = Properties.Settings.Default.tifPath + lh.画像名;
                File.Delete(imgPath);

                // 過去勤務票ヘッダを削除します
                lh.Delete();
            }

            // データベース更新
            pAdpMn.UpdateAll(dts);
        }

        /// -------------------------------------------------------------------------
        /// <summary>
        ///     基準年月以前の過去応援移動票ヘッダデータとその明細データを削除します</summary>
        /// <param name="sYYMM">
        ///     基準年月</param>     
        /// -------------------------------------------------------------------------
        private void deleteLastOuenDataArchived(int sYYMM)
        {
            // データ読み込み
            getOuenDataSet();

            // 基準年月以前の過去勤務票ヘッダデータを取得します
            var h = dts.過去応援移動票ヘッダ
                    .Where(a => a.RowState != DataRowState.Deleted && a.RowState != DataRowState.Detached &&
                                a.年 * 100 + a.月 < sYYMM);

            // foreach用の配列を作成
            var hLst = h.ToList();

            foreach (var lh in hLst)
            {
                // ヘッダIDが一致する過去応援移動票明細を取得します
                var m = dts.過去応援移動票明細
                    .Where(a => a.RowState != DataRowState.Deleted && a.RowState != DataRowState.Detached &&
                                a.ヘッダID == lh.ID);

                // foreach用の配列を作成
                var list = m.ToList();

                // 該当過去応援移動票明細を削除します
                foreach (var lm in list)
                {
                    DataSet1.過去応援移動票明細Row lRow = (DataSet1.過去応援移動票明細Row)dts.過去応援移動票明細.Rows.Find(lm.ID);
                    lRow.Delete();
                }

                // 画像ファイルを削除します
                string imgPath = Properties.Settings.Default.tifPath + lh.画像名;
                File.Delete(imgPath);

                // 過去応援移動票ヘッダを削除します
                lh.Delete();
            }

            // データベース更新
            ouenAdp.UpdateAll(dts);
        }

        /// -----------------------------------------------------------------------------
        /// <summary>
        ///     設定月数分経過した過去画像を削除する</summary>
        /// <param name="_dYYMM">
        ///     基準年月 (例：201401)</param>
        /// -----------------------------------------------------------------------------
        private void deleteImageArchived(int _dYYMM)
        {
            int _DataYYMM;
            string fileYYMM;

            // 設定月数分経過した過去画像を削除する            
            foreach (string files in System.IO.Directory.GetFiles(Properties.Settings.Default.tifPath, "*.tif"))
            {
                // ファイル名が規定外のファイルは読み飛ばします
                if (System.IO.Path.GetFileName(files).Length < 21) continue;

                //ファイル名より年月を取得する
                fileYYMM = System.IO.Path.GetFileName(files).Substring(0, 6);

                if (Utility.NumericCheck(fileYYMM))
                {
                    _DataYYMM = int.Parse(fileYYMM);

                    //基準年月以前なら削除する
                    if (_DataYYMM <= _dYYMM) File.Delete(files);
                }
            }
        }
        
        /// -------------------------------------------------------------------
        /// <summary>
        ///     応援移動票ヘッダデータと応援移動票明細データを全件削除します</summary>
        /// -------------------------------------------------------------------
        private void deleteOuenDataAll()
        {
            // 応援移動票データ読み込み
            getOuenDataSet();

            // 応援移動票明細全行削除
            var m = dts.応援移動票明細.Where(a => a.RowState != DataRowState.Deleted);
            foreach (var t in m)
            {
                t.Delete();
            }

            // 応援移動票ヘッダ全行削除
            var h = dts.応援移動票ヘッダ.Where(a => a.RowState != DataRowState.Deleted);
            foreach (var t in h)
            {
                t.Delete();
            }

            // データベース更新
            ouenAdp.UpdateAll(dts);

            // 後片付け
            dts.応援移動票明細.Dispose();
            dts.応援移動票ヘッダ.Dispose();
        }

        private void maskedTextBox3_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void txtYear_TextChanged(object sender, EventArgs e)
        {
            //// 曜日
            //DateTime eDate;
            //int tYY = Utility.StrtoInt(txtYear.Text);
            //string sDate = tYY.ToString() + "/" + Utility.EmptytoZero(txtMonth.Text) + "/" +
            //        Utility.EmptytoZero(txtDay.Text);

            //// 存在する日付と認識された場合、曜日を表示する
            //if (DateTime.TryParse(sDate, out eDate))
            //{
            //    txtWeekDay.Text = ("日月火水木金土").Substring(int.Parse(eDate.DayOfWeek.ToString("d")), 1);
            //}
            //else
            //{
            //    txtWeekDay.Text = string.Empty;
            //}
        }

        private void dGV_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            //if (editLogStatus)
            //{
            //    if (e.ColumnIndex == 0 || e.ColumnIndex == 1 || e.ColumnIndex == 3 || e.ColumnIndex == 4 ||
            //        e.ColumnIndex == 6 || e.ColumnIndex == 7 || e.ColumnIndex == 9 || e.ColumnIndex == 10 ||
            //        e.ColumnIndex == 12 || e.ColumnIndex == 13 || e.ColumnIndex == 15)
            //    {
            //        dGV.CommitEdit(DataGridViewDataErrorContexts.Commit);
            //        cellAfterValue = Utility.NulltoStr(dGV[e.ColumnIndex, e.RowIndex].Value);

            //        //// 変更のとき編集ログデータを書き込み
            //        //if (cellBeforeValue != cellAfterValue)
            //        //{
            //        //    logDataUpdate(e.RowIndex, cI, global.flgOn);
            //        //}
            //    }
            //}
        }

        private void txtYear_Enter(object sender, EventArgs e)
        {
            //if (editLogStatus)
            //{
            //    if (sender == txtYear) cellName = LOG_YEAR;
            //    if (sender == txtMonth) cellName = LOG_MONTH;
            //    if (sender == txtDay) cellName = LOG_DAY;
            //    //if (sender == txtSftCode) cellName = LOG_TAIKEICD;

            //    TextBox tb = (TextBox)sender;

            //    // 値を保持
            //    cellBeforeValue = Utility.NulltoStr(tb.Text);
            //}
        }

        private void txtYear_Leave(object sender, EventArgs e)
        {
            if (editLogStatus)
            {
                TextBox tb = (TextBox)sender;
                cellAfterValue = Utility.NulltoStr(tb.Text);

                //// 変更のとき編集ログデータを書き込み
                //if (cellBeforeValue != cellAfterValue)
                //{
                //    logDataUpdate(0, cI, global.flgOff);
                //}
            }
        }

        private void gcMultiRow1_CellValueChanged(object sender, CellEventArgs e)
        {
            if (!gl.ChangeValueStatus) return;

            if (e.RowIndex < 0) return;

            //// 過去データ表示のときは終了
            //if (dID != string.Empty) return;

            // 応援チェックのとき社員名を赤表示します
            if (e.CellName == "chkOuen")
            {
                if (gcMultiRow1[e.RowIndex, "chkOuen"].Value.ToString() == "True")
                {
                    gcMultiRow1[e.RowIndex, "lblName"].Style.ForeColor = Color.DeepPink;
                    gcMultiRow1[e.RowIndex, "lblLineNum"].Style.ForeColor = Color.DeepPink;
                    gcMultiRow1[e.RowIndex, "lblBmnCode"].Style.ForeColor = Color.DeepPink;
                    gcMultiRow1[e.RowIndex, "lblHinCode"].Style.ForeColor = Color.DeepPink;
                    gcMultiRow1[e.RowIndex, "lblSftName"].Style.ForeColor = Color.DeepPink;
                }
                else
                {
                    gcMultiRow1[e.RowIndex, "lblName"].Style.ForeColor = Color.Blue;
                    gcMultiRow1[e.RowIndex, "lblLineNum"].Style.ForeColor = Color.Blue;
                    gcMultiRow1[e.RowIndex, "lblBmnCode"].Style.ForeColor = Color.Blue;
                    gcMultiRow1[e.RowIndex, "lblHinCode"].Style.ForeColor = Color.Blue;
                    gcMultiRow1[e.RowIndex, "lblSftName"].Style.ForeColor = Color.Blue;
                }
            }

            // 社員番号のとき社員名を表示します
            if (e.CellName == "txtShainNum")
            {
                // ChangeValueイベントを発生させない
                gl.ChangeValueStatus = false;

                // 氏名を初期化
                gcMultiRow1[e.RowIndex, "lblName"].Value = string.Empty;

                // 奉行データベースより社員名を取得して表示します
                if (Utility.NulltoStr(gcMultiRow1[e.RowIndex, "txtShainNum"].Value) != string.Empty)
                {
                    // 接続文字列取得
                    string sc = sqlControl.obcConnectSting.get(_dbName);
                    sqlControl.DataControl sdCon = new Common.sqlControl.DataControl(sc);

                    string bCode = gcMultiRow1[e.RowIndex, "txtShainNum"].Value.ToString().PadLeft(10, '0');
                    SqlDataReader dR = sdCon.free_dsReader(Utility.getEmployee(bCode));

                    while (dR.Read())
                    {
                        // 社員名表示
                        gcMultiRow1[e.RowIndex, "lblName"].Value = dR["Name"].ToString().Trim();
                        
                        //// ライン
                        //string val = Utility.getHisCategory(hArray, dR["JobTypeID"].ToString());     
                        //if (Utility.StrtoInt(val) == 0)
                        //{
                        //    gcMultiRow1[e.RowIndex, "lblLineNum"].Value = val;
                        //}
                        //else
                        //{
                        //    gcMultiRow1[e.RowIndex, "lblLineNum"].Value = Utility.StrtoInt(val);
                        //}

                        //// 部門
                        //val = Utility.getHisCategory(hArray, dR["DutyID"].ToString());
                        //if (Utility.StrtoInt(val) == 0)
                        //{
                        //    gcMultiRow1[e.RowIndex, "lblBmnCode"].Value = val;
                        //}
                        //else
                        //{
                        //    gcMultiRow1[e.RowIndex, "lblBmnCode"].Value = Utility.StrtoInt(val);
                        //}

                        //// 製品群
                        //val = Utility.getHisCategory(hArray, dR["QualificationGradeID"].ToString());
                        //if (Utility.StrtoInt(val) == 0)
                        //{
                        //    gcMultiRow1[e.RowIndex, "lblHinCode"].Value = val;
                        //}
                        //else
                        //{
                        //    gcMultiRow1[e.RowIndex, "lblHinCode"].Value = Utility.StrtoInt(val);
                        //}
                    }

                    dR.Close();
                    sdCon.Close();

                    // 2018/03/21 コメント化
                    //// ChangeValueイベントステータスをtrueに戻す
                    //gl.ChangeValueStatus = true;
                }

                // 2018/03/21
                // ChangeValueイベントステータスをtrueに戻す
                gl.ChangeValueStatus = true;
            }
            
            // 勤務体系（シフト）コード
            if (e.CellName == "txtSftCode")
            {
                // ChangeValueイベントを発生させない
                gl.ChangeValueStatus = false;

                // シフト名を初期化
                gcMultiRow1[e.RowIndex, "lblSftName"].Value = string.Empty;

                if (Utility.NulltoStr(gcMultiRow1[e.RowIndex, "txtSftCode"].Value) != string.Empty)
                {
                    string lName = string.Empty;

                    // 奉行SQLServer接続文字列取得
                    string sc = sqlControl.obcConnectSting.get(_dbName);
                    sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

                    // 登録済み勤務体系（シフト）コード検証
                    StringBuilder sb = new StringBuilder();
                    sb.Clear();
                    sb.Append("select LaborSystemCode, LaborSystemName from tbLaborSystem ");
                    sb.Append("where LaborSystemCode = '" + gcMultiRow1[e.RowIndex, "txtSftCode"].Value.ToString().PadLeft(4, '0') + "'");

                    SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

                    while (dR.Read())
                    {
                        lName = dR["LaborSystemName"].ToString();
                        break;
                    }

                    dR.Close();
                    sdCon.Close();

                    gcMultiRow1[e.RowIndex, "lblSftName"].Value = lName;
                }

                // 開始時間が対象シフトコードの日替わり時刻以前のときバックカラーを変更する：2018/09/18
                if (!chkChangeDayStartTime(e.RowIndex))
                {
                    // 対象のシフトコードの開始時間と異なるときバックカラーを変更する
                    // 日替わり時刻以前チェックがfalseのときのみ実行：2018/09/18
                    chkSftStartTime(e.RowIndex);
                }

                // ChangeValueイベントステータスをtrueに戻す
                gl.ChangeValueStatus = true;
            }

            // 出勤時間
            if (e.CellName == "txtSh" || e.CellName == "txtSm")
            {
                // 開始時間が対象シフトコードの日替わり時刻以前のときバックカラーを変更する：2018/09/18
                if (!chkChangeDayStartTime(e.RowIndex))
                {
                    // 対象のシフトコードの開始時間と異なるときバックカラーを変更する
                    // 日替わり時刻以前チェックがfalseのときのみ実行：2018/09/18
                    chkSftStartTime(e.RowIndex);
                }
            }

            // 取消チェックのとき
            if (e.CellName == "chkTorikeshi")
            {
                if (gcMultiRow1[e.RowIndex, "chkTorikeshi"].Value.ToString() == "True")
                {
                    gcMultiRow1.Rows[e.RowIndex].BackColor = SystemColors.Control;
                    gcMultiRow1[e.RowIndex, "chkOuen"].ReadOnly = true;
                    gcMultiRow1[e.RowIndex, "chkSft"].ReadOnly = true;
                    gcMultiRow1[e.RowIndex, "txtShainNum"].ReadOnly = true;
                    gcMultiRow1[e.RowIndex, "txtJiyu1"].ReadOnly = true;
                    gcMultiRow1[e.RowIndex, "txtJiyu2"].ReadOnly = true;
                    gcMultiRow1[e.RowIndex, "txtJiyu3"].ReadOnly = true;
                    gcMultiRow1[e.RowIndex, "txtSftCode"].ReadOnly = true;
                    gcMultiRow1[e.RowIndex, "txtSh"].ReadOnly = true;
                    gcMultiRow1[e.RowIndex, "txtSm"].ReadOnly = true;
                    gcMultiRow1[e.RowIndex, "txtEh"].ReadOnly = true;
                    gcMultiRow1[e.RowIndex, "txtEm"].ReadOnly = true;
                    gcMultiRow1[e.RowIndex, "txtZanRe1"].ReadOnly = true;
                    gcMultiRow1[e.RowIndex, "txtZanH1"].ReadOnly = true;
                    gcMultiRow1[e.RowIndex, "txtZanM1"].ReadOnly = true;
                    gcMultiRow1[e.RowIndex, "txtZanRe2"].ReadOnly = true;
                    gcMultiRow1[e.RowIndex, "txtZanH2"].ReadOnly = true;
                    gcMultiRow1[e.RowIndex, "txtZanM2"].ReadOnly = true;
                    //gcMultiRow1[e.RowIndex, "btnCell"].ReadOnly = true;

                    gcMultiRow1[e.RowIndex, "chkOuen"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "chkSft"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "lblName"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "txtShainNum"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "lblLineNum"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "lblBmnCode"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "lblHinCode"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "txtJiyu1"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "txtJiyu2"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "txtJiyu3"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "txtSftCode"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "lblSftName"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "txtSh"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "txtSm"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "txtEh"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "txtEm"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "txtZanRe1"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "txtZanH1"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "txtZanM1"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "txtZanRe2"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "txtZanH2"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "txtZanM2"].Style.ForeColor = Color.LightGray;
                    //gcMultiRow1[e.RowIndex, "btnCell"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "labelCell12"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "labelCell13"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "labelCell14"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "labelCell9"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "labelCell10"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "labelCell16"].Style.ForeColor = Color.LightGray;
                    gcMultiRow1[e.RowIndex, "labelCell17"].Style.ForeColor = Color.LightGray;
                }
                else
                {
                    //gcMultiRow1.Rows[e.RowIndex].BackColor = Color.Empty;
                    gcMultiRow1[e.RowIndex, "chkOuen"].ReadOnly = false;
                    gcMultiRow1[e.RowIndex, "chkSft"].ReadOnly = false;
                    gcMultiRow1[e.RowIndex, "txtShainNum"].ReadOnly = false;
                    gcMultiRow1[e.RowIndex, "txtJiyu1"].ReadOnly = false;
                    gcMultiRow1[e.RowIndex, "txtJiyu2"].ReadOnly = false;
                    gcMultiRow1[e.RowIndex, "txtJiyu3"].ReadOnly = false;
                    gcMultiRow1[e.RowIndex, "txtSftCode"].ReadOnly = false;
                    gcMultiRow1[e.RowIndex, "txtSh"].ReadOnly = false;
                    gcMultiRow1[e.RowIndex, "txtSm"].ReadOnly = false;
                    gcMultiRow1[e.RowIndex, "txtEh"].ReadOnly = false;
                    gcMultiRow1[e.RowIndex, "txtEm"].ReadOnly = false;
                    gcMultiRow1[e.RowIndex, "txtZanRe1"].ReadOnly = false;
                    gcMultiRow1[e.RowIndex, "txtZanH1"].ReadOnly = false;
                    gcMultiRow1[e.RowIndex, "txtZanM1"].ReadOnly = false;
                    gcMultiRow1[e.RowIndex, "txtZanRe2"].ReadOnly = false;
                    gcMultiRow1[e.RowIndex, "txtZanH2"].ReadOnly = false;
                    gcMultiRow1[e.RowIndex, "txtZanM2"].ReadOnly = false;
                    //gcMultiRow1[e.RowIndex, "btnCell"].ReadOnly = false;
                    

                    if (gcMultiRow1[e.RowIndex, "chkOuen"].Value.ToString() == "True")
                    {
                        gcMultiRow1[e.RowIndex, "lblName"].Style.ForeColor = Color.DeepPink;
                        gcMultiRow1[e.RowIndex, "lblLineNum"].Style.ForeColor = Color.DeepPink;
                        gcMultiRow1[e.RowIndex, "lblBmnCode"].Style.ForeColor = Color.DeepPink;
                        gcMultiRow1[e.RowIndex, "lblHinCode"].Style.ForeColor = Color.DeepPink;
                        gcMultiRow1[e.RowIndex, "lblSftName"].Style.ForeColor = Color.DeepPink;
                    }
                    else
                    {
                        gcMultiRow1[e.RowIndex, "lblName"].Style.ForeColor = Color.Blue;
                        gcMultiRow1[e.RowIndex, "lblLineNum"].Style.ForeColor = Color.Blue;
                        gcMultiRow1[e.RowIndex, "lblBmnCode"].Style.ForeColor = Color.Blue;
                        gcMultiRow1[e.RowIndex, "lblHinCode"].Style.ForeColor = Color.Blue;
                        gcMultiRow1[e.RowIndex, "lblSftName"].Style.ForeColor = Color.Blue;
                    }
                                        
                    gcMultiRow1[e.RowIndex, "chkOuen"].Style.ForeColor = Color.Empty;
                    gcMultiRow1[e.RowIndex, "chkSft"].Style.ForeColor = Color.Empty;
                    //gcMultiRow1[e.RowIndex, "lblName"].Style.ForeColor = Color.Blue;
                    gcMultiRow1[e.RowIndex, "txtShainNum"].Style.ForeColor = Color.Empty;
                    //gcMultiRow1[e.RowIndex, "lblLineNum"].Style.ForeColor = Color.Blue;
                    //gcMultiRow1[e.RowIndex, "lblBmnCode"].Style.ForeColor = Color.Blue;
                    //gcMultiRow1[e.RowIndex, "lblHinCode"].Style.ForeColor = Color.Blue;
                    gcMultiRow1[e.RowIndex, "txtJiyu1"].Style.ForeColor = Color.Empty;
                    gcMultiRow1[e.RowIndex, "txtJiyu2"].Style.ForeColor = Color.Empty;
                    gcMultiRow1[e.RowIndex, "txtJiyu3"].Style.ForeColor = Color.Empty;
                    gcMultiRow1[e.RowIndex, "txtSftCode"].Style.ForeColor = Color.Empty;
                    //gcMultiRow1[e.RowIndex, "lblSftName"].Style.ForeColor = Color.Blue;
                    gcMultiRow1[e.RowIndex, "txtSh"].Style.ForeColor = Color.Empty;
                    gcMultiRow1[e.RowIndex, "txtSm"].Style.ForeColor = Color.Empty;
                    gcMultiRow1[e.RowIndex, "txtEh"].Style.ForeColor = Color.Empty;
                    gcMultiRow1[e.RowIndex, "txtEm"].Style.ForeColor = Color.Empty;
                    gcMultiRow1[e.RowIndex, "txtZanRe1"].Style.ForeColor = Color.Empty;
                    gcMultiRow1[e.RowIndex, "txtZanH1"].Style.ForeColor = Color.Empty;
                    gcMultiRow1[e.RowIndex, "txtZanM1"].Style.ForeColor = Color.Empty;
                    gcMultiRow1[e.RowIndex, "txtZanRe2"].Style.ForeColor = Color.Empty;
                    gcMultiRow1[e.RowIndex, "txtZanH2"].Style.ForeColor = Color.Empty;
                    gcMultiRow1[e.RowIndex, "txtZanM2"].Style.ForeColor = Color.Empty;
                    //gcMultiRow1[e.RowIndex, "btnCell"].Style.ForeColor = Color.Blue;
                    gcMultiRow1[e.RowIndex, "labelCell12"].Style.ForeColor = Color.Empty;
                    gcMultiRow1[e.RowIndex, "labelCell13"].Style.ForeColor = Color.Empty;
                    gcMultiRow1[e.RowIndex, "labelCell14"].Style.ForeColor = Color.Empty;
                    gcMultiRow1[e.RowIndex, "labelCell9"].Style.ForeColor = Color.Empty;
                    gcMultiRow1[e.RowIndex, "labelCell10"].Style.ForeColor = Color.Empty;
                    gcMultiRow1[e.RowIndex, "labelCell16"].Style.ForeColor = Color.Empty;
                    gcMultiRow1[e.RowIndex, "labelCell17"].Style.ForeColor = Color.Empty;
                }
            }
        }

        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     対象のシフトコードの開始時間と異なるときバックカラーを変更する </summary>
        /// <param name="r">
        ///     MultiRowの行インデックス </param>
        /// <param name="sftCode">
        ///     シフトコード </param>
        ///----------------------------------------------------------------------------------
        private void chkSftStartTime()
        {
            for (int r = 0; r < gcMultiRow1.RowCount; r++)
            {              
                string sG = Utility.NulltoStr(gcMultiRow1[r, "txtSh"].Value) + Utility.NulltoStr(gcMultiRow1[r, "txtSm"].Value);

                // 出勤時間空白は対象外とする
                if (sG != string.Empty)
                {
                    // 対象のシフトコード取得する
                    string sftCode = string.Empty;
                    DateTime sDt = DateTime.Now;

                    if (Utility.NulltoStr(gcMultiRow1[r, "txtSftCode"].Value) != string.Empty)
                    {
                        // 変更シフトコードあり
                        sftCode = gcMultiRow1[r, "txtSftCode"].Value.ToString().PadLeft(4, '0');
                    }
                    else if (Utility.NulltoStr(gcMultiRow2[0, "txtSftCode"].Value) != string.Empty)
                    {
                        // 標準シフトコード
                        sftCode = gcMultiRow2[0, "txtSftCode"].Value.ToString().PadLeft(4, '0');
                    }

                    // 有効なシフトコードが存在するとき
                    if (sftCode != string.Empty)
                    {
                        // 奉行SQLServer接続文字列取得
                        string sc = sqlControl.obcConnectSting.get(_dbName);
                        sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

                        // 勤務体系（シフト）コード情報取得
                        StringBuilder sb = new StringBuilder();
                        sb.Clear();
                        sb.Append("SELECT tbLaborSystem.LaborSystemID,tbLaborSystem.LaborSystemCode,");
                        sb.Append("tbLaborSystem.LaborSystemName,tbLaborTimeSpanRule.StartTime,");
                        sb.Append("tbLaborTimeSpanRule.EndTime ");
                        sb.Append("FROM tbLaborSystem inner join tbLaborTimeSpanRule ");
                        sb.Append("on tbLaborSystem.LaborSystemID = tbLaborTimeSpanRule.LaborSystemID ");
                        sb.Append("where tbLaborTimeSpanRule.LaborTimeSpanRuleType = 1 ");
                        sb.Append("and tbLaborSystem.LaborSystemCode = '").Append(sftCode).Append("'");

                        SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

                        bool bn = false;

                        while (dR.Read())
                        {
                            sDt = DateTime.Parse(dR["StartTime"].ToString());
                            bn = true;
                            break;
                        }

                        dR.Close();
                        sdCon.Close();

                        string sS = string.Empty;

                        if (bn)
                        {
                            // 開始時刻を取得したとき
                            sS = sDt.Hour.ToString().PadLeft(2, '0') + sDt.Minute.ToString().PadLeft(2, '0');
                        } 
                        
                        sG = Utility.NulltoStr(gcMultiRow1[r, "txtSh"].Value).PadLeft(2, '0') + Utility.NulltoStr(gcMultiRow1[r, "txtSm"].Value).PadLeft(2, '0');

                        if (!bn)
                        {
                            gcMultiRow1[r, "txtSh"].Style.BackColor = SystemColors.Window;
                            gcMultiRow1[r, "labelCell12"].Style.BackColor = SystemColors.Window;
                            gcMultiRow1[r, "txtSm"].Style.BackColor = SystemColors.Window;
                        }
                        else if (sS != sG)
                        {
                            gcMultiRow1[r, "txtSh"].Style.BackColor = Color.LightPink;
                            gcMultiRow1[r, "labelCell12"].Style.BackColor = Color.LightPink;
                            gcMultiRow1[r, "txtSm"].Style.BackColor = Color.LightPink;
                        }
                        else
                        {
                            gcMultiRow1[r, "txtSh"].Style.BackColor = SystemColors.Window;
                            gcMultiRow1[r, "labelCell12"].Style.BackColor = SystemColors.Window;
                            gcMultiRow1[r, "txtSm"].Style.BackColor = SystemColors.Window;
                        }
                    }
                }
            }
        }

        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     対象のシフトコードの開始時間と異なるときバックカラーを変更する </summary>
        /// <param name="r">
        ///     MultiRowの行インデックス </param>
        /// <param name="sftCode">
        ///     シフトコード </param>
        ///----------------------------------------------------------------------------------
        private void chkSftStartTime(int r)
        {
            string sG = Utility.NulltoStr(gcMultiRow1[r, "txtSh"].Value) + Utility.NulltoStr(gcMultiRow1[r, "txtSm"].Value);

            // 出勤時間空白は対象外とする
            if (sG != string.Empty)
            {
                // 対象のシフトコード取得する
                string sftCode = string.Empty;
                DateTime sDt = DateTime.Now;

                if (Utility.NulltoStr(gcMultiRow1[r, "txtSftCode"].Value) != string.Empty)
                {
                    // 変更シフトコードあり
                    sftCode = gcMultiRow1[r, "txtSftCode"].Value.ToString().PadLeft(4, '0');
                }
                else if (Utility.NulltoStr(gcMultiRow2[0, "txtSftCode"].Value) != string.Empty)
                {
                    // 標準シフトコード
                    sftCode = gcMultiRow2[0, "txtSftCode"].Value.ToString().PadLeft(4, '0');
                }

                // 有効なシフトコードが存在するとき
                if (sftCode != string.Empty)
                {
                    // 奉行SQLServer接続文字列取得
                    string sc = sqlControl.obcConnectSting.get(_dbName);
                    sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

                    // 勤務体系（シフト）コード情報取得
                    StringBuilder sb = new StringBuilder();
                    sb.Clear();
                    sb.Append("SELECT tbLaborSystem.LaborSystemID,tbLaborSystem.LaborSystemCode,");
                    sb.Append("tbLaborSystem.LaborSystemName,tbLaborTimeSpanRule.StartTime,");
                    sb.Append("tbLaborTimeSpanRule.EndTime ");
                    sb.Append("FROM tbLaborSystem inner join tbLaborTimeSpanRule ");
                    sb.Append("on tbLaborSystem.LaborSystemID = tbLaborTimeSpanRule.LaborSystemID ");
                    sb.Append("where tbLaborTimeSpanRule.LaborTimeSpanRuleType = 1 ");
                    sb.Append("and tbLaborSystem.LaborSystemCode = '").Append(sftCode).Append("'");

                    SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

                    bool bn = false;

                    while (dR.Read())
                    {
                        bn = true;
                        sDt = DateTime.Parse(dR["StartTime"].ToString());
                        break;
                    }

                    dR.Close();
                    sdCon.Close();

                    string sS = sDt.Hour.ToString().PadLeft(2, '0') + sDt.Minute.ToString().PadLeft(2, '0');
                    sG = Utility.NulltoStr(gcMultiRow1[r, "txtSh"].Value).PadLeft(2, '0') + Utility.NulltoStr(gcMultiRow1[r, "txtSm"].Value).PadLeft(2, '0');

                    if (!bn)
                    {
                        gcMultiRow1[r, "txtSh"].Style.BackColor = SystemColors.Window;
                        gcMultiRow1[r, "labelCell12"].Style.BackColor = SystemColors.Window;
                        gcMultiRow1[r, "txtSm"].Style.BackColor = SystemColors.Window;
                    }
                    else if (sS != sG)
                    {
                        gcMultiRow1[r, "txtSh"].Style.BackColor = Color.LightPink;
                        gcMultiRow1[r, "labelCell12"].Style.BackColor = Color.LightPink;
                        gcMultiRow1[r, "txtSm"].Style.BackColor = Color.LightPink;
                    }
                    else
                    {
                        gcMultiRow1[r, "txtSh"].Style.BackColor = SystemColors.Window;
                        gcMultiRow1[r, "labelCell12"].Style.BackColor = SystemColors.Window;
                        gcMultiRow1[r, "txtSm"].Style.BackColor = SystemColors.Window;
                    }
                }
            }
        }
        
        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     対象のシフトコードの日替わり時刻以前の開始時刻のときバックカラーを変更する ：
        ///     2018/09/18</summary>
        /// <param name="r">
        ///     MultiRowの行インデックス </param>
        /// <param name="sftCode">
        ///     シフトコード </param>
        ///----------------------------------------------------------------------------------
        private bool chkChangeDayStartTime(int r)
        {
            bool rtn = false;

            string stSt = Utility.NulltoStr(gcMultiRow1[r, "txtSh"].Value) +
                        Utility.NulltoStr(gcMultiRow1[r, "txtSm"].Value);

            // 出勤時間空白は対象外とする
            if (stSt == string.Empty)
            {
                return false;
            }

            int sG = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[r, "txtSh"].Value)) * 100 +
                     Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[r, "txtSm"].Value));


            // 対象のシフトコード取得する
            string sftCode = string.Empty;
            DateTime cDt = DateTime.Now;

            if (Utility.NulltoStr(gcMultiRow1[r, "txtSftCode"].Value) != string.Empty)
            {
                // 変更シフトコードあり
                sftCode = gcMultiRow1[r, "txtSftCode"].Value.ToString().PadLeft(4, '0');
            }
            else if (Utility.NulltoStr(gcMultiRow2[0, "txtSftCode"].Value) != string.Empty)
            {
                // 標準シフトコード
                sftCode = gcMultiRow2[0, "txtSftCode"].Value.ToString().PadLeft(4, '0');
            }

            // 有効なシフトコードが存在するとき
            if (sftCode != string.Empty)
            {
                // 奉行SQLServer接続文字列取得
                string sc = sqlControl.obcConnectSting.get(_dbName);
                sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

                // 勤務体系（シフト）コード情報取得
                StringBuilder sb = new StringBuilder();
                sb.Clear();
                sb.Append("SELECT tbLaborSystem.LaborSystemID,tbLaborSystem.LaborSystemCode,");
                sb.Append("tbLaborSystem.LaborSystemName,tbLaborTimeSpanRule.StartTime,");
                sb.Append("tbLaborTimeSpanRule.EndTime, tbLaborSystem.DayChangeTime ");
                sb.Append("FROM tbLaborSystem inner join tbLaborTimeSpanRule ");
                sb.Append("on tbLaborSystem.LaborSystemID = tbLaborTimeSpanRule.LaborSystemID ");
                sb.Append("where tbLaborTimeSpanRule.LaborTimeSpanRuleType = 1 ");
                sb.Append("and tbLaborSystem.LaborSystemCode = '").Append(sftCode).Append("'");

                SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

                bool bn = false;

                while (dR.Read())
                {
                    bn = true;
                    cDt = DateTime.Parse(dR["DayChangeTime"].ToString());
                    break;
                }

                dR.Close();
                sdCon.Close();

                int sS = cDt.Hour * 100 + cDt.Minute;

                if (!bn)
                {
                    gcMultiRow1[r, "txtSh"].Style.BackColor = SystemColors.Window;
                    gcMultiRow1[r, "labelCell12"].Style.BackColor = SystemColors.Window;
                    gcMultiRow1[r, "txtSm"].Style.BackColor = SystemColors.Window;
                    rtn = false;
                }
                else if (sS > sG)
                {
                    // 開始時刻が日替わり時刻以前のとき
                    gcMultiRow1[r, "txtSh"].Style.BackColor = Color.Red;
                    gcMultiRow1[r, "labelCell12"].Style.BackColor = Color.Red;
                    gcMultiRow1[r, "txtSm"].Style.BackColor = Color.Red;
                    rtn = true;
                }
                else
                {
                    gcMultiRow1[r, "txtSh"].Style.BackColor = SystemColors.Window;
                    gcMultiRow1[r, "labelCell12"].Style.BackColor = SystemColors.Window;
                    gcMultiRow1[r, "txtSm"].Style.BackColor = SystemColors.Window;
                    rtn = false;
                }
            }

            return rtn;
        }
        
        private void gcMultiRow1_EditingControlShowing(object sender, EditingControlShowingEventArgs e)
        {
            if (e.Control is TextBoxEditingControl)
            {
                //イベントハンドラが複数回追加されてしまうので最初に削除する
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress2);
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress3);

                // 数字のみ入力可能とする
                if (gcMultiRow1.CurrentCell.Name == "txtShainNum" || gcMultiRow1.CurrentCell.Name == "txtJiyu1" ||
                    gcMultiRow1.CurrentCell.Name == "txtJiyu2" || gcMultiRow1.CurrentCell.Name == "txtJiyu3" ||
                    gcMultiRow1.CurrentCell.Name == "txtSftCode" || gcMultiRow1.CurrentCell.Name == "txtSh" ||
                    gcMultiRow1.CurrentCell.Name == "txtSm" || gcMultiRow1.CurrentCell.Name == "txtEh" ||
                    gcMultiRow1.CurrentCell.Name == "txtEm" || 
                    gcMultiRow1.CurrentCell.Name == "txtZanRe1" || gcMultiRow1.CurrentCell.Name == "txtZanH1" ||
                    gcMultiRow1.CurrentCell.Name == "txtZanRe2" || gcMultiRow1.CurrentCell.Name == "txtZanH2")
                {
                    //イベントハンドラを追加する
                    e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
                }

                // 残業時は「０」「５」のみ入力可能とする
                if (gcMultiRow1.CurrentCell.Name == "txtZanM1" || gcMultiRow1.CurrentCell.Name == "txtZanM2")
                {
                    //イベントハンドラを追加する
                    e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress3);
                }
            }
        }

        private void gcMultiRow2_EditingControlShowing(object sender, EditingControlShowingEventArgs e)
        {
            if (e.Control is TextBoxEditingControl)
            {
                //イベントハンドラが複数回追加されてしまうので最初に削除する
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress2);

                // 数字のみ入力可能とする
                if (gcMultiRow2.CurrentCell.Name == "txtSftCode" ||
                    gcMultiRow2.CurrentCell.Name == "txtYear" || gcMultiRow2.CurrentCell.Name == "txtMonth" ||
                    gcMultiRow2.CurrentCell.Name == "txtDay")
                {
                    //イベントハンドラを追加する
                    e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
                }
            }
        }

        private void gcMultiRow2_CellValueChanged(object sender, CellEventArgs e)
        {
            if (!gl.ChangeValueStatus) return;

            if (e.RowIndex < 0) return;

            //// 過去データ表示のときは終了
            //if (dID != string.Empty) return;

            // 部署コードのとき部署名を表示します
            if (e.CellName == "txtBushoCode")
            {
                // ChangeValueイベントを発生させない
                gl.ChangeValueStatus = false;

                // 部署名を初期化
                gcMultiRow2[e.RowIndex, "lblShozoku"].Value = string.Empty;

                // 奉行データベースより部署名を取得して表示します
                if (Utility.NulltoStr(gcMultiRow2[e.RowIndex, "txtBushoCode"].Value) != string.Empty)
                {
                    string dName = string.Empty;
                    if (getDepartMentName(out dName, gcMultiRow2[e.RowIndex, "txtBushoCode"].Value.ToString(), e.RowIndex))
                    {
                        gcMultiRow2[e.RowIndex, "lblShozoku"].Value = dName;
                    }

                    // ChangeValueイベントステータスをtrueに戻す
                    gl.ChangeValueStatus = true;
                }
            }

            // 勤務体系（シフト）コード
            if (e.CellName == "txtSftCode")
            {
                // ChangeValueイベントを発生させない
                gl.ChangeValueStatus = false;

                // 対象年月日
                DateTime eDate;
                string sDate = Utility.NulltoStr(gcMultiRow2[0, "txtYear"].Value) + "/" +
                                  Utility.NulltoStr(gcMultiRow2[0, "txtMonth"].Value) + "/" + 
                                  Utility.NulltoStr(gcMultiRow2[0, "txtDay"].Value);

                string sHol = global.FLGOFF;
                if (DateTime.TryParse(sDate, out eDate))
                {
                    // 該当日が休日か調べる
                    if (dts.休日.Any(a => a.年月日 == eDate))
                    {
                        sHol = global.FLGON;
                    }
                }

                // 部署名を初期化
                gcMultiRow2[e.RowIndex, "lblShtName"].Value = string.Empty;

                // 就業奉行の勤務体系よりシフト名と開始終了時刻を取得して表示します
                if (Utility.NulltoStr(gcMultiRow2[e.RowIndex, "txtSftCode"].Value) != string.Empty)
                {           
                    // 奉行SQLServer接続文字列取得
                    string sc = sqlControl.obcConnectSting.get(_dbName);
                    sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

                    // 勤務体系（シフト）取得
                    StringBuilder sb = new StringBuilder();
                    sb.Clear();
                    sb.Append("select tbLaborSystem.LaborSystemCode, LaborSystemName, tbLaborSystem.LatterHalfStartTime, tbLaborSystem.FirstHalfEndTime,");
                    sb.Append("tbLaborSystem.DayChangeTime,a.StartTime,a.EndTime ");
                    sb.Append("from tbLaborSystem left join ");
                    sb.Append("(select * from tbLaborTimeSpanRule where LaborTimeItemID = 1) as a ");
                    sb.Append("on tbLaborSystem.LaborSystemID = a.LaborSystemID ");
                    sb.Append("where tbLaborSystem.LaborSystemCode = '" + gcMultiRow2[e.RowIndex, "txtSftCode"].Value.ToString().PadLeft(4, '0') + "'");

                    SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

                    bool bn = false;
                    DateTime dtStart = DateTime.Now;        // 開始時刻
                    DateTime dtEnd = DateTime.Now;          // 終了時刻
                    DateTime dtLatterStart = DateTime.Now;  // 後半開始時刻
                    DateTime dtFirstEnd = DateTime.Now;     // 前半終了時刻
                    DateTime dtChange = DateTime.Now;       // 日替わり時刻
                    string sftName = string.Empty;          // 勤務体系名称

                    lblStartTime.Text = string.Empty;
                    lblEndTime.Text = string.Empty;

                    while (dR.Read())
                    {
                        sftName = Utility.NulltoStr(dR["LaborSystemName"]);

                        if (!(dR["StartTime"] is DBNull))
                        {
                            bn = true;
                            dtStart = (DateTime)dR["StartTime"];
                        }

                        if (!(dR["EndTime"] is DBNull))
                        {
                            bn = true;
                            dtEnd = (DateTime)dR["EndTime"];
                        }

                        dtLatterStart = (DateTime)dR["LatterHalfStartTime"];
                        dtFirstEnd = (DateTime)dR["FirstHalfEndTime"];
                        dtChange = (DateTime)dR["DayChangeTime"];
                        break;
                    }

                    dR.Close();
                    sdCon.Close();

                    string msg = sftName;

                    if (bn)
                    {
                        msg +=  " " + dtStart.Hour.ToString().PadLeft(2, '0') + ":" + dtStart.Minute.ToString().PadLeft(2, '0') + "～";

                        lblStartTime.Text = dtStart.Hour.ToString().PadLeft(2, '0') + ":" + dtStart.Minute.ToString().PadLeft(2, '0') + "～" + 
                                            dtFirstEnd.Hour.ToString().PadLeft(2, '0') + ":" + dtFirstEnd.Minute.ToString().PadLeft(2, '0');
                    }

                    if (bn)
                    {
                        msg += dtEnd.Hour.ToString().PadLeft(2, '0') + ":" + dtEnd.Minute.ToString().PadLeft(2, '0');

                        lblEndTime.Text = dtLatterStart.Hour.ToString().PadLeft(2, '0') + ":" + dtLatterStart.Minute.ToString().PadLeft(2, '0') + "～" +
                                            dtEnd.Hour.ToString().PadLeft(2, '0') + ":" + dtEnd.Minute.ToString().PadLeft(2, '0');
                    }

                    // シフト名
                    gcMultiRow2[0, "lblShtName"].Value = msg;

                    // ChangeValueイベントステータスをtrueに戻す
                    gl.ChangeValueStatus = true;
                }

                // 対象のシフトコードの開始時間と異なるときバックカラーを変更する
                chkSftStartTime();
            }

            // 年月日
            if (e.CellName == "txtYear" || e.CellName == "txtMonth" || e.CellName == "txtDay")
            {
                // 曜日
                DateTime eDate;
                int tYY = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow2[e.RowIndex, "txtYear"].Value));
                string sDate = tYY.ToString() + "/" + Utility.NulltoStr(gcMultiRow2[e.RowIndex, "txtMonth"].Value) + "/" +
                        Utility.NulltoStr(gcMultiRow2[e.RowIndex, "txtDay"].Value);

                // 存在する日付と認識された場合、曜日を表示する
                if (DateTime.TryParse(sDate, out eDate))
                {
                    gcMultiRow2[e.RowIndex, "lblWeek"].Value = ("日月火水木金土").Substring(int.Parse(eDate.DayOfWeek.ToString("d")), 1);
                }
                else
                {
                    gcMultiRow2[e.RowIndex, "lblWeek"].Value = string.Empty;
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
            if (Utility.NulltoStr(gcMultiRow2[r, "txtBushoCode"].Value) != string.Empty)
            {
                string b = string.Empty;

                // 検索用部署コード
                if (Utility.StrtoInt(gcMultiRow2[r, "txtBushoCode"].Value.ToString()) != global.flgOff)
                {
                    b = gcMultiRow2[r, "txtBushoCode"].Value.ToString().Trim().PadLeft(15, '0');
                }
                else
                {
                    b = gcMultiRow2[r, "txtBushoCode"].Value.ToString().Trim().PadRight(15, ' ');
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

        ///-------------------------------------------------------------------
        /// <summary>
        ///     ライン・部門・製品群コード配列取得   </summary>
        /// <returns>
        ///     ID,コード配列</returns>
        ///-------------------------------------------------------------------
        private string[] getCategoryArray()
        {
            // 接続文字列取得
            string sc = sqlControl.obcConnectSting.get(_dbName);
            sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            StringBuilder sb = new StringBuilder();
            sb.Append("select CategoryID, CategoryCode from tbHistoryDivisionCategory");
            SqlDataReader dr = sdCon.free_dsReader(sb.ToString());

            int iX = 0;
            string[] hArray = new string[1];

            while (dr.Read())
            {
                if (iX > 0)
                {
                    Array.Resize(ref hArray, iX + 1);
                }

                hArray[iX] = dr["CategoryID"].ToString() + "," + dr["CategoryCode"].ToString();
                iX++;
            }

            dr.Close();
            sdCon.Close();

            return hArray;
        }

        private void gcMultiRow2_CellEnter(object sender, CellEventArgs e)
        {
            if (gcMultiRow2.EditMode == EditMode.EditProgrammatically)
            {
                gcMultiRow2.BeginEdit(true);
            }
        }

        private void gcMultiRow1_CellEnter(object sender, CellEventArgs e)
        {
            if (gcMultiRow1.EditMode == EditMode.EditProgrammatically)
            {
                gcMultiRow1.BeginEdit(true);
            }
        }

        private void lnkOuen_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (dts.過去応援移動票ヘッダ.Count == 0)
            {
                MessageBox.Show("応援移動票データがありません", "応援移動票データ登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // 応援移動票データ作成
            frmOuenCorrectPast frmC = new frmOuenCorrectPast(_dbName, _comName, string.Empty, false);
            frmC.ShowDialog();

            // 応援移動票データ読み込み
            getOuenDataSet();
        }

        private void gcMultiRow1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            string colName = gcMultiRow1.CurrentCell.Name;            

            if (colName == "chkOuen" || colName == "chkTorikeshi")
            {
                if (gcMultiRow1.IsCurrentCellDirty)
                {
                    gcMultiRow1.CommitEdit(DataErrorContexts.Commit);
                    gcMultiRow1.Refresh();
                }
            }
        }

        private void gcMultiRow1_CellLeave(object sender, CellEventArgs e)
        {
            // 2018/03/21
            if (gcMultiRow1.EditMode == EditMode.EditProgrammatically)
            {
                gcMultiRow1.EndEdit();
            }
        }

        private void gcMultiRow1_CellContentClick(object sender, CellEventArgs e)
        {
            if (gcMultiRow1[e.RowIndex, "chkTorikeshi"].Value.ToString() == "True")
            {
                return;
            }

            //if (e.CellName == "btnCell")
            //{
            //    //カレントデータの更新
            //    CurDataUpDate(cID[cI]);
                
            //    int sMID = Utility.StrtoInt(gcMultiRow1[e.RowIndex, "txtID"].Value.ToString());

            //    if (dts.過去勤務票明細.Any(a => a.ID == sMID))
            //    {
            //        var s = dts.過去勤務票明細.Single(a => a.ID == sMID);
            //        string kID = s.帰宅後勤務ID;
            //        frmKitakugo frm = new frmKitakugo(_dbName, sMID, kID, hArray, bs, true);
            //        frm.ShowDialog();

            //        // 帰宅後勤務データ再読み込み
            //        tAdp.Fill(dts.帰宅後勤務);

            //        //// 勤務票明細再読み込み
            //        //adpMn.勤務票明細TableAdapter.Fill(dts.過去勤務票明細);

            //        // データ再表示
            //        showOcrData(cI);
            //    }
            //}
        }

        private void lnkErrCheck_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 非ログ書き込み状態とする：2015/09/25
            editLogStatus = false;

            // OCRDataクラス生成
            OCRData ocr = new OCRData(_dbName, bs);

            // エラーチェックを実行
            if (getErrData(dID, ocr))
            {
                MessageBox.Show("エラーはありませんでした", "エラーチェック", MessageBoxButtons.OK, MessageBoxIcon.Information);
                gcMultiRow1.CurrentCell = null;
                gcMultiRow2.CurrentCell = null;

                // データ表示
                showOcrData(dID);
            }
            else
            {
                // カレントインデックスをエラーありインデックスで更新
                cI = ocr._errHeaderIndex;

                // データ表示
                showOcrData(dID);

                // エラー表示
                ErrShow(ocr);
            }
        }

        private void lnkDataMake_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 非ログ書き込み状態とする
            editLogStatus = false;

            // 就業奉行用CSVデータ出力
            textDataMake();
        }

        private void lnkRtn_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 非ログ書き込み状態とする
            editLogStatus = false;

            // フォームを閉じる
            this.Tag = END_BUTTON;
            this.Close();
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            Image img;

            img = Image.FromFile(_img);
            e.Graphics.DrawImage(img, 0, 0);
            e.HasMorePages = false;
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("画像を印刷します。よろしいですか？", "印刷確認", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
            {
                return;
            }

            // 印刷実行
            printDocument1.Print();
        }

        private void gcMultiRow2_CellLeave(object sender, CellEventArgs e)
        {
            // 2018/03/21
            if (gcMultiRow2.EditMode == EditMode.EditProgrammatically)
            {
                gcMultiRow2.EndEdit();
            }

        }
    }
}
