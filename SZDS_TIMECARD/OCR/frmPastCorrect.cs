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
    public partial class frmPastCorrect : Form
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
        /// ------------------------------------------------------------
        public frmPastCorrect(string dbName, string comName, string sID)
        {
            InitializeComponent();

            _dbName = dbName;       // データベース名
            _comName = comName;     // 会社名
            dID = sID;              // 処理モード
            
            // テーブルアダプターマネージャーに勤務票ヘッダ、明細テーブルアダプターを割り付ける
            adpMn.過去勤務票ヘッダTableAdapter = hAdp;
            adpMn.過去勤務票明細TableAdapter = iAdp;

            // 休日テーブル読み込み
            kAdp.Fill(dts.休日);
        }

        // データアダプターオブジェクト
        DataSet1TableAdapters.TableAdapterManager adpMn = new DataSet1TableAdapters.TableAdapterManager();
        DataSet1TableAdapters.過去勤務票ヘッダTableAdapter hAdp = new DataSet1TableAdapters.過去勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.過去勤務票明細TableAdapter iAdp = new DataSet1TableAdapters.過去勤務票明細TableAdapter();
        DataSet1TableAdapters.休日TableAdapter kAdp = new DataSet1TableAdapters.休日TableAdapter();

        DataSet1TableAdapters.応援移動票ヘッダTableAdapter ohAdp = new DataSet1TableAdapters.応援移動票ヘッダTableAdapter();
        DataSet1TableAdapters.応援移動票明細TableAdapter omAdp = new DataSet1TableAdapters.応援移動票明細TableAdapter();

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
            
            // データセットへデータを読み込みます
            getDataSet();

            // 部署別残業理由シートデータ配列取得
            bs = new xlsData();
            bs.zArray = bs.getShiftCode();

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
            foreach (var t in dts.勤務票ヘッダ.OrderBy(a => a.ID))
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

            this.gcMultiRow2.AllowUserToAddRows = false;                    //手動による行追加を禁止する
            this.gcMultiRow2.AllowUserToDeleteRows = false;                 //手動による行削除を禁止する
            this.gcMultiRow2.Rows.Clear();                                  //行数をクリア
            this.gcMultiRow2.RowCount = 1;                                  //行数を設定

            //multirow編集モード
            gcMultiRow1.EditMode = EditMode.EditProgrammatically;

            this.gcMultiRow1.AllowUserToAddRows = false;                    //手動による行追加を禁止する
            this.gcMultiRow1.AllowUserToDeleteRows = false;                 //手動による行削除を禁止する
            this.gcMultiRow1.Rows.Clear();                                  //行数をクリア
            this.gcMultiRow1.RowCount = global.MAX_GYO;                     //行数を設定
        }

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
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
            if (e.KeyChar != '0' && e.KeyChar != '5' && e.KeyChar != '\b' && e.KeyChar != '\t')
                e.Handled = true;
        }

        private void frmCorrect_Shown(object sender, EventArgs e)
        {
            
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

        private void btnRtn_Click(object sender, EventArgs e)
        {
        }

        private void frmCorrect_FormClosing(object sender, FormClosingEventArgs e)
        {
            //「受入データ作成終了」「勤務票データなし」以外での終了のとき
            if (this.Tag.ToString() != END_MAKEDATA && this.Tag.ToString() != END_NODATA)
            {
                if (MessageBox.Show("終了します。よろしいですか", "終了確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    e.Cancel = true;
                    return;
                }
            }

            // データベース更新
            adpMn.UpdateAll(dts);

            // 解放する
            this.Dispose();
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
                linkLabel3.Enabled = true;

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
                linkLabel3.Enabled = false;

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
                        
                        // ライン
                        string val = Utility.getHisCategory(hArray, dR["JobTypeID"].ToString());     
                        if (Utility.StrtoInt(val) == 0)
                        {
                            gcMultiRow1[e.RowIndex, "lblLineNum"].Value = val;
                        }
                        else
                        {
                            gcMultiRow1[e.RowIndex, "lblLineNum"].Value = Utility.StrtoInt(val);
                        }

                        // 部門
                        val = Utility.getHisCategory(hArray, dR["DutyID"].ToString());
                        if (Utility.StrtoInt(val) == 0)
                        {
                            gcMultiRow1[e.RowIndex, "lblBmnCode"].Value = val;
                        }
                        else
                        {
                            gcMultiRow1[e.RowIndex, "lblBmnCode"].Value = Utility.StrtoInt(val);
                        }

                        // 製品群
                        val = Utility.getHisCategory(hArray, dR["QualificationGradeID"].ToString());
                        if (Utility.StrtoInt(val) == 0)
                        {
                            gcMultiRow1[e.RowIndex, "lblHinCode"].Value = val;
                        }
                        else
                        {
                            gcMultiRow1[e.RowIndex, "lblHinCode"].Value = Utility.StrtoInt(val);
                        }
                    }

                    dR.Close();
                    sdCon.Close();

                    // ChangeValueイベントステータスをtrueに戻す
                    gl.ChangeValueStatus = true;
                }
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
                
                // 対象のシフトコードの開始時間と異なるときバックカラーを変更する
                chkSftStartTime(e.RowIndex);

                // ChangeValueイベントステータスをtrueに戻す
                gl.ChangeValueStatus = true;
            }

            // 出勤時間
            if (e.CellName == "txtSh" || e.CellName == "txtSm")
            {
                // 対象のシフトコードの開始時間と異なるときバックカラーを変更する
                chkSftStartTime(e.RowIndex);
            }


            // 取消チェックのとき
            if (e.CellName == "chkTorikeshi")
            {
                if (gcMultiRow1[e.RowIndex, "chkTorikeshi"].Value.ToString() == "True")
                {
                    gcMultiRow1.Rows[e.RowIndex].BackColor = SystemColors.Control;

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

                        while (dR.Read())
                        {
                            sDt = DateTime.Parse(dR["StartTime"].ToString());
                            break;
                        }

                        dR.Close();
                        sdCon.Close();

                        string sS = sDt.Hour.ToString().PadLeft(2, '0') + sDt.Minute.ToString().PadLeft(2, '0');
                        sG = Utility.NulltoStr(gcMultiRow1[r, "txtSh"].Value).PadLeft(2, '0') + Utility.NulltoStr(gcMultiRow1[r, "txtSm"].Value).PadLeft(2, '0');

                        if (sS != sG)
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

                    while (dR.Read())
                    {
                        sDt = DateTime.Parse(dR["StartTime"].ToString());
                        break;
                    }

                    dR.Close();
                    sdCon.Close();

                    string sS = sDt.Hour.ToString().PadLeft(2, '0') + sDt.Minute.ToString().PadLeft(2, '0');
                    sG = Utility.NulltoStr(gcMultiRow1[r, "txtSh"].Value).PadLeft(2, '0') + Utility.NulltoStr(gcMultiRow1[r, "txtSm"].Value).PadLeft(2, '0');

                    if (sS != sG)
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

                // 部署別勤務体系配列よりシフト名を取得して表示します
                if (Utility.NulltoStr(gcMultiRow2[e.RowIndex, "txtSftCode"].Value) != string.Empty)
                {
                    string dName = string.Empty;
                    if (bs.getBushoSft(out dName, gcMultiRow2[e.RowIndex, "txtBushoCode"].Value.ToString(), gcMultiRow2[e.RowIndex, "txtSftCode"].Value.ToString(), sHol))
                    {
                        gcMultiRow2[e.RowIndex, "lblShtName"].Value = dName;
                    }

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
            
        }

        private void gcMultiRow1_CellContentClick(object sender, CellEventArgs e)
        {
            if (gcMultiRow1[e.RowIndex, "chkTorikeshi"].Value == null)
            {
                return;
            }

            if (gcMultiRow1[e.RowIndex, "chkTorikeshi"].Value.ToString() == "True")
            {
                return;
            }

            //if (e.CellName == "btnCell")
            //{
            //    int sMID = Utility.StrtoInt(gcMultiRow1[e.RowIndex, "txtID"].Value.ToString());

            //    if (dts.過去勤務票明細.Any(a => a.ID == sMID))
            //    {
            //        var s = dts.過去勤務票明細.Single(a => a.ID == sMID);
            //        string kID = s.帰宅後勤務ID;

            //        if (getKitakuData(kID))
            //        {
            //            frmKitakugo frm = new frmKitakugo(_dbName, sMID, kID, hArray, bs, false);
            //            frm.ShowDialog();

            //            // 帰宅後勤務データ再読み込み
            //            tAdp.Fill(dts.帰宅後勤務);
            //        }
            //        else
            //        {
            //            MessageBox.Show("帰宅後勤務データはありません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //        }
            //    }
            //}
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     帰宅後勤務データ存在するか？ </summary>
        /// <param name="sID">
        ///     過去勤務票明細ID </param>
        /// <returns>
        ///     true:あり、false:なし</returns>
        ///------------------------------------------------------------------
        private bool getKitakuData(string sID)
        {
            if (!dts.帰宅後勤務.Any(a => a.勤務票帰宅後ID == sID))
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            Image img;

            img = Image.FromFile(_img);
            e.Graphics.DrawImage(img, 0, 0);
            e.HasMorePages = false;
        }

        private void btnPrn_Click(object sender, EventArgs e)
        {
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

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // フォームを閉じる
            this.Tag = END_BUTTON;
            this.Close();
        }
    }
}
