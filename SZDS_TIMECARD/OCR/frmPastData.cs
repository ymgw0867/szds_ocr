using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
//using System.Data.SqlClient;
using System.Data.OleDb;
using SZDS_TIMECARD.Common;

namespace SZDS_TIMECARD.OCR
{
    public partial class frmPastData : Form
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
        public frmPastData(string dbName, string comName, string sID)
        {
            InitializeComponent();

            _dbName = dbName;       // データベース名
            _comName = comName;     // 会社名
            dID = sID;              // 処理モード

            hAdp.Fill(dts.過去勤務票ヘッダ);
            iAdp.Fill(dts.過去勤務票明細);
        }

        // データアダプターオブジェクト
        DataSet1TableAdapters.TableAdapterManager adpMn = new DataSet1TableAdapters.TableAdapterManager();
        DataSet1TableAdapters.過去勤務票ヘッダTableAdapter hAdp = new DataSet1TableAdapters.過去勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.過去勤務票明細TableAdapter iAdp = new DataSet1TableAdapters.過去勤務票明細TableAdapter();

        // データセットオブジェクト
        DataSet1 dts = new DataSet1();

        /// <summary>
        ///     カレントデータRowsインデックス</summary>
        int cI = 0;

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

        string _dbName = string.Empty;              // 会社領域データベース識別番号
        string _comNo = string.Empty;               // 会社番号
        string _comName = string.Empty;             // 会社名

        // dataGridView1_CellEnterステータス
        bool gridViewCellEnterStatus = true;

        global gl = new global();

        private void frmCorrect_Load(object sender, EventArgs e)
        {
            this.pictureBox1.Image = new Bitmap(pictureBox1.Width, pictureBox1.Height);

            // フォーム最大値
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小値
            Utility.WindowsMinSize(this, this.Width, this.Height);

            //元号を取得
            lblGengou.Text = Properties.Settings.Default.gengou;

            // キャプション
            this.Text = "過去勤怠データＩ／Ｐ票表示";

            // グリッドビュー定義
            GridviewSet gs = new GridviewSet();
            gs.Setting_Shain(dGV);

            // データ表示
            showOcrData();
            
            // tagを初期化
            this.Tag = string.Empty;
        }

        #region データグリッドビューカラム定義
        private static string cCheck = "col1";      // 取消
        private static string cShainNum = "col2";   // 社員番号
        private static string cName = "col3";       // 氏名
        private static string cKinmu = "col4";      // 勤務記号
        private static string cZR1 = "col5";        // 残業理由１
        private static string cZan1 = "col6";       // 残業１
        private static string cZR2 = "col7";        // 残業理由２
        private static string cZan2 = "col8";       // 残業２
        private static string cSH = "col11";        // 開始時
        private static string cEH = "col14";        // 終了時
        private static string cJiyu1 = "col15";     // 事由１
        private static string cJiyu2 = "col15";     // 事由２
        private static string cJiyu3 = "col15";     // 事由３
        private static string cOuenChk = "col16";   // 応援あり
        private static string cSftCode = "col16";   // 変更シフトコード
        private static string cSftChk = "col16";    // シフト通りチェック
        private static string cTorikeshi = "col16"; // 取消
        private static string cID = "colID";        // ID
        private static string cSzCode = "colSzCode";  // 所属コード
        private static string cSzName = "colSzName";  // 所属名

        #endregion

        ///----------------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビュークラス </summary>
        ///----------------------------------------------------------------------------
        private class GridviewSet
        {
            ///----------------------------------------------------------------------------
            /// <summary>
            ///     社員用データグリッドビューの定義を行います</summary> 
            /// <param name="gv">
            ///     データグリッドビューオブジェクト</param>
            ///----------------------------------------------------------------------------
            public void Setting_Shain(DataGridView gv)
            {
                try
                {
                    // データグリッドビューの基本設定
                    setGridView_Properties(gv);

                    // カラムコレクションを空にします
                    gv.Columns.Clear();

                    // 行数をクリア            
                    gv.Rows.Clear();

                    //各列幅指定
                    DataGridViewCheckBoxColumn column = new DataGridViewCheckBoxColumn();
                    gv.Columns.Add(column);
                    gv.Columns[0].Name = cCheck;
                    gv.Columns[0].HeaderText = "取消";

                    gv.Columns.Add(cShainNum, "社員番号");
                    gv.Columns.Add(cName, "氏名");
                    gv.Columns.Add(cKinmu, "記号");
                    gv.Columns.Add(cSH, "出勤");
                    gv.Columns.Add(cEH, "退勤");
                    gv.Columns.Add(cZH, "普");
                    gv.Columns.Add(cZE, "");
                    gv.Columns.Add(cZM, "通");
                    gv.Columns.Add(cSIH, "深");
                    gv.Columns.Add(cSIE, "");
                    gv.Columns.Add(cSIM, "夜");

                    gv.Columns.Add(cID, "");        // 明細ID
                    gv.Columns.Add(cSzCode, "");    // 所属コード
                    gv.Columns.Add(cSzName, "");    // 所属名
                    gv.Columns[cID].Visible = false;
                    gv.Columns[cSzCode].Visible = false;
                    gv.Columns[cSzName].Visible = false;

                    foreach (DataGridViewColumn c in gv.Columns)
                    {
                        // 幅
                        if (c.Name == cShainNum )
                        {
                            c.Width = 70;
                        }
                        else if (c.Name == cName)
                        {
                            c.Width = 157;
                        }
                        else if (c.Name == cKinmu)
                        {
                            c.Width = 40;
                        }
                        else if (c.Name == cSE || c.Name == cEE || c.Name == cZE || c.Name == cSIE)
                        {
                            c.Width = 10;
                        }
                        else
                        {
                            c.Width = 30;
                        }
                                                
                        // 表示位置
                        if (c.Index < 2 || c.Name == cKinmu || c.Name == cSE || c.Name == cEE || c.Name == cZE || c.Name == cSIE)
                        {
                            c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        }
                        else if (c.Name == cName)
                        {
                            c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                        }
                        else
                        {
                            c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                        }

                        if (c.Name == cSH || c.Name == cEH || c.Name == cZH || c.Name == cSIH) 
                            c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

                        if (c.Name == cSM || c.Name == cEM || c.Name == cZM || c.Name == cSIM) 
                            c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;

                        // 編集可否
                        //if (c.Name == cName || c.Name == cSE || c.Name == cEE || c.Name == cZE || c.Name == cSIE)
                        //    c.ReadOnly = true;
                        //else c.ReadOnly = false;
                        
                        c.ReadOnly = true;

                        // 区切り文字
                        if (c.Name == cSE || c.Name == cEE || c.Name == cZE || c.Name == cSIE)
                            c.DefaultCellStyle.Font = new Font("ＭＳＰゴシック", 8, FontStyle.Regular);

                        // 入力可能桁数
                        if (c.Name != cCheck)
                        {
                            DataGridViewTextBoxColumn col = (DataGridViewTextBoxColumn)c;

                            if (c.Name == cSIH)
                            {
                                col.MaxInputLength = 1;
                            }
                            else if (c.Name == cShainNum)
                            {
                                col.MaxInputLength = 5;
                            }
                            else
                            {
                                col.MaxInputLength = 2;
                            }
                        }

                        // ソート禁止
                        c.SortMode = DataGridViewColumnSortMode.NotSortable;
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            ///----------------------------------------------------------------------------
            /// <summary>
            ///     データグリッドビュー基本設定</summary>
            /// <param name="gv">
            ///     データグリッドビューオブジェクト</param>
            ///----------------------------------------------------------------------------
            private void setGridView_Properties(DataGridView gv)
            {
                // 列スタイルを変更する
                gv.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                gv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                gv.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // データフォント指定
                gv.DefaultCellStyle.Font = new Font("Meiryo UI", (Single)11, FontStyle.Regular);

                // 行の高さ
                gv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                gv.ColumnHeadersHeight = 20;
                gv.RowTemplate.Height = 24;

                // 全体の高さ
                //gv.Height = 362;
                gv.Height = 430;

                // 全体の幅
                gv.Width = 580;

                // 奇数行の色
                //gv.AlternatingRowsDefaultCellStyle.BackColor = Color.LightBlue;

                //テキストカラーの設定
                gv.RowsDefaultCellStyle.ForeColor = Color.Navy;       
                gv.DefaultCellStyle.SelectionBackColor = Color.Empty;
                gv.DefaultCellStyle.SelectionForeColor = Color.Navy;

                // 行ヘッダを表示しない
                gv.RowHeadersVisible = false;

                // 選択モード
                gv.SelectionMode = DataGridViewSelectionMode.CellSelect;
                gv.MultiSelect = false;

                // データグリッドビュー編集不可
                gv.ReadOnly = true;

                // 追加行表示しない
                gv.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                gv.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                gv.AllowUserToOrderColumns = false;

                // 列サイズ変更不可
                gv.AllowUserToResizeColumns = false;

                // 行サイズ変更禁止
                gv.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //gv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                //TAB動作
                gv.StandardTab = false;

                // 編集モード
                //gv.EditMode = DataGridViewEditMode.EditOnEnter;
            }
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
            if (e.Control is DataGridViewTextBoxEditingControl)
            {
                // 数字のみ入力可能とする
                if (dGV.CurrentCell.ColumnIndex != 0 && dGV.CurrentCell.ColumnIndex != 2)
                {
                    //イベントハンドラが複数回追加されてしまうので最初に削除する
                    e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                    e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress2);

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

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (!gl.ChangeValueStatus) return;

            if (e.RowIndex < 0) return;

            string colName = dGV.Columns[e.ColumnIndex].Name;

            //// 過去データ表示のときは終了
            //if (dID != string.Empty) return;

            // 社員番号のとき社員名を表示します
            if (colName == cShainNum)
            {
                // ChangeValueイベントを発生させない
                gl.ChangeValueStatus = false;

                // 氏名を初期化
                dGV[cName, e.RowIndex].Value = string.Empty;

                // 奉行データベースより社員名を取得して表示します
                if (Utility.NulltoStr(dGV[cShainNum, e.RowIndex].Value) != string.Empty)
                {
                    dbControl.DataControl dCon = new dbControl.DataControl(_dbName);
                    string sYY = (Utility.StrtoInt(txtYear.Text) + Properties.Settings.Default.rekiHosei).ToString();
                    string sMM = Utility.StrtoInt(txtMonth.Text).ToString();
                    string sDD = Utility.StrtoInt(txtDay.Text).ToString();
                    string sNum = Utility.NulltoStr(dGV[cShainNum, e.RowIndex].Value);
                    OleDbDataReader dR = dCon.GetEmployeeBase(sYY, sMM, sDD, sNum);
                    while (dR.Read())
                    {
                        // 所属名・社員名表示
                        lblShozoku.Text = dR["DepartmentName"].ToString().Trim();
                        dGV[cName, e.RowIndex].Value = dR["Name"].ToString().Trim();
                        dGV[cSzCode, e.RowIndex].Value = dR["DepartmentCode"].ToString().Trim().Substring(10, 5);
                        dGV[cSzName, e.RowIndex].Value = dR["DepartmentName"].ToString().Trim();
                    }

                    dR.Close();
                    dCon.Close();

                    // 時刻区切り文字
                    dGV[cSE, e.RowIndex].Value = ":";
                    dGV[cEE, e.RowIndex].Value = ":";
                    dGV[cSIE, e.RowIndex].Value = ":";
                    dGV[cZE, e.RowIndex].Value = ":";
                }

                // ChangeValueイベントステータスをtrueに戻す
                gl.ChangeValueStatus = true;
            }
        }

        private void frmCorrect_Shown(object sender, EventArgs e)
        {
            if (dID != string.Empty) btnRtn.Focus();
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
            // フォームを閉じる
            this.Tag = END_BUTTON;
            this.Close();
        }

        private void frmCorrect_FormClosing(object sender, FormClosingEventArgs e)
        {
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
        }

        private void btnMinus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor > gl.ZOOM_MIN)
            {
                leadImg.ScaleFactor -= gl.ZOOM_STEP;
            }
            gl.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }

        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            //if (e.RowIndex < 0) return;

            string colName = dGV.Columns[e.ColumnIndex].Name;

            if (colName == cSH || colName == cSE || colName == cEH || colName == cEE ||
                colName == cZH || colName == cZE || colName == cSIH || colName == cSIE)
            {
                e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;
            }
        }

        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            string colName = dGV.Columns[dGV.CurrentCell.ColumnIndex].Name;
            //if (colName == cKyuka || colName == cCheck)
            //{
            //    if (dGV.IsCurrentCellDirty)
            //    {
            //        dGV.CommitEdit(DataGridViewDataErrorContexts.Commit);
            //        dGV.RefreshEdit();
            //    }
            //}
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
        }

        private void dataGridView1_CellEnter_1(object sender, DataGridViewCellEventArgs e)
        {
            // エラー表示時には処理を行わない
            if (!gridViewCellEnterStatus) return;
 
            string ColH = string.Empty;
            string ColM = dGV.Columns[dGV.CurrentCell.ColumnIndex].Name;

            // 開始時間または終了時間を判断
            if (ColM == cSM)            // 開始時刻
            {
                ColH = cSH;
            }
            else if (ColM == cEM)       // 終了時刻
            {
                ColH = cEH;
            }
            else if (ColM == cZM)       // 時間外
            {
                ColH = cZH;
            }
            else if (ColM == cSIM)      // 深夜
            {
                ColH = cSIH;
            }
            else
            {
                return;
            }

            // 時が入力済みで分が未入力のとき分に"00"を表示します
            if (dGV[ColH, dGV.CurrentRow.Index].Value != null)
            {
                if (dGV[ColH, dGV.CurrentRow.Index].Value.ToString().Trim() != string.Empty)
                {
                    if (dGV[ColM, dGV.CurrentRow.Index].Value == null)
                    {
                        dGV[ColM, dGV.CurrentRow.Index].Value = "00";
                    }
                    else if (dGV[ColM, dGV.CurrentRow.Index].Value.ToString().Trim() == string.Empty)
                    {
                        dGV[ColM, dGV.CurrentRow.Index].Value = "00";
                    }
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

        private void maskedTextBox3_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void txtYear_TextChanged(object sender, EventArgs e)
        {
            // 曜日
            DateTime eDate;
            int tYY = Utility.StrtoInt(txtYear.Text) + Properties.Settings.Default.rekiHosei;
            string sDate = tYY.ToString() + "/" + Utility.EmptytoZero(txtMonth.Text) + "/" +
                    Utility.EmptytoZero(txtDay.Text);

            // 存在する日付と認識された場合、曜日を表示する
            if (DateTime.TryParse(sDate, out eDate))
            {
                txtWeekDay.Text = ("日月火水木金土").Substring(int.Parse(eDate.DayOfWeek.ToString("d")), 1);
            }
            else
            {
                txtWeekDay.Text = string.Empty;
            }
        }
    }
}
