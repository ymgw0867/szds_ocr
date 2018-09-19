using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using SZDS_TIMECARD.Common;
using GrapeCity.Win.MultiRow;

namespace SZDS_TIMECARD.config
{
    public partial class frmKitakugoWork : Form
    {
        public frmKitakugoWork(string dbName)
        {
            InitializeComponent();

            _dbName = dbName;
            kAdp.Fill(dts.帰宅後勤務);
        }

        string _dbName = string.Empty;
        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.帰宅後勤務TableAdapter kAdp = new DataSet1TableAdapters.帰宅後勤務TableAdapter();

        bool changeValueStatus = true;
        const int MODE_ADD = 0;     // 追加モード
        const int MODE_EDIT = 1;    // 更新モード
        int fMode = 0;
        int fID = 0;
        
        private void frmKitakugoWork_Load(object sender, EventArgs e)
        {
            Utility.WindowsMaxSize(this, this.Width, this.Height);
            Utility.WindowsMinSize(this, this.Width, this.Height);
            
            // Tabキーの既定のショートカットキーを解除する。
            gcMultiRow1.ShortcutKeyManager.Unregister(Keys.Tab);
            gcMultiRow1.ShortcutKeyManager.Unregister(Keys.Enter);

            // Tabキーのショートカットキーにユーザー定義のショートカットキーを割り当てる。
            gcMultiRow1.ShortcutKeyManager.Register(new clsKeyTab.CustomMoveToNextContorol(), Keys.Tab);
            gcMultiRow1.ShortcutKeyManager.Register(new clsKeyTab.CustomMoveToNextContorol(), Keys.Enter);
            
            GridViewSetting(dataGridView1);
            gcMrSetting();

            GridViewShow(dataGridView1);
            dispClear();
            gcMultiRow1.CurrentCell = null;
        }
        
        // グリッドビューカラム名
        private string cDate = "c1";
        private string cSNum = "c2";
        private string cName = "c3";
        private string cShukkin = "c4";
        private string cTaikin = "c5";
        private string cZanRe1 = "c6";
        private string cZanGyo1 = "c7";
        private string cZanRe2 = "c8";
        private string cZanGyo2 = "c9";
        private string cSft = "c10";
        private string cJiyu = "c11";
        private string cID = "c12";
        private string cUpDate = "c13";
        
        ///------------------------------------------------------------------------
        /// <summary>
        ///     グリッドビューの定義を行います  </summary>
        /// <param name="dg">
        ///     データグリッドビューオブジェクト</param>
        ///------------------------------------------------------------------------
        private void GridViewSetting(DataGridView dg)
        {
            try
            {
                dg.EnableHeadersVisualStyles = false;
                dg.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                dg.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                //フォームサイズ定義

                // 列スタイルを変更する

                dg.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                dg.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                dg.ColumnHeadersDefaultCellStyle.Font = new Font("メイリオ", 9, FontStyle.Regular);

                // データフォント指定
                dg.DefaultCellStyle.Font = new Font("メイリオ", 9, FontStyle.Regular);

                // 行の高さ
                dg.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                dg.ColumnHeadersHeight = 20;
                dg.RowTemplate.Height = 20;

                // 全体の高さ
                dg.Height = 180;

                // 全体の幅
                //dg.Width = 583;

                // 奇数行の色
                dg.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                //各列幅指定
                dg.Columns.Add(cDate, "年月日");
                dg.Columns.Add(cSNum, "社員番号");
                dg.Columns.Add(cName, "氏名");
                dg.Columns.Add(cShukkin, "出勤");
                dg.Columns.Add(cTaikin, "退勤");
                dg.Columns.Add(cZanRe1, "理由");
                dg.Columns.Add(cZanGyo1, "残業");
                dg.Columns.Add(cZanRe2, "理由");
                dg.Columns.Add(cZanGyo2, "残業");
                dg.Columns.Add(cSft, "シフト");
                dg.Columns.Add(cJiyu, "事由");
                dg.Columns.Add(cUpDate, "更新年月日");
                dg.Columns.Add(cID, "ID");

                dg.Columns[cID].Visible = false;

                dg.Columns[cDate].Width = 110;
                dg.Columns[cSNum].Width = 90;
                dg.Columns[cName].Width = 110;
                dg.Columns[cShukkin].Width = 80;
                dg.Columns[cTaikin].Width = 80;
                dg.Columns[cZanRe1].Width = 80;
                dg.Columns[cZanGyo1].Width = 80;
                dg.Columns[cZanRe2].Width = 80;
                dg.Columns[cZanGyo2].Width = 80;
                dg.Columns[cJiyu].Width = 80;
                dg.Columns[cSft].Width = 80;
                dg.Columns[cUpDate].Width = 160;

                //dg.Columns[cMemo].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dg.Columns[cDate].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dg.Columns[cSNum].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dg.Columns[cShukkin].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dg.Columns[cTaikin].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dg.Columns[cZanRe1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dg.Columns[cZanGyo1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dg.Columns[cZanRe2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dg.Columns[cZanGyo2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dg.Columns[cJiyu].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dg.Columns[cSft].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dg.Columns[cUpDate].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 行ヘッダを表示しない
                dg.RowHeadersVisible = false;

                // 選択モード
                dg.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                dg.MultiSelect = false;

                // 編集不可とする
                dg.ReadOnly = true;

                // 追加行表示しない
                dg.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                dg.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                dg.AllowUserToOrderColumns = false;

                // 列サイズ変更不可
                dg.AllowUserToResizeColumns = false;

                // 行サイズ変更禁止
                dg.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //dg.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                //TAB動作
                dg.StandardTab = false;

                //// ソート禁止
                //foreach (DataGridViewColumn c in dg.Columns)
                //{
                //    c.SortMode = DataGridViewColumnSortMode.NotSortable;
                //}
                //dg.Columns[cDay].SortMode = DataGridViewColumnSortMode.NotSortable;

                // 罫線
                dg.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                dg.CellBorderStyle = DataGridViewCellBorderStyle.None;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
            }
        }

        private void gcMrSetting()
        {
            //MultiRow編集モード
            gcMultiRow1.EditMode = EditMode.EditProgrammatically;

            this.gcMultiRow1.AllowUserToAddRows = false;                    // 手動による行追加を禁止する
            this.gcMultiRow1.AllowUserToDeleteRows = false;                 // 手動による行削除を禁止する
            this.gcMultiRow1.Rows.Clear();                                  // 行数をクリア
            this.gcMultiRow1.RowCount = 1;                                  // 行数を設定
            this.gcMultiRow1.ReadOnly = false;
            this.gcMultiRow1.HideSelection = true;                          // GcMultiRow コントロールがフォーカスを失ったとき、セルの選択状態を非表示にする
        }

        ///------------------------------------------------------------------------
        /// <summary>
        ///     休日データをグリッドビューへ表示します </summary>
        /// <param name="tempGrid">
        ///     データグリッドビューオブジェクト</param>
        ///------------------------------------------------------------------------
        private void GridViewShow(DataGridView g)
        {
            string sc = sqlControl.obcConnectSting.get(_dbName);
            sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            int iX = 0;
            g.RowCount = 0;

            foreach (var t in dts.帰宅後勤務.OrderByDescending(a => a.年).ThenByDescending(a => a.月).ThenByDescending(a => a.日).ThenBy(a => a.ID))
            {
                g.Rows.Add();

                g[cDate, iX].Value = t.年.ToString().PadLeft(2, '0') + "/" + t.月.ToString().PadLeft(2, '0') + "/" + t.日.ToString().PadLeft(2, '0');
                g[cSNum, iX].Value = t.社員番号.ToString().PadLeft(6, '0');

                SqlDataReader dR = sdCon.free_dsReader(Utility.getEmployee(t.社員番号.ToString().PadLeft(10, '0')));
                g[cName, iX].Value = string.Empty;

                while (dR.Read())
                {
                    g[cName, iX].Value = dR["Name"].ToString().Trim();
                }

                dR.Close();

                g[cShukkin, iX].Value = t.出勤時.ToString().PadLeft(2, '0') + ":" + t.出勤分.ToString().PadLeft(2, '0');
                g[cTaikin, iX].Value = t.退勤時.ToString().PadLeft(2, '0') + ":" + t.退勤分.ToString().PadLeft(2, '0');
                g[cZanRe1, iX].Value = t.残業理由1.ToString();
                g[cZanGyo1, iX].Value = t.残業時1.ToString() + "." + t.残業分1.ToString();
                g[cZanRe2, iX].Value = t.残業理由2.ToString();
                g[cZanGyo2, iX].Value = t.残業時2.ToString() + "." + t.残業分2.ToString();

                if (t.IsシフトコードNull())
                {
                    g[cSft, iX].Value = "";
                }
                else
                {
                    g[cSft, iX].Value = t.シフトコード.ToString();
                }

                g[cJiyu, iX].Value = t.事由1.ToString();
                g[cUpDate, iX].Value = t.更新年月日;
                g[cID, iX].Value = t.ID.ToString();

                iX++;
            }

            g.CurrentCell = null;

            sdCon.Close();
        }

        private void gcMultiRow1_CellEnter(object sender, CellEventArgs e)
        {
            if (gcMultiRow1.EditMode == EditMode.EditProgrammatically)
            {
                gcMultiRow1.BeginEdit(true);
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

        private void gcMultiRow1_EditingControlShowing(object sender, EditingControlShowingEventArgs e)
        {
            if (e.Control is TextBoxEditingControl)
            {
                //イベントハンドラが複数回追加されてしまうので最初に削除する
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress2);
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress3);

                // 数字のみ入力可能とする
                if (gcMultiRow1.CurrentCell.Name == "txtShainNum" ||
                    gcMultiRow1.CurrentCell.Name == "txtYear" || gcMultiRow1.CurrentCell.Name == "txtMonth" ||
                    gcMultiRow1.CurrentCell.Name == "txtDay" || 
                    gcMultiRow1.CurrentCell.Name == "txtSh" || gcMultiRow1.CurrentCell.Name == "txtSm" ||
                    gcMultiRow1.CurrentCell.Name == "txtEh" || gcMultiRow1.CurrentCell.Name == "txtEm" ||
                    gcMultiRow1.CurrentCell.Name == "txtZanRe1" || gcMultiRow1.CurrentCell.Name == "txtZanH1" ||
                    gcMultiRow1.CurrentCell.Name == "txtZanRe2" || gcMultiRow1.CurrentCell.Name == "txtZanH2" ||
                    gcMultiRow1.CurrentCell.Name == "txtJiyu")
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

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 閉じる
            Close();
        }

        private void frmKitakugoWork_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            Dispose();
        }
        
        private void gcMultiRow1_CellValueChanged(object sender, CellEventArgs e)
        {
            if (!changeValueStatus)
            {
                return;
            }

            if (e.CellName == "cmbSft")
            {
                ComboBoxCell comboBoxCell = gcMultiRow1[0, "cmbSft"] as ComboBoxCell;
                object selectedValue = comboBoxCell.Value;
                int selectedIndex = -1;
                if (selectedValue != null)
                {
                    selectedIndex = comboBoxCell.Items.IndexOf(selectedValue);
                }

                if (selectedIndex == 0)
                {
                    gcMultiRow1[0, "lblSftName"].Value = "単独残業"; 
                }
                else if (selectedIndex == 1)
                {
                    gcMultiRow1[0, "lblSftName"].Value = "単独休出";
                }
                else
                {
                    gcMultiRow1[0, "lblSftName"].Value = "";
                }
            }


            // 奉行データベースより社員名を取得して表示します
            if (Utility.NulltoStr(gcMultiRow1[e.RowIndex, "txtShainNum"].Value) != string.Empty)
            {
                // 接続文字列取得
                string sc = sqlControl.obcConnectSting.get(_dbName);
                sqlControl.DataControl sdCon = new Common.sqlControl.DataControl(sc);

                string bCode = gcMultiRow1[e.RowIndex, "txtShainNum"].Value.ToString().PadLeft(10, '0');
                SqlDataReader dR = sdCon.free_dsReader(Utility.getEmployee(bCode));

                changeValueStatus = false;

                gcMultiRow1[e.RowIndex, "lblName"].Value = "該当者なし";
                gcMultiRow1[e.RowIndex, "lblName"].Style.ForeColor = Color.Red;

                while (dR.Read())
                {
                    // 社員名表示
                    gcMultiRow1[e.RowIndex, "lblName"].Value = dR["Name"].ToString().Trim();
                    gcMultiRow1[e.RowIndex, "lblName"].Style.ForeColor = Color.Black;
                }

                dR.Close();
                sdCon.Close();

                changeValueStatus = true;
            }

            if (e.CellName == "txtYear" || e.CellName == "txtMonth" || e.CellName == "txtDay" ||
                e.CellName == "txtSh" || e.CellName == "txtSm" || e.CellName == "txtEh" || e.CellName == "txtEm" || 
                e.CellName == "txtJiyu")
            {
                changeValueStatus = false;

                int val = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[e.RowIndex, e.CellName].Value));

                if (e.CellName == "txtYear")
                {
                    if (val < 2017)
                    {
                        gcMultiRow1[e.RowIndex, e.CellName].Value = "";
                    }
                }

                if (e.CellName == "txtMonth")
                {
                    if (val < 1 || val > 12)
                    {
                        gcMultiRow1[e.RowIndex, e.CellName].Value = "";
                    }
                }

                if (e.CellName == "txtDay")
                {
                    if (val < 1 || val > 31)
                    {
                        gcMultiRow1[e.RowIndex, e.CellName].Value = "";
                    }
                }

                if (e.CellName == "txtSh" || e.CellName == "txtEh")
                {
                    if (val > 23)
                    {
                        gcMultiRow1[e.RowIndex, e.CellName].Value = "";
                    }
                }

                if (e.CellName == "txtSm" || e.CellName == "txtEm")
                {
                    if (val > 59)
                    {
                        gcMultiRow1[e.RowIndex, e.CellName].Value = "";
                    }
                }

                if (e.CellName == "txtJiyu")
                {
                    if (val != 30)
                    {
                        gcMultiRow1[e.RowIndex, e.CellName].Value = "";
                    }
                }

                changeValueStatus = true;
            }
        }

        ///-----------------------------------------------------------
        /// <summary>
        ///     画面初期化 </summary>
        ///-----------------------------------------------------------
        private void dispClear()
        {
            gcMultiRow1[0, "txtShainNum"].Value = string.Empty;
            gcMultiRow1[0, "lblName"].Value = string.Empty;
            gcMultiRow1[0, "txtYear"].Value = string.Empty;
            gcMultiRow1[0, "txtMonth"].Value = string.Empty;
            gcMultiRow1[0, "txtDay"].Value = string.Empty;
            gcMultiRow1[0, "txtSh"].Value = string.Empty;
            gcMultiRow1[0, "txtSm"].Value = string.Empty;
            gcMultiRow1[0, "txtEh"].Value = string.Empty;
            gcMultiRow1[0, "txtEm"].Value = string.Empty;
            gcMultiRow1[0, "txtZanRe1"].Value = string.Empty;
            gcMultiRow1[0, "txtZanH1"].Value = string.Empty;
            gcMultiRow1[0, "txtZanM1"].Value = string.Empty;
            gcMultiRow1[0, "txtZanRe2"].Value = string.Empty;
            gcMultiRow1[0, "txtZanH2"].Value = string.Empty;
            gcMultiRow1[0, "txtZanM2"].Value = string.Empty;
            gcMultiRow1[0, "txtJiyu"].Value = string.Empty;
            gcMultiRow1[0, "lblSftName"].Value = string.Empty;
            gcMultiRow1[0, "cmbSft"].Value = null;
            gcMultiRow1[0, "cmbSh"].Value = null;   // 2018/03/08
            gcMultiRow1[0, "cmbEh"].Value = null;   // 2018/03/08

            fMode = MODE_ADD;
            fID = 0;

            lnkLblDelete.Enabled = false;
            lnkLblClr.Enabled = false;
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     エラーチェック </summary>
        /// <param name="g">
        ///     GcMultiRowオブジェクト</param>
        /// <returns>
        ///     true:エラーなし、false:エラーあり</returns>
        ///-------------------------------------------------------------------
        private bool errCheck(GcMultiRow g)
        {
            bool rtn = true;

            // 社員番号未入力
            if (Utility.NulltoStr(gcMultiRow1[0, "txtShainNum"].Value) == string.Empty)
            {
                MessageBox.Show("社員番号が未入力です", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtShainNum"];
                return false;
            }

            // 社員番号、奉行データベース登録確認
            if (Utility.NulltoStr(gcMultiRow1[0, "txtShainNum"].Value) != string.Empty)
            {
                // 接続文字列取得
                string sc = sqlControl.obcConnectSting.get(_dbName);
                sqlControl.DataControl sdCon = new Common.sqlControl.DataControl(sc);

                string bCode = gcMultiRow1[0, "txtShainNum"].Value.ToString().PadLeft(10, '0');
                SqlDataReader dR = sdCon.free_dsReader(Utility.getEmployee(bCode));

                if (!dR.HasRows)
                {
                    rtn = false;
                }

                dR.Close();
                sdCon.Close();

                if (!rtn)
                {
                    MessageBox.Show("マスター未登録の社員番号です", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtShainNum"];
                    return false;
                }
            }
            
            int yy = Utility.StrtoInt(Utility.NulltoStr(g[0, "txtYear"].Value));
            int mm = Utility.StrtoInt(Utility.NulltoStr(g[0, "txtMonth"].Value));
            int dd = Utility.StrtoInt(Utility.NulltoStr(g[0, "txtDay"].Value));
            
            // 日付
            DateTime dt;
            if (!DateTime.TryParse((yy + "/" + mm + "/" + dd), out dt))
            {
                MessageBox.Show("日付が正しくありません","確認",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtYear"];
                return false;
            }
            
            // 勤務時間
            if (Utility.NulltoStr(gcMultiRow1[0, "txtSh"].Value) == string.Empty)
            {
                MessageBox.Show("出勤時刻が未入力です", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtSh"];
                return false;
            }
            
            if (Utility.NulltoStr(gcMultiRow1[0, "txtSm"].Value) == string.Empty)
            {
                MessageBox.Show("出勤時刻が未入力です", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtSm"];
                return false;
            }
            
            if (Utility.NulltoStr(gcMultiRow1[0, "txtEh"].Value) == string.Empty)
            {
                MessageBox.Show("退勤時刻が未入力です", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtEh"];
                return false;
            }
            
            if (Utility.NulltoStr(gcMultiRow1[0, "txtEm"].Value) == string.Empty)
            {
                MessageBox.Show("退勤時刻が未入力です", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtEm"];
                return false;
            }

            // 出勤時刻
            DateTime sDt;
            if (!DateTime.TryParse(gcMultiRow1[0, "txtSh"].Value.ToString() + ":" + gcMultiRow1[0, "txtSm"].Value.ToString() + ":00", out sDt))
            {
                MessageBox.Show("出勤時刻が正しくありません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtSh"];
                return false;
            }

            // 退勤時刻
            DateTime eDt;
            if (!DateTime.TryParse(gcMultiRow1[0, "txtEh"].Value.ToString() + ":" + gcMultiRow1[0, "txtEm"].Value.ToString() + ":00", out eDt))
            {
                MessageBox.Show("退勤時刻が正しくありません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtEh"];
                return false;
            }

            // 残業
            if (Utility.NulltoStr(gcMultiRow1[0, "txtZanRe1"].Value) != string.Empty && (
                Utility.NulltoStr(gcMultiRow1[0, "txtZanH1"].Value) == string.Empty &&
                Utility.NulltoStr(gcMultiRow1[0, "txtZanM1"].Value) == string.Empty))
            {
                MessageBox.Show("残業時間が未入力です", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtZanH1"];
                return false;
            }

            if (Utility.NulltoStr(gcMultiRow1[0, "txtZanRe1"].Value) == string.Empty && (
                Utility.NulltoStr(gcMultiRow1[0, "txtZanH1"].Value) != string.Empty ||
                Utility.NulltoStr(gcMultiRow1[0, "txtZanM1"].Value) != string.Empty))
            {
                MessageBox.Show("残業理由が未入力です", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtZanRe1"];
                return false;
            }

            if (Utility.NulltoStr(gcMultiRow1[0, "txtZanH1"].Value) != string.Empty &&
                Utility.NulltoStr(gcMultiRow1[0, "txtZanM1"].Value) == string.Empty)
            {
                MessageBox.Show("残業時間が未入力です", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtZanM1"];
                return false;
            }

            if (Utility.NulltoStr(gcMultiRow1[0, "txtZanH1"].Value) == string.Empty &&
                Utility.NulltoStr(gcMultiRow1[0, "txtZanM1"].Value) != string.Empty)
            {
                MessageBox.Show("残業時間が未入力です", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtZanH1"];
                return false;
            }


            if (Utility.NulltoStr(gcMultiRow1[0, "txtZanRe2"].Value) != string.Empty && (
                Utility.NulltoStr(gcMultiRow1[0, "txtZanH2"].Value) == string.Empty &&
                Utility.NulltoStr(gcMultiRow1[0, "txtZanM2"].Value) == string.Empty))
            {
                MessageBox.Show("残業時間が未入力です", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtZanH2"];
                return false;
            }

            if (Utility.NulltoStr(gcMultiRow1[0, "txtZanRe2"].Value) == string.Empty && (
                Utility.NulltoStr(gcMultiRow1[0, "txtZanH2"].Value) != string.Empty ||
                Utility.NulltoStr(gcMultiRow1[0, "txtZanM2"].Value) != string.Empty))
            {
                MessageBox.Show("残業理由が未入力です", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtZanRe2"];
                return false;
            }

            if (Utility.NulltoStr(gcMultiRow1[0, "txtZanH2"].Value) != string.Empty &&
                Utility.NulltoStr(gcMultiRow1[0, "txtZanM2"].Value) == string.Empty)
            {
                MessageBox.Show("残業時間が未入力です", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtZanM2"];
                return false;
            }

            if (Utility.NulltoStr(gcMultiRow1[0, "txtZanH2"].Value) == string.Empty &&
                Utility.NulltoStr(gcMultiRow1[0, "txtZanM2"].Value) != string.Empty)
            {
                MessageBox.Show("残業時間が未入力です", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtZanH2"];
                return false;
            }

            // 実働時間
            double w = Utility.GetTimeSpan(sDt, eDt).TotalMinutes;

            // 残業時間
            double z = (Utility.StrtoDouble(Utility.NulltoStr(gcMultiRow1[0, "txtZanH1"].Value)) * 60) + (Utility.StrtoDouble(Utility.NulltoStr(gcMultiRow1[0, "txtZanM1"].Value)) * 60 / 10) +
                       (Utility.StrtoDouble(Utility.NulltoStr(gcMultiRow1[0, "txtZanH2"].Value)) * 60) + (Utility.StrtoDouble(Utility.NulltoStr(gcMultiRow1[0, "txtZanM2"].Value)) * 60 / 10);

            if (z > w)
            {
                MessageBox.Show("残業時間が実働時間以上です", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtZanH1"];
                return false;
            }

            // シフトコード
            ComboBoxCell comboBoxCell = gcMultiRow1[0, "cmbSft"] as ComboBoxCell;
            object selectedValue = comboBoxCell.Value;
            if (selectedValue == null)
            {
                MessageBox.Show("シフトコードを選択してください", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "cmbSft"];
                return false;
            }

            // 事由
            string sJiyu = Utility.NulltoStr(gcMultiRow1[0, "txtJiyu"].Value);
            if (sJiyu != "30" && sJiyu != string.Empty)
            {
                MessageBox.Show("帰宅後勤務で使用可能な事由は「呼出回数」のみです", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtJiyu"];
                return false;
            }

            // 当日・翌日
            comboBoxCell = gcMultiRow1[0, "cmbSh"] as ComboBoxCell;
            selectedValue = comboBoxCell.Value;
            if (selectedValue == null)
            {
                MessageBox.Show("出勤時刻の当日・翌日を選択してください", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "cmbSh"];
                return false;
            }

            comboBoxCell = gcMultiRow1[0, "cmbEh"] as ComboBoxCell;
            selectedValue = comboBoxCell.Value;
            if (selectedValue == null)
            {
                MessageBox.Show("退勤時刻の当日・翌日を選択してください", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "cmbEh"];
                return false;
            }
            return true;
        }

        private void lnkLblUpdate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // エラーチェック
            if (!errCheck(gcMultiRow1))
            {
                return;
            }

            // 残業未入力のとき警告
            if (!zanWarning())
            {
                if (MessageBox.Show("残業が未入力です。このまま登録してよろしいですか", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                {
                    return;
                }
            }

            if (fMode == MODE_ADD)
            {
                // 追加モード
                if (!dataInsert(gcMultiRow1))
                {
                    MessageBox.Show("データ登録に失敗しました");
                    return;
                }
            }
            else if (fMode == MODE_EDIT)
            {
                // 更新モード
                if (!dataUpdate(gcMultiRow1, fID))
                {
                    MessageBox.Show("データ更新に失敗しました");
                    return;
                }
            }

            // データベース更新
            kAdp.Update(dts.帰宅後勤務);
            kAdp.Fill(dts.帰宅後勤務);
            GridViewShow(dataGridView1);
            dispClear();
        }

        ///-----------------------------------------------------------------
        /// <summary>
        ///     残業未入力チェック </summary>
        /// <returns>
        ///     true:入力あり, false:入力なし</returns>
        ///-----------------------------------------------------------------
        private bool zanWarning()
        {
            if (Utility.NulltoStr(gcMultiRow1[0, "txtZanH1"].Value) == string.Empty &&
                Utility.NulltoStr(gcMultiRow1[0, "txtZanM1"].Value) == string.Empty &&
                Utility.NulltoStr(gcMultiRow1[0, "txtZanH2"].Value) == string.Empty &&
                Utility.NulltoStr(gcMultiRow1[0, "txtZanM2"].Value) == string.Empty)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        
        private bool dataInsert(GcMultiRow g)
        {
            try
            {
                DataSet1.帰宅後勤務Row r = dts.帰宅後勤務.New帰宅後勤務Row();
                setRowData(ref r, gcMultiRow1);
                dts.帰宅後勤務.Add帰宅後勤務Row(r);

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        private bool dataUpdate(GcMultiRow g, int sID)
        {
            try
            {
                DataSet1.帰宅後勤務Row r = dts.帰宅後勤務.FindByID(sID);
                setRowData(ref r, gcMultiRow1);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        private void setRowData(ref DataSet1.帰宅後勤務Row r, GcMultiRow g)
        {
            r.勤務票帰宅後ID = "";
            r.年 = Utility.StrtoInt(Utility.NulltoStr(g[0, "txtYear"].Value));
            r.月 = Utility.StrtoInt(Utility.NulltoStr(g[0, "txtMonth"].Value));
            r.日 = Utility.StrtoInt(Utility.NulltoStr(g[0, "txtDay"].Value));
            r.社員番号 = Utility.NulltoStr(g[0, "txtShainNum"].Value);
            r.出勤時 = Utility.NulltoStr(g[0, "txtSh"].Value);
            r.出勤分 = Utility.NulltoStr(g[0, "txtSm"].Value);
            r.退勤時 = Utility.NulltoStr(g[0, "txtEh"].Value);
            r.退勤分 = Utility.NulltoStr(g[0, "txtEm"].Value);
            r.残業理由1 = Utility.NulltoStr(g[0, "txtZanRe1"].Value);
            r.残業時1 = Utility.NulltoStr(g[0, "txtZanH1"].Value);
            r.残業分1 = Utility.NulltoStr(g[0, "txtZanM1"].Value);
            r.残業理由2 = Utility.NulltoStr(g[0, "txtZanRe2"].Value);
            r.残業時2 = Utility.NulltoStr(g[0, "txtZanH2"].Value);
            r.残業分2 = Utility.NulltoStr(g[0, "txtZanM2"].Value);
            r.事由1 = Utility.NulltoStr(g[0, "txtJiyu"].Value);
            r.事由2 = string.Empty;
            r.事由3 = string.Empty;

            ComboBoxCell comboBoxCell = gcMultiRow1[0, "cmbSft"] as ComboBoxCell;
            object selectedValue = comboBoxCell.Value;
            r.シフトコード = selectedValue.ToString();

            // 出勤日：2018/03/08
            comboBoxCell = gcMultiRow1[0, "cmbSh"] as ComboBoxCell;
            selectedValue = comboBoxCell.Value;

            if (selectedValue.ToString() == "当日")
            {
                r.出勤日 = string.Empty;
            }
            else
            {
                // 翌日
                r.出勤日 = selectedValue.ToString();
            }

            // 退勤日：2018/03/08
            comboBoxCell = gcMultiRow1[0, "cmbEh"] as ComboBoxCell;
            selectedValue = comboBoxCell.Value;

            if (selectedValue.ToString() == "当日")
            {
                r.退勤日 = string.Empty;
            }
            else
            {
                // 翌日
                r.退勤日 = selectedValue.ToString();
            }

            r.取消 = string.Empty;
            r.データ領域名 = _dbName;
            r.更新年月日 = DateTime.Now;
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            selectViewData(dataGridView1, gcMultiRow1);
        }

        private void selectViewData(DataGridView dg, GcMultiRow g)
        {
            if (dg.SelectedRows.Count == 0)
            {
                return;
            }

            int r = dg.SelectedRows[0].Index;
            
            fID = Utility.StrtoInt(dg[cID, r].Value.ToString());

            if (dts.帰宅後勤務.Any(a => a.ID == fID))
            {
                var s = dts.帰宅後勤務.Single(a => a.ID == fID);
                g[0, "txtShainNum"].Value = s.社員番号.ToString();
                g[0, "txtYear"].Value = s.年.ToString();
                g[0, "txtMonth"].Value = s.月.ToString();
                g[0, "txtDay"].Value = s.日.ToString();
                g[0, "txtSh"].Value = s.出勤時.ToString();
                g[0, "txtSm"].Value = s.出勤分.ToString().PadLeft(2, '0');
                g[0, "txtEh"].Value = s.退勤時.ToString();
                g[0, "txtEm"].Value = s.退勤分.ToString().PadLeft(2, '0');
                g[0, "txtZanRe1"].Value = s.残業理由1.ToString();
                g[0, "txtZanH1"].Value = s.残業時1.ToString();
                g[0, "txtZanM1"].Value = s.残業分1.ToString();
                g[0, "txtZanRe2"].Value = s.残業理由2.ToString();
                g[0, "txtZanH2"].Value = s.残業時2.ToString();
                g[0, "txtZanM2"].Value = s.残業分2.ToString();
                g[0, "cmbSft"].Value = s.シフトコード.ToString().PadLeft(3, '0');
                g[0, "txtJiyu"].Value = s.事由1.ToString();

                // 2018/03/08
                if (s.Is出勤日Null() || s.出勤日 == string.Empty)
                {
                    g[0, "cmbSh"].Value = "当日";
                }
                else
                {
                    g[0, "cmbSh"].Value = s.出勤日;
                }

                // 2018/03/08
                if (s.Is退勤日Null() || s.退勤日 == string.Empty)
                {
                    g[0, "cmbEh"].Value = "当日";
                }
                else
                {
                    g[0, "cmbEh"].Value = s.退勤日;
                }
            }
            else
            {
                MessageBox.Show("データの取得に失敗しました");
                return;
            }

            lnkLblUpdate.Enabled = true;
            lnkLblDelete.Enabled = true;
            lnkLblClr.Enabled = true;

            fMode = MODE_EDIT;
        }

        private void lnkLblClr_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            dispClear();
        }

        private void gcMultiRow1_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void lnkLblDelete_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("データを削除します。よろしいですか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;

            int sID = int.Parse(dataGridView1[cID, dataGridView1.SelectedRows[0].Index].Value.ToString());
            
            if (dataDelete(sID))
            {            
                MessageBox.Show("削除しました", "結果", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            
            // データベース更新
            kAdp.Update(dts.帰宅後勤務);
            kAdp.Fill(dts.帰宅後勤務);
            GridViewShow(dataGridView1);
            dispClear();
        }

        ///--------------------------------------------------------------
        /// <summary>
        ///     データを削除する </summary>
        /// <param name="sID">
        ///     レコードID</param>
        /// <returns>
        ///     true:削除成功、false:削除失敗</returns>
        ///--------------------------------------------------------------
        private bool dataDelete(int sID)
        {
            try
            {
                // 削除データ取得（エラー回避のためDataRowState.Deleted と DataRowState.Detachedは除外して抽出する）
                var d = dts.帰宅後勤務.Where(a => a.RowState != DataRowState.Deleted && a.RowState != DataRowState.Detached && a.ID == sID);

                // foreach用の配列を作成する
                var list = d.ToList();

                // 削除
                foreach (var it in list)
                {
                    DataSet1.帰宅後勤務Row dl = dts.帰宅後勤務.FindByID(it.ID);
                    dl.Delete();
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "削除失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
        }

    }
}
