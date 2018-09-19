using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using GrapeCity.Win.MultiRow;
using SZDS_TIMECARD.Common;
using System.Data.SqlClient;

namespace SZDS_TIMECARD.OCR
{
    public partial class frmKitakugo : Form
    {
        public frmKitakugo(string dbName, int mID, string gID, string [] _hArray, xlsData exlms, bool sKbn)
        {
            InitializeComponent();

            sMID = mID;
            _gID = gID;
            _dbName = dbName;
            hArray = _hArray;
            bs = exlms;
            _sKbn = sKbn;

            // データ読み込み
            kAdp.Fill(dts.帰宅後勤務);

            if (_sKbn)
            {
                hAdp.Fill(dts.勤務票ヘッダ);
                mAdp.Fill(dts.勤務票明細);
            }
            else
            {
                phAdp.Fill(dts.過去勤務票ヘッダ);
                pmAdp.Fill(dts.過去勤務票明細);
            }
        }

        string _dbName = string.Empty;
        string[] hArray = null;
        Common.xlsData bs;
        private int sMID { get; set; }
        bool _sKbn = false;

        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.帰宅後勤務TableAdapter kAdp = new DataSet1TableAdapters.帰宅後勤務TableAdapter();
        DataSet1TableAdapters.勤務票ヘッダTableAdapter hAdp = new DataSet1TableAdapters.勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.勤務票明細TableAdapter mAdp = new DataSet1TableAdapters.勤務票明細TableAdapter();
        DataSet1TableAdapters.過去勤務票ヘッダTableAdapter phAdp = new DataSet1TableAdapters.過去勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.過去勤務票明細TableAdapter pmAdp = new DataSet1TableAdapters.過去勤務票明細TableAdapter();

        global gl = new global();
        int fMode = 0;
        string kID = string.Empty;    // 帰宅後勤務レコードID
        string _gID = string.Empty;

        private void frmKitakugo_Load(object sender, EventArgs e)
        {
            gcMrSetting();

            if (_sKbn)
            {
                dataLoad();
            }
            else
            {
                pastDataLoad();
            }
        }

        private void dataLoad()
        {
            if (dts.帰宅後勤務.Any(a => a.勤務票帰宅後ID == _gID))
            {
                // データ登録済み
                var s = dts.帰宅後勤務.Single(a => a.勤務票帰宅後ID == _gID);

                gcMultiRow1[0, "txtYear"].Value = s.年.ToString();
                gcMultiRow1[0, "txtMonth"].Value = s.月.ToString();
                gcMultiRow1[0, "txtDay"].Value = s.日.ToString();
                gcMultiRow1[0, "txtShainNum"].Value = s.社員番号.ToString();
                gcMultiRow1[0, "txtJiyu1"].Value = s.事由1;
                gcMultiRow1[0, "txtJiyu2"].Value = s.事由2;
                gcMultiRow1[0, "txtJiyu3"].Value = s.事由3;

                //gcMultiRow1[0, "txtSftCode"].Value = s.シフトコード;

                // 勤務体系（シフト）コード
                var m = dts.勤務票明細.Single(a => a.ID == s.勤務票明細Row.ID);
                if (m.シフトコード != string.Empty)
                {
                    // シフト変更のとき
                    gcMultiRow1[0, "txtSftCode"].Value = m.シフトコード;
                }
                else if (m.勤務票ヘッダRow != null)
                    {
                        // 既定のシフトコードのとき
                        gcMultiRow1[0, "txtSftCode"].Value = m.勤務票ヘッダRow.シフトコード;
                    }
                else
                {
                    gcMultiRow1[0, "txtSftCode"].Value = string.Empty;
                }

                // 部署コード
                if (m.勤務票ヘッダRow != null)
                {
                    gcMultiRow1[0, "txtBushoCode"].Value = m.勤務票ヘッダRow.部署コード;
                }
                else
                {
                    gcMultiRow1[0, "txtBushoCode"].Value = string.Empty;
                }

                gcMultiRow1[0, "txtSh"].Value = s.出勤時;
                gcMultiRow1[0, "txtSm"].Value = s.出勤分;
                gcMultiRow1[0, "txtEh"].Value = s.退勤時;
                gcMultiRow1[0, "txtEm"].Value = s.退勤分;
                gcMultiRow1[0, "txtZanRe1"].Value = s.残業理由1;
                gcMultiRow1[0, "txtZanH1"].Value = s.残業時1;
                gcMultiRow1[0, "txtZanM1"].Value = s.残業分1;
                gcMultiRow1[0, "txtZanRe2"].Value = s.残業理由2;
                gcMultiRow1[0, "txtZanH2"].Value = s.残業時2;
                gcMultiRow1[0, "txtZanM2"].Value = s.残業分2;

                // データ編集モード
                fMode = global.FORM_EDITMODE;
                kID = _gID;

                // 編集モード
                btnDel.Enabled = true;
                button1.Enabled = true;
            }
            else
            {
                if (dts.勤務票明細.Any(a => a.ID == sMID))
                {
                    var s = dts.勤務票明細.Single(a => a.ID == sMID);

                    // 新規登録のとき
                    if (s.勤務票ヘッダRow != null)
                    {
                        gcMultiRow1[0, "txtYear"].Value = s.勤務票ヘッダRow.年.ToString();
                        gcMultiRow1[0, "txtMonth"].Value = s.勤務票ヘッダRow.月.ToString();
                        gcMultiRow1[0, "txtDay"].Value = s.勤務票ヘッダRow.日.ToString();
                    }
                    else
                    {
                        gcMultiRow1[0, "txtMonth"].Value = string.Empty;
                        gcMultiRow1[0, "txtDay"].Value = string.Empty;
                        gcMultiRow1[0, "txtShainNum"].Value = string.Empty;
                    }
                    
                    gcMultiRow1[0, "txtShainNum"].Value = s.社員番号.ToString();

                    // 勤務体系（シフト）コード
                    if (s.シフトコード != string.Empty)
                    {
                        // シフト変更のとき
                        gcMultiRow1[0, "txtSftCode"].Value = s.シフトコード;
                    }
                    else
                    {
                        // 既定のシフトコードのとき
                        gcMultiRow1[0, "txtSftCode"].Value = s.勤務票ヘッダRow.シフトコード;
                    }

                    // 部署コード
                    if (s.勤務票ヘッダRow != null)
                    {
                        gcMultiRow1[0, "txtBushoCode"].Value = s.勤務票ヘッダRow.部署コード;
                    }
                    else
                    {
                        gcMultiRow1[0, "txtBushoCode"].Value = string.Empty;
                    }
                }

                // データ追加モード
                fMode = global.FORM_ADDMODE;

                DateTime dt = DateTime.Now;
                kID = dt.Year.ToString() + dt.Month.ToString().PadLeft(2, '0') + dt.Day.ToString().PadLeft(2, '0') + dt.Hour.ToString().PadLeft(2, '0') + dt.Minute.ToString().PadLeft(2, '0') + dt.Second.ToString().PadLeft(2, '0'); 
                btnDel.Enabled = false;
            }
        }


        private void pastDataLoad()
        {
            if (dts.帰宅後勤務.Any(a => a.勤務票帰宅後ID == _gID))
            {
                // データ登録済み
                var s = dts.帰宅後勤務.Single(a => a.勤務票帰宅後ID == _gID);

                gcMultiRow1[0, "txtYear"].Value = s.年.ToString();
                gcMultiRow1[0, "txtMonth"].Value = s.月.ToString();
                gcMultiRow1[0, "txtDay"].Value = s.日.ToString();
                gcMultiRow1[0, "txtShainNum"].Value = s.社員番号.ToString();
                gcMultiRow1[0, "txtJiyu1"].Value = s.事由1;
                gcMultiRow1[0, "txtJiyu2"].Value = s.事由2;
                gcMultiRow1[0, "txtJiyu3"].Value = s.事由3;

                //gcMultiRow1[0, "txtSftCode"].Value = s.シフトコード;

                // 勤務体系（シフト）コード
                var m = dts.過去勤務票明細.Single(a => a.ID == s.過去勤務票明細Row.ID);
                if (m.シフトコード != string.Empty)
                {
                    // シフト変更のとき
                    gcMultiRow1[0, "txtSftCode"].Value = m.シフトコード;
                }
                else if (m.過去勤務票ヘッダRow != null)
                {
                    // 既定のシフトコードのとき
                    gcMultiRow1[0, "txtSftCode"].Value = m.過去勤務票ヘッダRow.シフトコード;
                }
                else
                {
                    gcMultiRow1[0, "txtSftCode"].Value = string.Empty;
                }

                // 部署コード
                if (m.過去勤務票ヘッダRow != null)
                {
                    gcMultiRow1[0, "txtBushoCode"].Value = m.過去勤務票ヘッダRow.部署コード;
                }
                else
                {
                    gcMultiRow1[0, "txtBushoCode"].Value = string.Empty;
                }

                gcMultiRow1[0, "txtSh"].Value = s.出勤時;
                gcMultiRow1[0, "txtSm"].Value = s.出勤分;
                gcMultiRow1[0, "txtEh"].Value = s.退勤時;
                gcMultiRow1[0, "txtEm"].Value = s.退勤分;
                gcMultiRow1[0, "txtZanRe1"].Value = s.残業理由1;
                gcMultiRow1[0, "txtZanH1"].Value = s.残業時1;
                gcMultiRow1[0, "txtZanM1"].Value = s.残業分1;
                gcMultiRow1[0, "txtZanRe2"].Value = s.残業理由2;
                gcMultiRow1[0, "txtZanH2"].Value = s.残業時2;
                gcMultiRow1[0, "txtZanM2"].Value = s.残業分2;

                // データ編集モード
                fMode = global.FORM_EDITMODE;
                kID = _gID;

                // 過去データ閲覧のとき
                btnDel.Visible = false;
                button1.Visible = false;
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

            // 閲覧モード
            if (!_sKbn)
            {
                gcMultiRow1.ReadOnly = true;
            }
            else
            {
                gcMultiRow1.ReadOnly = false;
            }
        }

        private void frmKitakugo_Shown(object sender, EventArgs e)
        {
            gcMultiRow1.CurrentCell = null;
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

        private void gcMultiRow1_EditingControlShowing(object sender, GrapeCity.Win.MultiRow.EditingControlShowingEventArgs e)
        {
            if (e.Control is TextBoxEditingControl)
            {
                //イベントハンドラが複数回追加されてしまうので最初に削除する
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress2);
                e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress3);

                // 数字のみ入力可能とする
                if (gcMultiRow1.CurrentCell.Name == "txtJiyu1" ||
                    gcMultiRow1.CurrentCell.Name == "txtJiyu2" || gcMultiRow1.CurrentCell.Name == "txtJiyu3" ||
                    gcMultiRow1.CurrentCell.Name == "txtSh" || gcMultiRow1.CurrentCell.Name == "txtSm" || 
                    gcMultiRow1.CurrentCell.Name == "txtEh" || gcMultiRow1.CurrentCell.Name == "txtEm" ||
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

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmKitakugo_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
        }

        private void gcMultiRow1_CellValueChanged(object sender, CellEventArgs e)
        {
            if (!gl.ChangeValueStatus) return;

            if (e.RowIndex < 0) return;

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

                //// 対象のシフトコードの開始時間と異なるときバックカラーを変更する
                //chkSftStartTime(e.RowIndex);

                // ChangeValueイベントステータスをtrueに戻す
                gl.ChangeValueStatus = true;
            }
        }

        private void gcMultiRow1_CellEnter(object sender, CellEventArgs e)
        {
            if (gcMultiRow1.EditMode == EditMode.EditProgrammatically)
            {
                gcMultiRow1.BeginEdit(true);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string msg = string.Empty;

            if (errCheck())
            {
                if (MessageBox.Show("帰宅後勤務データを登録します。よろしいですか", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                {
                    return;
                }

                if (fMode == global.FORM_ADDMODE)
                {
                    dataAdd();
                    msg = "データを追加登録しました";

                    // 勤務票明細に帰宅後勤務IDを書き込み
                    var s = dts.勤務票明細.Single(a => a.ID == sMID);
                    s.帰宅後勤務ID = kID;
                    mAdp.Update(dts);
                }
                else if (fMode == global.FORM_EDITMODE)
                {
                    dataUpdate();
                    msg = "データを更新しました";
                }

                kAdp.Update(dts.帰宅後勤務);

                MessageBox.Show(msg, "帰宅後勤務", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // フォームを閉じる
                this.Close();
            }
        }

        private bool errCheck()
        {
            bool rtn = true;

            string sDate;
            DateTime eDate;

            // 対象年月日
            sDate = Utility.NulltoStr(gcMultiRow1[0, "txtYear"].Value) + "/" + Utility.NulltoStr(gcMultiRow1[0, "txtMonth"].Value) + "/" + Utility.NulltoStr(gcMultiRow1[0, "txtDay"].Value);
            if (!DateTime.TryParse(sDate, out eDate))
            {
                MessageBox.Show("年月日が正しくありません", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtYear"];
                return false;
            }

            int iX = 0;
            string k = string.Empty;    // 特別休暇記号
            string yk = string.Empty;   // 有給記号

            // 事由コード
            if (Utility.NulltoStr(gcMultiRow1[0, "txtJiyu1"].Value) != string.Empty)
            {
                if (!Utility.chkJiyu(gcMultiRow1[0, "txtJiyu1"].Value.ToString(), _dbName))
                {
                    MessageBox.Show("事由１が正しくありません", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtJiyu1"];
                    return false;
                }
            }

            if (Utility.NulltoStr(gcMultiRow1[0, "txtJiyu2"].Value) != string.Empty)
            {
                if (!Utility.chkJiyu(gcMultiRow1[0, "txtJiyu2"].Value.ToString(), _dbName))
                {
                    MessageBox.Show("事由２が正しくありません", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtJiyu2"];
                    return false;
                }
            }

            if (Utility.NulltoStr(gcMultiRow1[0, "txtJiyu3"].Value) != string.Empty)
            {
                if (!Utility.chkJiyu(gcMultiRow1[0, "txtJiyu3"].Value.ToString(), _dbName))
                {
                    MessageBox.Show("事由３が正しくありません", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtJiyu3"];
                    return false;
                }
            }

            // 始業時刻・終業時刻チェック
            if (!errCheckTime(gcMultiRow1, 1)) return false;

            // 残業理由
            if (!Utility.chkZangyoRe(Utility.NulltoStr(gcMultiRow1[0, "txtZanRe1"].Value), 
                Utility.NulltoStr(gcMultiRow1[0, "txtZanH1"].Value), 
                Utility.NulltoStr(gcMultiRow1[0, "txtZanM1"].Value)))
            {
                MessageBox.Show("残業理由が未記入です", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtZanRe1"];
                return false;
            }

            if (!Utility.chkZangyoRe2(Utility.NulltoStr(gcMultiRow1[0, "txtZanRe1"].Value), 
                Utility.NulltoStr(gcMultiRow1[0, "txtZanH1"].Value), 
                Utility.NulltoStr(gcMultiRow1[0, "txtZanM1"].Value)))
            {
                MessageBox.Show("残業時間が未記入です", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtZanH1"];
                return false;
            }

            // 部署別残業理由Excelシート登録チェック
            string reName = string.Empty;
            if (Utility.NulltoStr(gcMultiRow1[0, "txtZanRe1"].Value) != string.Empty)
            {
                if (!bs.getBushoZanRe(out reName, gcMultiRow1[0, "txtBushoCode"].Value.ToString(), gcMultiRow1[0, "txtZanRe1"].Value.ToString()))
                {
                    MessageBox.Show("該当部署に登録されていない残業理由です", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtZanRe1"];
                    return false;
                }
            }

            // 残業理由2
            if (!Utility.chkZangyoRe(Utility.NulltoStr(gcMultiRow1[0, "txtZanRe2"].Value), 
                Utility.NulltoStr(gcMultiRow1[0, "txtZanH2"].Value), 
                Utility.NulltoStr(gcMultiRow1[0, "txtZanM2"].Value)))
            {
                MessageBox.Show("残業理由が未記入です", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtZanRe2"];
                return false;
            }

            if (!Utility.chkZangyoRe2(Utility.NulltoStr(gcMultiRow1[0, "txtZanRe2"].Value), 
                Utility.NulltoStr(gcMultiRow1[0, "txtZanH2"].Value), 
                Utility.NulltoStr(gcMultiRow1[0, "txtZanM2"].Value)))
            {
                MessageBox.Show("残業時間が未記入です", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtZanH2"];
                return false;
            }

            // 部署別残業理由Excelシート登録チェック
            if (Utility.NulltoStr(gcMultiRow1[0, "txtZanRe2"].Value) != string.Empty)
            {
                if (!bs.getBushoZanRe(out reName, gcMultiRow1[0, "txtBushoCode"].Value.ToString(), gcMultiRow1[0, "txtZanRe2"].Value.ToString()))
                {
                    MessageBox.Show("該当部署に登録されていない残業理由です", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtZanRe2"];
                    return false;
                }
            }

            //// 残業分単位
            //if (!chkZangyoMin(m.残業分1))
            //{
            //    setErrStatus(eZanM1, iX - 1, "残業分単位は０または５です");
            //    return false;
            //}

            //if (!chkZangyoMin(m.残業分2))
            //{
            //    setErrStatus(eZanM2, iX - 1, "残業分単位は０または５です");
            //    return false;
            //}

            // 残業と出退勤時刻記入
            if (!errCheckZanShEh(gcMultiRow1))
            {
                MessageBox.Show("残業があるときは出退勤時刻の記入が必要です", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtSh"];
                return false;
            }
            
            return rtn;
        }

        ///---------------------------------------------------------------------------
        /// <summary>
        ///     残業のとき出勤退時刻が記入されているか </summary>
        /// <param name="m">
        ///     DataSet1.勤務票明細Row　</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///---------------------------------------------------------------------------
        private bool errCheckZanShEh(GcMultiRow gc)
        {
            bool rtn = true;

            if (Utility.NulltoStr(gcMultiRow1[0, "txtZanRe1"].Value) != string.Empty || 
                Utility.NulltoStr(gcMultiRow1[0, "txtZanH1"].Value) != string.Empty ||
                Utility.NulltoStr(gcMultiRow1[0, "txtZanM1"].Value) != string.Empty || 
                Utility.NulltoStr(gcMultiRow1[0, "txtZanRe2"].Value) != string.Empty ||
                Utility.NulltoStr(gcMultiRow1[0, "txtZanH2"].Value) != string.Empty ||
                Utility.NulltoStr(gcMultiRow1[0, "txtZanM2"].Value) != string.Empty)
            {
                if (Utility.NulltoStr(gcMultiRow1[0, "txtSh"].Value) == string.Empty || 
                    Utility.NulltoStr(gcMultiRow1[0, "txtSm"].Value) == string.Empty || 
                    Utility.NulltoStr(gcMultiRow1[0, "txtEh"].Value) == string.Empty || 
                    Utility.NulltoStr(gcMultiRow1[0, "txtEm"].Value) == string.Empty)
                {
                    return false;
                }
            }

            return rtn;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     時間記入チェック </summary>
        /// <param name="obj">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="Tani">
        ///     分記入単位</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckTime(GcMultiRow gc, int Tani)
        {
            // 出勤時間と退勤時間
            string Sh = Utility.NulltoStr(gc[0, "txtSh"].Value).Trim();
            string Sm = Utility.NulltoStr(gc[0, "txtSm"].Value).Trim();
            string Eh = Utility.NulltoStr(gc[0, "txtEh"].Value).Trim();
            string Em = Utility.NulltoStr(gc[0, "txtEm"].Value).Trim();

            string sTimeW = Sh + Sm;
            string eTimeW = Eh + Em;

            if (sTimeW == string.Empty && eTimeW == string.Empty)
            {
                MessageBox.Show("出退勤時刻が未入力です", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtSh"];
                return false;
            }

            if (sTimeW != string.Empty && eTimeW == string.Empty)
            {
                MessageBox.Show("退勤時刻が未入力です", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtEh"];
                return false;
            }

            if (sTimeW == string.Empty && eTimeW != string.Empty)
            {
                MessageBox.Show("出勤時刻が未入力です", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                gcMultiRow1.Focus();
                gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtSh"];
                return false;
            }

            // 記入のとき
            if (Sh != string.Empty || Sm != string.Empty || Eh != string.Empty || Em != string.Empty) 
            {
                // 数字範囲、単位チェック
                if (!Utility.checkHourSpan(Sh))
                {
                    MessageBox.Show("出勤時刻が正しくありません", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtSh"];
                    return false;
                }

                if (!Utility.checkMinSpan(Sm, Tani))
                {
                    MessageBox.Show("出勤時刻が正しくありません", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtSm"];
                    return false;
                }

                if (!Utility.checkHourSpan(Eh))
                {
                    MessageBox.Show("退勤時刻が正しくありません", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtEh"];
                    return false;
                }

                if (!Utility.checkMinSpan(Em, Tani))
                {
                    MessageBox.Show("退勤時刻が正しくありません", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtEm"];
                    return false;
                }
            }

            return true;
        }

        ///-----------------------------------------------------------------
        /// <summary>
        ///     帰宅後勤務データ新規登録 </summary>
        ///-----------------------------------------------------------------
        private void dataAdd()
        {
            try
            {
                DataSet1.帰宅後勤務Row r = dts.帰宅後勤務.New帰宅後勤務Row();
                r.勤務票帰宅後ID = kID;
                r.年 = Utility.StrtoInt(gcMultiRow1[0, "txtYear"].Value.ToString());
                r.月 = Utility.StrtoInt(gcMultiRow1[0, "txtMonth"].Value.ToString());
                r.日 = Utility.StrtoInt(gcMultiRow1[0, "txtDay"].Value.ToString());
                r.社員番号 = gcMultiRow1[0, "txtShainNum"].Value.ToString();
                r.出勤時 = gcMultiRow1[0, "txtSh"].Value.ToString();
                r.出勤分 = gcMultiRow1[0, "txtSm"].Value.ToString();
                r.退勤時 = gcMultiRow1[0, "txtEh"].Value.ToString();
                r.退勤分 = gcMultiRow1[0, "txtEm"].Value.ToString();

                // 残業理由１：先頭ゼロは除去
                string sN = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[0, "txtZanRe1"].Value)).ToString();
                if (sN != global.FLGOFF)
                {
                    // 残業理由記入あり
                    r.残業理由1 = sN;
                    r.残業時1 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[0, "txtZanH1"].Value)).ToString();
                    r.残業分1 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[0, "txtZanM1"].Value)).ToString();
                }
                else
                {
                    // 残業理由記入なし
                    r.残業理由1 = string.Empty;
                    r.残業時1 = string.Empty;
                    r.残業分1 = string.Empty;
                }

                // 残業理由２：先頭ゼロは除去
                sN = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[0, "txtZanRe2"].Value)).ToString();
                if (sN != global.FLGOFF)
                {
                    // 残業理由記入あり
                    r.残業理由2 = sN;
                    r.残業時2 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[0, "txtZanH2"].Value)).ToString();
                    r.残業分2 = Utility.StrtoInt(Utility.NulltoStr(gcMultiRow1[0, "txtZanM2"].Value)).ToString();
                }
                else
                {
                    // 残業理由記入なし
                    r.残業理由2 = string.Empty;
                    r.残業時2 = string.Empty;
                    r.残業分2 = string.Empty;
                }

                r.事由1 = Utility.NulltoStr(gcMultiRow1[0, "txtJiyu1"].Value);
                r.事由2 = Utility.NulltoStr(gcMultiRow1[0, "txtJiyu2"].Value);
                r.事由3 = Utility.NulltoStr(gcMultiRow1[0, "txtJiyu3"].Value);
                //r.シフトコード = gcMultiRow1[0, "txtSftCode"].Value.ToString();
                r.取消 = string.Empty;
                r.データ領域名 = _dbName;
                r.更新年月日 = DateTime.Now;

                dts.帰宅後勤務.Add帰宅後勤務Row(r);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        
        ///-----------------------------------------------------------------
        /// <summary>
        ///     帰宅後勤務データ更新 </summary>
        ///-----------------------------------------------------------------
        private void dataUpdate()
        {
            try
            {
                DataSet1.帰宅後勤務Row r = dts.帰宅後勤務.Single(a => a.勤務票帰宅後ID == kID);
                r.年 = Utility.StrtoInt(gcMultiRow1[0, "txtYear"].Value.ToString());
                r.月 = Utility.StrtoInt(gcMultiRow1[0, "txtMonth"].Value.ToString());
                r.日 = Utility.StrtoInt(gcMultiRow1[0, "txtDay"].Value.ToString());
                r.社員番号 = gcMultiRow1[0, "txtShainNum"].Value.ToString();
                r.出勤時 = gcMultiRow1[0, "txtSh"].Value.ToString();
                r.出勤分 = gcMultiRow1[0, "txtSm"].Value.ToString();
                r.退勤時 = gcMultiRow1[0, "txtEh"].Value.ToString();
                r.退勤分 = gcMultiRow1[0, "txtEm"].Value.ToString();
                r.残業理由1 = Utility.NulltoStr(gcMultiRow1[0, "txtZanRe1"].Value);
                r.残業時1 = Utility.NulltoStr(gcMultiRow1[0, "txtZanH1"].Value);
                r.残業分1 = Utility.NulltoStr(gcMultiRow1[0, "txtZanM1"].Value);
                r.残業理由2 = Utility.NulltoStr(gcMultiRow1[0, "txtZanRe2"].Value);
                r.残業時2 = Utility.NulltoStr(gcMultiRow1[0, "txtZanH2"].Value);
                r.残業分2 = Utility.NulltoStr(gcMultiRow1[0, "txtZanM2"].Value);
                r.事由1 = Utility.NulltoStr(gcMultiRow1[0, "txtJiyu1"].Value);
                r.事由2 = Utility.NulltoStr(gcMultiRow1[0, "txtJiyu2"].Value);
                r.事由3 = Utility.NulltoStr(gcMultiRow1[0, "txtJiyu3"].Value);
                //r.シフトコード = gcMultiRow1[0, "txtSftCode"].Value.ToString();
                r.取消 = string.Empty;
                r.データ領域名 = _dbName;
                r.更新年月日 = DateTime.Now;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("帰宅後勤務データを削除します。よろしいですか", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            dataDelete();
            kAdp.Update(dts.帰宅後勤務);

            // 勤務票明細に帰宅後勤務IDを書き込み
            var s = dts.勤務票明細.Single(a => a.ID == sMID);
            s.帰宅後勤務ID = string.Empty;
            mAdp.Update(dts);

            MessageBox.Show("データを削除しました", "帰宅後勤務", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // フォームを閉じる
            this.Close();
        }

        private void dataDelete()
        {
            DataSet1.帰宅後勤務Row r = dts.帰宅後勤務.Single(a => a.勤務票帰宅後ID == kID);
            r.Delete();
        }
    }
}
