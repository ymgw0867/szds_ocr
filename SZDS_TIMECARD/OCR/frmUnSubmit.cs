using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using SZDS_TIMECARD.Common;


namespace SZDS_TIMECARD.OCR
{
    public partial class frmUnSubmit : Form
    {
        public frmUnSubmit(string DBName, string ComName)
        {
            InitializeComponent();

            _DBName = DBName;
            _ComName = ComName;

            hAdp.Fill(dts.過去勤務票ヘッダ);
            iAdp.Fill(dts.過去勤務票明細);
            ohAdp.Fill(dts.過去応援移動票ヘッダ);
            oiAdp.Fill(dts.過去応援移動票明細);
        }

        string _DBName = string.Empty;
        string _ComName = string.Empty;

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.過去勤務票ヘッダTableAdapter hAdp = new DataSet1TableAdapters.過去勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.過去勤務票明細TableAdapter iAdp = new DataSet1TableAdapters.過去勤務票明細TableAdapter();
        DataSet1TableAdapters.過去応援移動票ヘッダTableAdapter ohAdp = new DataSet1TableAdapters.過去応援移動票ヘッダTableAdapter();
        DataSet1TableAdapters.過去応援移動票明細TableAdapter oiAdp = new DataSet1TableAdapters.過去応援移動票明細TableAdapter();

        private void frmUnSubmit_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            // 所属名コンボボックスのデータソースをセットする
            Utility.ComboBumon.loadBusho(comboBox2, _DBName);

            // データグリッド定義
            GridviewSet(dataGridView1);

            // 画面初期化
            DispClear();

            // 会社領域名表示
            this.Text += "【" + _ComName + "】";
        }

        private void dataComboSet()
        {
            comboBox2.Items.Clear();

            // 過去勤務票明細の所属名を取得する
            OleDbCommand sCom = new OleDbCommand();
            mdbControl dCon = new mdbControl();
            dCon.dbConnect(sCom);
            StringBuilder sb = new StringBuilder();

            sb.Append("select distinct 部署コード from ");
            sb.Append("(SELECT 過去勤務票明細.部署コード from ");
            sb.Append("過去勤務票ヘッダ inner join 過去勤務票明細 ");
            sb.Append("on 過去勤務票ヘッダ.ID = 過去勤務票明細.ヘッダID ");
            sb.Append("where 過去勤務票ヘッダ.データ領域名 = ?) as a");

            sCom.CommandText = sb.ToString();
            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@db", _DBName);
            OleDbDataReader dR = sCom.ExecuteReader();

            while (dR.Read())
            {
                comboBox2.Items.Add(dR["所属名"].ToString());
            }
            dR.Close();

            sCom.Connection.Close();
        }

        // カラム定義
        private string ColDate = "c0";
        private string ColSz = "c1";
        private string ColSznm = "c2";
        private string ColCode = "c3";
        private string ColName = "c4";
        private string ColMemo = "c5";
        private string Colet = "c6";
        private string Colzn = "c7";
        private string Colsi = "c8";
        private string ColID = "c9";
        private string ColKinmuCode = "c10";

        /// <summary>
        /// データグリッドビューの定義を行います
        /// </summary>
        private void GridviewSet(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 10, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", 10, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                tempDGV.ColumnHeadersHeight = 22;
                tempDGV.RowTemplate.Height = 22;

                // 全体の高さ
                tempDGV.Height = 530;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;

                // 各列幅指定
                tempDGV.Columns.Add(ColDate, "年月日");
                tempDGV.Columns.Add(ColSz, "コード");
                tempDGV.Columns.Add(ColSznm, "部署名");
                tempDGV.Columns.Add(ColCode, "社員番号");
                tempDGV.Columns.Add(ColName, "氏名");
                tempDGV.Columns.Add(ColMemo, "備考");
                tempDGV.Columns.Add(ColID, "id");

                tempDGV.Columns[ColDate].Width = 110;
                tempDGV.Columns[ColSz].Width = 80;
                tempDGV.Columns[ColSznm].Width = 240;
                tempDGV.Columns[ColCode].Width = 80;
                tempDGV.Columns[ColName].Width = 170;
                tempDGV.Columns[ColMemo].Width = 110;

                tempDGV.Columns[ColID].Visible = false;

                tempDGV.Columns[ColName].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[ColDate].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[ColSz].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[ColCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                // 編集可否
                tempDGV.ReadOnly = true;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = true;

                // 追加行表示しない
                tempDGV.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                tempDGV.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                tempDGV.AllowUserToOrderColumns = false;

                // 列サイズ変更禁止
                tempDGV.AllowUserToResizeColumns = true;

                // 行サイズ変更禁止
                tempDGV.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 画面初期化
        /// </summary>
        private void DispClear()
        {
            txtYear.Text = string.Empty;
            txtMonth.Text = string.Empty;
            txtDay.Text = string.Empty;
            comboBox2.Text = string.Empty;
            comboBox2.SelectedIndex = -1;

            comboBox1.SelectedIndex = 0;
        }

        private void btnSel_Click(object sender, EventArgs e)
        {
        }

        private bool errCheck()
        {
            try 
	        {
                if (txtYear.Text != string.Empty && !Utility.NumericCheck(txtYear.Text))
                {
                    txtYear.Focus();
                    throw new Exception("年が正しくありません");
                }

                if (txtMonth.Text != string.Empty && !Utility.NumericCheck(txtMonth.Text))
                {
                    txtMonth.Focus();
                    throw new Exception("月が正しくありません");
                }

                if (txtMonth.Text != string.Empty)
                {
                    if (int.Parse(txtMonth.Text) < 1 || int.Parse(txtMonth.Text) > 12)
                    {
                        txtMonth.Focus();
                        throw new Exception("月が正しくありません");
                    }
                }

                if (txtDay.Text != string.Empty)
                {
                    if (int.Parse(txtDay.Text) < 1 || int.Parse(txtDay.Text) > 31)
                    {
                        txtDay.Focus();
                        throw new Exception("日が正しくありません");
                    }
                }
            }
	        catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
	        }
            return true;
        }

        /// <summary>
        /// データ検索
        /// </summary>
        private void DataSelect(DataGridView g)
        {
            this.Cursor = Cursors.WaitCursor;

            //dbControl.DataControl dCon = new dbControl.DataControl(_DBName);

            // 奉行SQLServer接続文字列取得
            string sc = sqlControl.obcConnectSting.get(_DBName);
            sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            StringBuilder sb = new StringBuilder();

            // データグリッドビューの表示を初期化する
            g.Rows.Clear();
            
            var s = dts.過去勤務票明細.Where(a => a.過去勤務票ヘッダRow != null)
                .OrderByDescending(a => a.過去勤務票ヘッダRow.年).ThenByDescending(a => a.過去勤務票ヘッダRow.月).ThenByDescending(a => a.過去勤務票ヘッダRow.日).ThenBy(a => a.過去勤務票ヘッダRow.部署コード).ThenBy(a => a.社員番号);
            
            s = s.Where(a => a.過去勤務票ヘッダRow.データ領域名 == _DBName)
                .OrderByDescending(a => a.過去勤務票ヘッダRow.年).ThenByDescending(a => a.過去勤務票ヘッダRow.月).ThenByDescending(a => a.過去勤務票ヘッダRow.日).ThenBy(a => a.過去勤務票ヘッダRow.部署コード).ThenBy(a => a.社員番号);

            if (txtYear.Text.Trim() != string.Empty)
            {
                s = s.Where(a => a.過去勤務票ヘッダRow.年 == Utility.StrtoInt(txtYear.Text))
                    .OrderByDescending(a => a.過去勤務票ヘッダRow.年).ThenByDescending(a => a.過去勤務票ヘッダRow.月).ThenByDescending(a => a.過去勤務票ヘッダRow.日).ThenBy(a => a.過去勤務票ヘッダRow.部署コード).ThenBy(a => a.社員番号);
            }

            if (txtMonth.Text.Trim() != string.Empty)
            {
                s = s.Where(a => a.過去勤務票ヘッダRow.月 == Utility.StrtoInt(txtMonth.Text))
                   .OrderByDescending(a => a.過去勤務票ヘッダRow.年).ThenByDescending(a => a.過去勤務票ヘッダRow.月).ThenByDescending(a => a.過去勤務票ヘッダRow.日).ThenBy(a => a.過去勤務票ヘッダRow.部署コード).ThenBy(a => a.社員番号);
            }

            if (txtDay.Text.Trim() != string.Empty)
            {
                s = s.Where(a => a.過去勤務票ヘッダRow.日 == Utility.StrtoInt(txtDay.Text))
                    .OrderByDescending(a => a.過去勤務票ヘッダRow.年).ThenByDescending(a => a.過去勤務票ヘッダRow.月).ThenByDescending(a => a.過去勤務票ヘッダRow.日).ThenBy(a => a.過去勤務票ヘッダRow.部署コード).ThenBy(a => a.社員番号);
            }
            
            if (txtShainNum.Text.Trim() != string.Empty)
            {
                s = s.Where(a => a.社員番号 == txtShainNum.Text)
                    .OrderByDescending(a => a.過去勤務票ヘッダRow.年).ThenByDescending(a => a.過去勤務票ヘッダRow.月).ThenByDescending(a => a.過去勤務票ヘッダRow.日).ThenBy(a => a.過去勤務票ヘッダRow.部署コード).ThenBy(a => a.社員番号);
            }

            if (comboBox2.SelectedIndex != -1)
            {
                Utility.ComboBumon cmb = (Utility.ComboBumon)comboBox2.SelectedItem;
                s = s.Where(a => a.過去勤務票ヘッダRow.部署コード == cmb.code.ToString())
                    .OrderByDescending(a => a.過去勤務票ヘッダRow.年).ThenByDescending(a => a.過去勤務票ヘッダRow.月).ThenByDescending(a => a.過去勤務票ヘッダRow.日).ThenBy(a => a.過去勤務票ヘッダRow.部署コード).ThenBy(a => a.社員番号);
            }


            foreach (var t in s)
            {
                g.Rows.Add();
                g[ColDate, g.Rows.Count - 1].Value = t.過去勤務票ヘッダRow.年.ToString() + "/" + t.過去勤務票ヘッダRow.月.ToString() + "/" + t.過去勤務票ヘッダRow.日.ToString();
                g[ColSz, g.Rows.Count - 1].Value = t.過去勤務票ヘッダRow.部署コード.ToString();

                // 奉行データベースより部署名を取得して表示します
                // DepartmentCode（部署コード）
                string strCode = string.Empty;
                if (Utility.NumericCheck(t.過去勤務票ヘッダRow.部署コード.ToString()))
                {
                    strCode = t.過去勤務票ヘッダRow.部署コード.ToString().PadLeft(15, '0');
                }
                else
                {
                    strCode = t.過去勤務票ヘッダRow.部署コード.ToString().PadRight(15, ' ');
                }

                g[ColSznm, g.Rows.Count - 1].Value = Utility.getDepartmentName(sdCon, strCode);
                g[ColCode, g.Rows.Count - 1].Value = t.社員番号.PadLeft(6, '0');                
                g[ColName, g.Rows.Count - 1].Value = t.社員名;
                g[ColMemo, g.Rows.Count - 1].Value = string.Empty;

                g[ColID, g.Rows.Count - 1].Value = t.過去勤務票ヘッダRow.ID.ToString();
            }
            
            dataGridView1.CurrentCell = null;

            sdCon.Close();

            // 終了
            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("該当するデータはありませんでした", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Text = "【" + _ComName + "】";
            }
            else
            {
                this.Text = "【" + _ComName + "】 " + dataGridView1.RowCount.ToString("#,##0") + "件"; 
            }

            this.Cursor = Cursors.Default;
        }


        ///-------------------------------------------------------------------
        /// <summary>
        ///     応援移動票データ検索 </summary>
        /// <param name="g">
        ///     datagridView オブジェクト</param>
        ///-------------------------------------------------------------------
        private void ouenDataSelect(DataGridView g)
        {
            this.Cursor = Cursors.WaitCursor;

            // 奉行SQLServer接続文字列取得
            string sc = sqlControl.obcConnectSting.get(_DBName);
            sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            // データグリッドビューの表示を初期化する
            g.Rows.Clear();

            var s = dts.過去応援移動票明細.Where(a => a.過去応援移動票ヘッダRow != null && a.過去応援移動票ヘッダRow.データ領域名 == _DBName)
                .OrderByDescending(a => a.過去応援移動票ヘッダRow.年).ThenByDescending(a => a.過去応援移動票ヘッダRow.月).ThenByDescending(a => a.過去応援移動票ヘッダRow.日).ThenBy(a => a.過去応援移動票ヘッダRow.部署コード).ThenBy(a => a.社員番号);

            if (txtYear.Text.Trim() != string.Empty)
            {
                s = s.Where(a => a.過去応援移動票ヘッダRow.年 == Utility.StrtoInt(txtYear.Text))
                    .OrderByDescending(a => a.過去応援移動票ヘッダRow.年).ThenByDescending(a => a.過去応援移動票ヘッダRow.月).ThenByDescending(a => a.過去応援移動票ヘッダRow.日).ThenBy(a => a.過去応援移動票ヘッダRow.部署コード).ThenBy(a => a.社員番号);
            }

            if (txtMonth.Text.Trim() != string.Empty)
            {
                s = s.Where(a => a.過去応援移動票ヘッダRow.月 == Utility.StrtoInt(txtMonth.Text))
                .OrderByDescending(a => a.過去応援移動票ヘッダRow.年).ThenByDescending(a => a.過去応援移動票ヘッダRow.月).ThenByDescending(a => a.過去応援移動票ヘッダRow.日).ThenBy(a => a.過去応援移動票ヘッダRow.部署コード).ThenBy(a => a.社員番号);
            }

            if (txtDay.Text.Trim() != string.Empty)
            {
                s = s.Where(a => a.過去応援移動票ヘッダRow.日 == Utility.StrtoInt(txtDay.Text))
                    .OrderByDescending(a => a.過去応援移動票ヘッダRow.年).ThenByDescending(a => a.過去応援移動票ヘッダRow.月).ThenByDescending(a => a.過去応援移動票ヘッダRow.日).ThenBy(a => a.過去応援移動票ヘッダRow.部署コード).ThenBy(a => a.社員番号);
            }

            if (txtShainNum.Text.Trim() != string.Empty)
            {
                s = s.Where(a => a.社員番号 == txtShainNum.Text)
                    .OrderByDescending(a => a.過去応援移動票ヘッダRow.年).ThenByDescending(a => a.過去応援移動票ヘッダRow.月).ThenByDescending(a => a.過去応援移動票ヘッダRow.日).ThenBy(a => a.過去応援移動票ヘッダRow.部署コード).ThenBy(a => a.社員番号);
            }

            if (comboBox2.SelectedIndex != -1)
            {
                Utility.ComboBumon cmb = (Utility.ComboBumon)comboBox2.SelectedItem;
                s = s.Where(a => a.過去応援移動票ヘッダRow.部署コード == cmb.code.ToString())
                    .OrderByDescending(a => a.過去応援移動票ヘッダRow.年).ThenByDescending(a => a.過去応援移動票ヘッダRow.月).ThenByDescending(a => a.過去応援移動票ヘッダRow.日).ThenBy(a => a.過去応援移動票ヘッダRow.部署コード).ThenBy(a => a.社員番号);
            }


            foreach (var t in s)
            {
                g.Rows.Add();
                g[ColDate, g.Rows.Count - 1].Value = t.過去応援移動票ヘッダRow.年.ToString() + "/" + t.過去応援移動票ヘッダRow.月.ToString() + "/" + t.過去応援移動票ヘッダRow.日.ToString();
                g[ColSz, g.Rows.Count - 1].Value = t.過去応援移動票ヘッダRow.部署コード.ToString();

                // 奉行データベースより部署名を取得して表示します
                // DepartmentCode（部署コード）
                string strCode = string.Empty;
                if (Utility.NumericCheck(t.過去応援移動票ヘッダRow.部署コード.ToString()))
                {
                    strCode = t.過去応援移動票ヘッダRow.部署コード.ToString().PadLeft(15, '0');
                }
                else
                {
                    strCode = t.過去応援移動票ヘッダRow.部署コード.ToString().PadRight(15, ' ');
                }

                g[ColSznm, g.Rows.Count - 1].Value = Utility.getDepartmentName(sdCon, strCode);
                g[ColCode, g.Rows.Count - 1].Value = t.社員番号.PadLeft(6, '0');

                if (t.Is社員名Null())
                {
                    g[ColName, g.Rows.Count - 1].Value = "";
                }
                else
                {
                    g[ColName, g.Rows.Count - 1].Value = t.社員名;
                }

                if (t.データ区分 == 1)
                {
                    g[ColMemo, g.Rows.Count - 1].Value = "日中";
                }
                else if (t.データ区分 == 2)
                {
                    g[ColMemo, g.Rows.Count - 1].Value = "残業";
                }

                g[ColID, g.Rows.Count - 1].Value = t.過去応援移動票ヘッダRow.ID.ToString();
            }

            dataGridView1.CurrentCell = null;
            sdCon.Close();

            // 終了
            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("該当するデータはありませんでした", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Text = "【" + _ComName + "】";
            }
            else
            {
                this.Text = "【" + _ComName + "】 " + dataGridView1.RowCount.ToString("#,##0") + "件";
            }


            this.Cursor = Cursors.Default;
        }

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     データグリッドへ表示する </summary>
        /// <param name="sdR">
        ///     データリーダーオブジェクト</param>
        /// <param name="stutas">
        ///     ステータス</param>
        /// <param name="g">
        ///     datagridviewオブジェクト</param>
        ///-----------------------------------------------------------------------
        //private void gridShow(OleDbDataReader sdR, DataGridView g)
        //{
        //    g.Rows.Add();
        //    g[ColDate, g.Rows.Count - 1].Value = sdR["年"].ToString() + "/" + sdR["月"].ToString() + "/" + sdR["日"].ToString();
        //    g[ColSz, g.Rows.Count - 1].Value = sdR["部署コード"].ToString();

        //    // 奉行データベースより部署名を取得して表示します
        //    if (Utility.NulltoStr(g[ColSz, g.Rows.Count - 1].Value) != string.Empty)
        //    {
        //        // DepartmentCode（部署コード）
        //        string strCode = string.Empty;
        //        if (Utility.NumericCheck(g[ColSz, g.Rows.Count - 1].Value.ToString()))
        //        {
        //            strCode = g[ColSz, g.Rows.Count - 1].Value.ToString().PadLeft(15, '0');
        //        }
        //        else
        //        {
        //            strCode = g[ColSz, g.Rows.Count - 1].Value.ToString().PadRight(15, ' ');
        //        }

        //        g[ColSznm, g.Rows.Count - 1].Value = Utility.getDepartmentName(_DBName, strCode);
        //    }
        //    else
        //    {
        //        g[ColSznm, g.Rows.Count - 1].Value = string.Empty;
        //    }

        //    g[ColCode, g.Rows.Count - 1].Value = sdR["社員番号"].ToString().PadLeft(6, '0');

        //    // 社員名取得
        //    dbControl.DataControl dCon = new dbControl.DataControl(_DBName);
        //    StringBuilder sb = new StringBuilder();
        //    sb.Clear();
        //    sb.Append("select EmployeeNo,RetireCorpDate,Name from tbEmployeeBase ");
        //    sb.Append("where EmployeeNo = '" + sdR["社員番号"].ToString().PadLeft(10, '0') + "'");

        //    OleDbDataReader dR = dCon.FreeReader(sb.ToString());

        //    while (dR.Read())
        //    {
        //        g[ColName, g.Rows.Count - 1].Value = dR["Name"].ToString();
        //    }

        //    dR.Close();
        //    dCon.Close();

        //    g[ColID, g.Rows.Count - 1].Value = sdR["ID"].ToString();
        //}

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmUnSubmit_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            string rID = string.Empty;

            rID = dataGridView1[ColID, dataGridView1.SelectedRows[0].Index].Value.ToString();

            if (rID != string.Empty)
            {
                if (comboBox1.SelectedIndex == 0)
                {
                    this.Hide();
                    OCR.frmCorrectPast frm = new frmCorrectPast(_DBName, _ComName, rID, true);
                    frm.ShowDialog();
                    this.Show();
                }
                else if (comboBox1.SelectedIndex == 1)
                {
                    this.Hide();
                    OCR.frmOuenCorrectPast frm = new frmOuenCorrectPast(_DBName, _ComName, rID, true);
                    frm.ShowDialog();
                    this.Show();
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Close();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (errCheck())
            {
                if (comboBox1.SelectedIndex == 0)
                {
                    DataSelect(dataGridView1);
                }
                else
                {
                    ouenDataSelect(dataGridView1);
                }
            }
        }
    }
}
