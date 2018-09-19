using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SZDS_TIMECARD.Common;
using System.Data.SqlClient;

namespace SZDS_TIMECARD.OCR
{
    public partial class frmOCRIndex : Form
    {
        public frmOCRIndex(string dbName, DataSet1 _dts, DataSet1TableAdapters.勤務票ヘッダTableAdapter _hAdp, DataSet1TableAdapters.勤務票明細TableAdapter _mAdp)
        {
            InitializeComponent();

            dts = _dts;
            hAdp = _hAdp;
            mAdp = _mAdp;

            //hAdp.Fill(dts.勤務票ヘッダ);
            //mAdp.Fill(dts.勤務票明細);

            _dbName = dbName;
        }

        DataSet1 dts = null;
        DataSet1TableAdapters.勤務票ヘッダTableAdapter hAdp = new DataSet1TableAdapters.勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.勤務票明細TableAdapter mAdp = new DataSet1TableAdapters.勤務票明細TableAdapter();

        string _dbName = string.Empty;

        private void frmOCRIndex_Load(object sender, EventArgs e)
        {
            Utility.WindowsMaxSize(this, this.Width, this.Height);
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // 部署コンボボックスロード
            Utility.ComboBumon.loadBusho(comboBox1, _dbName);

            dateTimePicker1.Checked = false;

            // データグリッドビュー定義
            GridViewSetting(dataGridView1);

            // データグリッドビュー表示
            GridViewShowData(dataGridView1, 0, 0, 0, string.Empty);
        }


        #region グリッドカラム定義
        string colDate = "c1";
        string colBushoCode = "c2";
        string colBushoName = "c3";
        string colNinzu = "c4";
        string colID = "c5";
        #endregion

        /// ------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        /// ------------------------------------------------------------------
        public void GridViewSetting(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更するe

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 11, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("Meiryo UI", (float)11, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 22;
                tempDGV.RowTemplate.Height = 22;

                // 全体の高さ
                tempDGV.Height = 554;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = false;

                // カラム定義
                tempDGV.Columns.Add(colDate, "日付");
                tempDGV.Columns.Add(colBushoCode, "部署コード");
                tempDGV.Columns.Add(colBushoName, "部署名");
                tempDGV.Columns.Add(colNinzu, "人数");
                tempDGV.Columns.Add(colID, "ID");

                // IDは非表示
                tempDGV.Columns[colID].Visible = false;

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

                // 表示位置
                tempDGV.Columns[colDate].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colBushoCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[colNinzu].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                              
                // 各列幅指定
                tempDGV.Columns[colDate].Width = 110;
                tempDGV.Columns[colBushoCode].Width = 100;
                tempDGV.Columns[colNinzu].Width = 80;
                tempDGV.Columns[colBushoName].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                // 編集可否
                tempDGV.ReadOnly = true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GridViewShowData(DataGridView g, int sYY, int sMM, int sDD, string sBushoCode)
        {
            try
            {
                g.Rows.Clear();
                int i = 0;

                var s = dts.勤務票ヘッダ.OrderBy(a => a.年).ThenBy(a => a.月).ThenBy(a => a.日).ThenBy(a => a.部署コード);

                // 日付検索
                if (sYY != 0)
                {
                    s = s.Where(a => a.年 == sYY && a.月 == sMM && a.日 == sDD).OrderBy(a => a.年).ThenBy(a => a.月).ThenBy(a => a.日).ThenBy(a => a.部署コード);
                }

                // 部署検索
                if (sBushoCode != string.Empty)
                {
                    s = s.Where(a => a.部署コード == sBushoCode).OrderBy(a => a.年).ThenBy(a => a.月).ThenBy(a => a.日).ThenBy(a => a.部署コード);
                }

                foreach (DataSet1.勤務票ヘッダRow t in s)
                {
                    g.Rows.Add();
                    g[colDate, i].Value = t.年.ToString() + "/" + t.月.ToString() + "/" + t.日.ToString();
                    g[colBushoCode, i].Value = t.部署コード;
                    
                    // 奉行データベースより部署名を取得して表示します
                    if (Utility.NulltoStr(t.部署コード) != string.Empty)
                    {
                        string dName = string.Empty;
                        if (getDepartMentName(out dName, t.部署コード.ToString()))
                        {
                            g[colBushoName, i].Value = dName;
                        }
                    }

                    int nin = 0;
                    foreach (var m in t.Get勤務票明細Rows())
	                {
                        if (m.社員番号 != string.Empty)
                        {
                            nin++;
                        }
	                }

                    g[colNinzu, i].Value = nin.ToString();

                    g[colID, i].Value = t.ID.ToString();

                    i++;
                }

                g.CurrentCell = null;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            }

            // 勤務票情報がないとき
            if (g.RowCount == 0)
            {
                MessageBox.Show("勤怠データI/P票が存在しませんでした", "データなし", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            // ID初期化
            hdID = string.Empty;

            // 確認
            string msg = dataGridView1[colDate, e.RowIndex].Value.ToString() + " " + dataGridView1[colBushoName, e.RowIndex].Value.ToString() + " が選択されました。よろしいですか？";
            if (MessageBox.Show(msg,"確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            // ID取得
            hdID = dataGridView1[colID, e.RowIndex].Value.ToString();

            // 閉じる
            this.Close();
        }

        // 選択したデータのID
        public string hdID { get; set; }

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
        private bool getDepartMentName(out string dName, string dCode)
        {
            bool rtn = false;
            int c = 0;

            // 部署名を初期化
            dName = string.Empty;

            // 奉行データベースより部署名を取得して表示します
            if (Utility.NulltoStr(dCode) != string.Empty)
            {
                string b = string.Empty;

                // 検索用部署コード
                if (Utility.StrtoInt(dCode) != global.flgOff)
                {
                    b = dCode.Trim().PadLeft(15, '0');
                }
                else
                {
                    b = dCode.Trim().PadRight(15, ' ');
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

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmOCRIndex_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            getSelectData();
        }

        private void getSelectData()
        {
            int sYY = 0;
            int sMM = 0;
            int sDD = 0;
            string sBusho = string.Empty;
            
            if (dateTimePicker1.Checked)
            {
                sYY = dateTimePicker1.Value.Year;
                sMM = dateTimePicker1.Value.Month;
                sDD = dateTimePicker1.Value.Day;
            }

            if (comboBox1.SelectedIndex != -1)
            {
                Utility.ComboBumon cmb = (Utility.ComboBumon)comboBox1.SelectedItem;
                sBusho = cmb.code; 
            }

            // データ検索
            GridViewShowData(dataGridView1, sYY, sMM, sDD, sBusho);
        }
    }
}
