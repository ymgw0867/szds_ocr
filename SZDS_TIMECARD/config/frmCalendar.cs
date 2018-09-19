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
using SZDS_TIMECARD.config;

namespace SZDS_TIMECARD.config
{
    public partial class frmCalendar : Form
    {
        public frmCalendar()
        {
            InitializeComponent();

            adp.Fill(dts.休日);
        }

        private void frmCalendar_Load(object sender, EventArgs e)
        {
            Utility.WindowsMaxSize(this, this.Width, this.Height);
            Utility.WindowsMinSize(this, this.Width, this.Height);

            GridViewSetting(dataGridView1); // グリッドビュー設定
            ComboYear();                    // 対象年コンボボックス値セット
            GridViewShow(dataGridView1);    // グリッドビュー表示
            DispClr();                      // 画面初期化

            // 休日コンボボックス値セット
            Utility.comboHoliday.Load(comboBox1);
        }

        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.休日TableAdapter adp = new DataSet1TableAdapters.休日TableAdapter();
        

        // ID
        string _ID;

        // 登録モード
        int _fMode = 0;

        // グリッドビューカラム名
        private string cDate = "c1";
        private string cGekkyu = "c2";
        private string cJikyu = "c3";
        private string cMemo = "c4";
        private string cID = "c5";

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
                dg.Height = 322;

                // 全体の幅
                //dg.Width = 583;

                // 奇数行の色
                //dg.AlternatingRowsDefaultCellStyle.BackColor = Color.LightBlue;

                //各列幅指定
                dg.Columns.Add(cDate, "年月日");
                dg.Columns.Add(cMemo, "名称");
                dg.Columns.Add(cID, "ID");
                dg.Columns[cID].Visible = false;

                dg.Columns[cDate].Width = 110;
                dg.Columns[cMemo].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dg.Columns[cDate].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
              
                // 行ヘッダを表示しない
                dg.RowHeadersVisible = false;

                // 選択モード
                dg.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                dg.MultiSelect = true;

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

                // ソート禁止
                foreach (DataGridViewColumn c in dg.Columns)
                {
                    c.SortMode = DataGridViewColumnSortMode.NotSortable;
                }
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

        /// <summary>
        /// 休日対象年
        /// </summary>
        private void ComboYear()
        {
            comboBox2.Items.Clear();

            var s = dts.休日.Select(a => new
            {
                y = a.年月日.Year,
            }).Distinct();
            
            foreach (var t in s)
            {
                comboBox2.Items.Add(t.y.ToString());
            }
            
            // 今年を初期表示とする
            comboBox2.SelectedIndex = -1;
            for (int i = 0; i < comboBox2.Items.Count; i++)
            {
                if (comboBox2.Items[i].ToString() == DateTime.Today.Year.ToString())
                {
                    comboBox2.SelectedIndex = i;
                    break;
                }
            }

            // 当年の休日が登録されていないときは一番最近の年を初期表示とする
            if (comboBox2.Items.Count != 0)
            {
                if (comboBox2.SelectedIndex == -1) comboBox2.SelectedIndex = comboBox2.Items.Count - 1;
            }
        }

        ///------------------------------------------------------------------------
        /// <summary>
        ///     休日データをグリッドビューへ表示します </summary>
        /// <param name="tempGrid">
        ///     データグリッドビューオブジェクト</param>
        ///------------------------------------------------------------------------
        private void GridViewShow(DataGridView tempGrid)
        {
            if (comboBox2.Text != string.Empty)
            {
                int iX = 0;
                tempGrid.RowCount = 0;

                foreach (var t in dts.休日.OrderBy(a => a.年月日))
                {
                    if (t.年月日.Year.ToString() != comboBox2.Text)
                    {
                        continue;
                    }

                    tempGrid.Rows.Add();

                    tempGrid[cDate, iX].Value = DateTime.Parse(t.年月日.ToString()).ToShortDateString();
                    tempGrid[cMemo, iX].Value = t.名称;
                    tempGrid[cID, iX].Value = t.ID.ToString();
                    iX++;
                }

                tempGrid.CurrentCell = null;
            }
        }

        private void DispClr()
        {
            txtDate.Text = string.Empty;
            comboBox1.Text = string.Empty;

            lnkLblUpdate.Enabled = false;
            lnkLblDelete.Enabled = false;
            lnkLblClr.Enabled = false;
            monthCalendar1.Enabled = true;

            _fMode = 0;
        }

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            txtDate.Text = monthCalendar1.SelectionStart.ToString("yyyy/MM/dd (ddd)");

            // 休日名称を表示
            string md = monthCalendar1.SelectionStart.ToString("MM/dd");
            Utility.comboHoliday.selectedIndex(comboBox1, md);
            
            lnkLblUpdate.Enabled = true;
            lnkLblClr.Enabled = true;
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     休日テーブルに休日データを新規に登録する </summary>
        /// <param name="dt">
        ///     対象となる日付</param>
        ///----------------------------------------------------------------
        private void dataInsert(DateTime dt)
        {
            DataSet1.休日Row r = dts.休日.New休日Row();
            r.年月日 = DateTime.Parse(dt.ToShortDateString());
            r.名称 = comboBox1.Text;
            r.備考 = string.Empty;
            r.更新年月日 = DateTime.Now;
            dts.休日.Add休日Row(r);
            adp.Update(dts.休日);
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     休日データを更新する </summary>
        /// <param name="dt">
        ///     対象となる日付</param>
        ///----------------------------------------------------------------
        private void dataUpdate(DateTime dt)
        {
            if (dts.休日.Any(a => a.ID == int.Parse(_ID)))
            {
                var s = dts.休日.Single(a => a.ID == int.Parse(_ID));
                s.年月日 = DateTime.Parse(txtDate.Text.Substring(0, 10));
                s.名称 = comboBox1.Text;
                s.備考 = string.Empty;
                s.更新年月日 = DateTime.Now;
                adp.Update(dts.休日);
            }
            else
            {
                MessageBox.Show("休日カレンダーの更新に失敗しました。","更新エラー",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
            }
        }

        ///--------------------------------------------------------------
        /// <summary>
        ///     休日データを削除する </summary>
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
                var d = dts.休日.Where(a => a.RowState != DataRowState.Deleted && a.RowState != DataRowState.Detached && a.ID == sID);

                // foreach用の配列を作成する
                var list = d.ToList();

                // 削除
                foreach (var it in list)
                {
                    DataSet1.休日Row dl = dts.休日.FindByID(it.ID);
                    dl.Delete();
                }

                return true;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString(), "削除失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     グリッドビューの選択された行データを表示する </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト</param>
        ///---------------------------------------------------------------------
        private void GetGridViewData(DataGridView g)
        {
            if (g.SelectedRows.Count == 0) return;

            int r = g.SelectedRows[0].Index;

            string y = g[cDate, r].Value.ToString();

            txtDate.Text = DateTime.Parse(y).ToString("yyyy/MM/dd (ddd)");
            comboBox1.Text = g[cMemo, r].Value.ToString();

            _ID = g[cID, r].Value.ToString();

            lnkLblUpdate.Enabled = true;
            lnkLblDelete.Enabled = true;
            lnkLblClr.Enabled = true;
            //monthCalendar1.Enabled = false;
            _fMode = 1;
        }

        private void btnClr_Click(object sender, EventArgs e)
        {
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     データ削除 </summary>
        /// <param name="sender">
        ///     </param>
        /// <param name="e">
        ///     </param>
        ///---------------------------------------------------------------------
        private void btnDelete_Click(object sender, EventArgs e)
        {
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     休日データを検索する </summary>
        /// <param name="dt">
        ///     対象となる日付</param>
        /// <returns>
        ///     true:データなし、false:データあり</returns>
        ///---------------------------------------------------------------------
        private bool dataSearch(DateTime dt)
        {
            string s2 = dt.ToShortDateString();

            if (dts.休日.Any(a => a.年月日.ToShortDateString() == s2))
            {
                return false;
            }
                          
            return true;
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
        }

        private void frmCalendar_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            GetGridViewData(dataGridView1);
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            GridViewShow(dataGridView1);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            frmCalenderBatch frm = new config.frmCalenderBatch();
            frm.ShowDialog();
            adp.Fill(dts.休日);
            GridViewShow(dataGridView1);    // グリッドビュー表示
        }

        private void linkLabel4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            DispClr();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("削除する休日を選択してください", "休日未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show(dataGridView1.SelectedRows.Count.ToString() + "件の休日を削除します。よろしいですか？", "休日削除", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;

            int dC = 0;

            for (int i = 0; i < dataGridView1.SelectedRows.Count; i++)
            {
                int r = dataGridView1.SelectedRows[i].Index;
                int sID = int.Parse(dataGridView1[cID, r].Value.ToString());

                if (dataDelete(sID))
                {
                    dC++;
                }
            }

            MessageBox.Show(dC.ToString() + "件の休日を削除しました", "結果", MessageBoxButtons.OK, MessageBoxIcon.Information);

            adp.Update(dts.休日);
            adp.Fill(dts.休日);
            ComboYear();
            //GridViewShow(dataGridView1);
            DispClr();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (txtDate.Text == string.Empty)
            {
                MessageBox.Show("日付が選択されていません", "休日設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            switch (_fMode)
            {
                case 0:
                    if (dataSearch(monthCalendar1.SelectionStart))
                    {
                        if (MessageBox.Show(txtDate.Text + " を登録しますか？", "休日登録", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;
                        dataInsert(monthCalendar1.SelectionStart);
                    }
                    else
                    {
                        MessageBox.Show("既に登録済みの日付です", "休日設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    break;
                case 1:
                    if (MessageBox.Show(txtDate.Text + " を更新しますか？", "休日登録", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;
                    dataUpdate(DateTime.Parse(txtDate.Text));
                    break;
                default:
                    break;
            }

            adp.Fill(dts.休日);
            ComboYear();
            GridViewShow(dataGridView1);
            DispClr();
        }
    }
}
