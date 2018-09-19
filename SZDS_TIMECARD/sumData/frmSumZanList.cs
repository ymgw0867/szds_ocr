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
using Excel = Microsoft.Office.Interop.Excel;

namespace SZDS_TIMECARD.sumData
{
    public partial class frmSumZanList : Form
    {
        public frmSumZanList(string dbName)
        {
            InitializeComponent();

            _dbName = dbName;

            hAdp.Fill(dts.過去勤務票ヘッダ);
            dAdp.Fill(dts.休日);
        }

        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.残業集計TableAdapter adp = new DataSet1TableAdapters.残業集計TableAdapter();
        DataSet1TableAdapters.過去勤務票ヘッダTableAdapter hAdp = new DataSet1TableAdapters.過去勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.休日TableAdapter dAdp = new DataSet1TableAdapters.休日TableAdapter();

        xlsData bs;
        string _dbName = string.Empty;

        const string KAKARI_TOTAL = "係合計";
        const string KA_TOTAL = "課合計";
        const string SEIZOU_TOTAL = "製造合計";
        const string BU_TOTAL = "部合計";
        const string KANSETSU_TOTAL = "間接合計";
        const string ALL_TOTAL = "全社合計";

        private void button1_Click(object sender, EventArgs e)
        {
            testSummary();
        }

        private void testSummary()
        {           

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

            // 部署別理由別で残業時間を集計 /////////////////////////////////////////////////////////////////////
            var s = dts.残業集計
                .GroupBy(a => a.部署コード)
                .Select(g => new
                {
                    buCode = g.Key,
                    hhh = g.GroupBy(b => b.残業理由)
                    .Select(h => new
                    {
                        zanRe = h.Key,
                        zanH = h.Sum(a => a.残業時 * 60 + a.残業分)
                    }).OrderBy(a => a.zanRe)
                });

            foreach (var t in s)
            {
                foreach (var i in t.hhh)
                {
                    MessageBox.Show(t.buCode + " " + i.zanRe.ToString() + " " + i.zanH.ToString()); 
                }
            }
        }

        private void frmSumZanList_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            // グリッドビュー定義
            GridviewSet(dataGridView1);

            // 年月初期値
            txtYear.Text = DateTime.Today.Year.ToString();
            txtMonth.Text = DateTime.Today.Month.ToString();

            linkLabel1.Enabled = false;
            linkLabel3.Enabled = false;
        }

        // カラム定義
        private string ColDate = "c0";
        private string ColSz = "c1";
        private string ColSznm = "c2";
        private string ColNin = "c3";
        private string ColKeikaku = "c4";
        private string ColZisseki = "c5";
        private string Col1 = "c6";
        private string Col2 = "c7";
        private string Col3 = "c8";
        private string Col4 = "c9";
        private string Col5 = "c10";
        private string Col6 = "c11";
        private string Col7 = "c12";
        private string Col8 = "c13";
        private string Col9 = "c14";
        private string Col10 = "c15";
        private string ColbyDay = "c16";
        private string ColbyMan = "c17";
        private string ColID = "c18";
        private string ColToZan = "c19";

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
                tempDGV.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                tempDGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

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
                tempDGV.Height = 618;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;

                // 各列幅指定
                tempDGV.Columns.Add(ColSz, "課係班");
                tempDGV.Columns.Add(ColSznm, "部署名");
                tempDGV.Columns.Add(ColNin, "人員");
                tempDGV.Columns.Add(ColKeikaku, "計画");
                tempDGV.Columns.Add(ColZisseki, "実績");
                tempDGV.Columns.Add(Col1, "1");
                tempDGV.Columns.Add(Col2, "2");
                tempDGV.Columns.Add(Col3, "3");
                tempDGV.Columns.Add(Col4, "4");
                tempDGV.Columns.Add(Col5, "5");
                tempDGV.Columns.Add(Col6, "6");
                tempDGV.Columns.Add(Col7, "7");
                tempDGV.Columns.Add(Col8, "8");
                tempDGV.Columns.Add(Col9, "9");
                tempDGV.Columns.Add(Col10, "10");
                tempDGV.Columns.Add(ColbyDay, "日当り");
                tempDGV.Columns.Add(ColbyMan, "1人当り");
                tempDGV.Columns.Add(ColToZan, "当日残業");
                tempDGV.Columns.Add(ColID, "id");

                tempDGV.Columns[ColSz].Width = 72;
                tempDGV.Columns[ColSznm].Width = 200;
                tempDGV.Columns[ColNin].Width = 54;
                tempDGV.Columns[ColKeikaku].Width = 70;
                tempDGV.Columns[ColZisseki].Width = 70;
                tempDGV.Columns[Col1].Width = 60;
                tempDGV.Columns[Col2].Width = 60;
                tempDGV.Columns[Col3].Width = 60;
                tempDGV.Columns[Col4].Width = 60;
                tempDGV.Columns[Col5].Width = 60;
                tempDGV.Columns[Col6].Width = 60;
                tempDGV.Columns[Col7].Width = 60;
                tempDGV.Columns[Col8].Width = 60;
                tempDGV.Columns[Col9].Width = 60;
                tempDGV.Columns[Col10].Width = 60;
                tempDGV.Columns[ColbyDay].Width = 70;
                tempDGV.Columns[ColbyMan].Width = 70;
                tempDGV.Columns[ColToZan].Width = 60;

                tempDGV.Columns[ColID].Visible = false;

                tempDGV.Columns[ColSznm].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[ColSz].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempDGV.Columns[ColNin].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[ColKeikaku].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[ColZisseki].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[Col1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[Col2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[Col3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[Col4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[Col5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[Col6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[Col7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[Col8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[Col9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[Col10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[ColbyDay].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[ColbyMan].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                tempDGV.Columns[ColToZan].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                tempDGV.Columns[ColNin].DefaultCellStyle.Format = "#,##0";
                tempDGV.Columns[ColKeikaku].DefaultCellStyle.Format = "#,##0.0";
                tempDGV.Columns[ColZisseki].DefaultCellStyle.Format = "#,##0.0";
                tempDGV.Columns[Col1].DefaultCellStyle.Format = "#,##0.0";
                tempDGV.Columns[Col2].DefaultCellStyle.Format = "#,##0.0";
                tempDGV.Columns[Col3].DefaultCellStyle.Format = "#,##0.0";
                tempDGV.Columns[Col4].DefaultCellStyle.Format = "#,##0.0";
                tempDGV.Columns[Col5].DefaultCellStyle.Format = "#,##0.0";
                tempDGV.Columns[Col6].DefaultCellStyle.Format = "#,##0.0";
                tempDGV.Columns[Col7].DefaultCellStyle.Format = "#,##0.0";
                tempDGV.Columns[Col8].DefaultCellStyle.Format = "#,##0.0";
                tempDGV.Columns[Col9].DefaultCellStyle.Format = "#,##0.0";
                tempDGV.Columns[Col10].DefaultCellStyle.Format = "#,##0.0";
                tempDGV.Columns[ColbyDay].DefaultCellStyle.Format = "#,##0.0";
                tempDGV.Columns[ColbyMan].DefaultCellStyle.Format = "#,##0.0";
                tempDGV.Columns[ColToZan].DefaultCellStyle.Format = "#,##0.0";

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

                // 罫線
                //tempDGV.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                //tempDGV.CellBorderStyle = DataGridViewCellBorderStyle.None;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        ///--------------------------------------------------------------------------------
        /// <summary>
        ///     グリッドに部署名を表示して表のテンプレートを作成する </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト</param>
        /// <param name="yy">
        ///     対象年</param>
        /// <param name="mm">
        ///     対象月</param>
        ///--------------------------------------------------------------------------------
        private void gridTemp(DataGridView g, int yy, int mm)
        {
            bs = new xlsData();
            bs.zpArray = bs.getZanPlan();
            
            string szCode = string.Empty;
            
            // 奉行データベース接続
            string sc = sqlControl.obcConnectSting.get(_dbName);
            sqlControl.DataControl sdCon = new Common.sqlControl.DataControl(sc);

            g.Rows.Clear();

            try
            {
                for (int i = 1; i <= bs.zpArray.GetLength(0); i++)
                {
                    // 部署コードが5桁未満のときは対象外
                    if (bs.zpArray[i, 1].ToString().Length < 5)
                    {
                        continue;
                    }

                    // 対象年月以外のときは対象外
                    if (Utility.StrtoInt(bs.zpArray[i, 2].ToString()) != yy || Utility.StrtoInt(bs.zpArray[i, 3].ToString()) != mm)
                    {
                        continue;
                    }

                    // 人員数が「０」のときは対象外
                    if (Utility.StrtoInt(bs.zpArray[i, 4].ToString()) == global.flgOff)
                    {
                        continue;
                    }

                    if (szCode != string.Empty)
                    {
                        // 製造部門各合計行
                        if (szCode.Substring(1, 1) == "1")
                        {
                            if (szCode.Substring(1, 1) != bs.zpArray[i, 1].ToString().Substring(1, 1))
                            {
                                g.Rows.Add();
                                g[ColSz, g.Rows.Count - 1].Value = KAKARI_TOTAL;
                                g.Rows[g.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightGray;

                                g.Rows.Add();
                                g[ColSz, g.Rows.Count - 1].Value = KA_TOTAL;
                                g.Rows[g.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightGray;

                                g.Rows.Add();
                                g[ColSz, g.Rows.Count - 1].Value = SEIZOU_TOTAL;
                                g.Rows[g.Rows.Count - 1].DefaultCellStyle.BackColor = Color.SteelBlue;
                                g.Rows[g.Rows.Count - 1].DefaultCellStyle.ForeColor = Color.White;
                            }
                            else if (szCode.Substring(0, 3) != bs.zpArray[i, 1].ToString().Substring(0, 3))
                            {
                                g.Rows.Add();
                                g[ColSz, g.Rows.Count - 1].Value = KAKARI_TOTAL;
                                g.Rows[g.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightGray;

                                g.Rows.Add();
                                g[ColSz, g.Rows.Count - 1].Value = KA_TOTAL;
                                g.Rows[g.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightGray;
                            }
                            else if (szCode.Substring(0, 4) != bs.zpArray[i, 1].ToString().Substring(0, 4))
                            {
                                g.Rows.Add();
                                g[ColSz, g.Rows.Count - 1].Value = KAKARI_TOTAL;
                                g.Rows[g.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightGray;
                            }
                        }
                        else
                        {
                            // 間接部門各合計行
                            if (szCode.Substring(0, 3) != bs.zpArray[i, 1].ToString().Substring(0, 3))
                            {
                                // 間接部門
                                g.Rows.Add();
                                g[ColSz, g.Rows.Count - 1].Value = BU_TOTAL;
                                g.Rows[g.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightGray;
                            }
                        }
                    }

                    g.Rows.Add();
                    g[ColSz, g.Rows.Count - 1].Value = Utility.NulltoStr(bs.zpArray[i, 1]);
                    g[ColSznm, g.Rows.Count - 1].Value = getDepartmentName(_dbName, Utility.NulltoStr(bs.zpArray[i, 1]), sdCon);
                    g[ColNin, g.Rows.Count - 1].Value = Utility.NulltoStr(bs.zpArray[i, 4]);
                    g[ColKeikaku, g.Rows.Count - 1].Value = Utility.StrtoInt(Utility.NulltoStr(bs.zpArray[i, 5])) * Utility.StrtoInt(lblWdays.Text) / Utility.StrtoInt(lblKDays.Text);
                    g[ColZisseki, g.Rows.Count - 1].Value = (double)(0);
                    g[ColbyDay, g.Rows.Count - 1].Value = (double)(0);
                    g[ColbyMan, g.Rows.Count - 1].Value = (double)(0);

                    szCode = bs.zpArray[i, 1].ToString();
                }

                // 終了処理
                g.Rows.Add();
                g[ColSz, g.Rows.Count - 1].Value = BU_TOTAL;
                g.Rows[g.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightGray;

                g.Rows.Add();
                g[ColSz, g.Rows.Count - 1].Value = KANSETSU_TOTAL;
                g.Rows[g.Rows.Count - 1].DefaultCellStyle.BackColor = Color.SteelBlue;
                g.Rows[g.Rows.Count - 1].DefaultCellStyle.ForeColor = Color.White;

                g.Rows.Add();
                g[ColSz, g.Rows.Count - 1].Value = ALL_TOTAL;
                g.Rows[g.Rows.Count - 1].DefaultCellStyle.BackColor = Color.SteelBlue;
                g.Rows[g.Rows.Count - 1].DefaultCellStyle.ForeColor = Color.White;

                g.CurrentCell = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                // 奉行データベース接続解除
                if (sdCon.Cn.State == ConnectionState.Open)
                {
                    sdCon.Close();
                }
            }
        }

        ///------------------------------------------------------------------------------
        /// <summary>
        ///     奉行データベースの部門マスターから部門名を取得する </summary>
        /// <param name="_dbName">
        ///     データベース名</param>
        /// <param name="dCode">
        ///     部署コード</param>
        /// <param name="sdCon">
        ///     sqlControl.DataControl </param>
        /// <returns>
        ///     部署名</returns>
        ///------------------------------------------------------------------------------
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

            //// 接続文字列取得
            //string sc = sqlControl.obcConnectSting.get(_dbName);
            //sqlControl.DataControl sdCon = new Common.sqlControl.DataControl(sc);

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

            //sdCon.Close();

            return dName;
        }

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     各合計行表示 </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト</param>
        ///---------------------------------------------------------------
        private void setSectionTotal(DataGridView g)
        {
            zTotal[] z = new zTotal[4];

            // 集計配列初期化
            for (int i = 0; i < 4; i++)
            {
                z[i] = new zTotal();
                z[i].zNin = 0;
                z[i].zKeikaku = 0;
                z[i].zZisseki = 0;
                z[i].zRe1 = 0;
                z[i].zRe2 = 0;
                z[i].zRe3 = 0;
                z[i].zRe4 = 0;
                z[i].zRe5 = 0;
                z[i].zRe6 = 0;
                z[i].zRe7 = 0;
                z[i].zRe8 = 0;
                z[i].zRe9 = 0;
                z[i].zRe10 = 0;
                z[i].zByDay = 0;
                z[i].zByMan = 0;
                z[i].zToZan = 0;
            }
            
            for (int i = 0; i < g.RowCount; i++)
            {
                if (g[ColSz, i].Value.ToString() == KAKARI_TOTAL)
                {
                    // 製造部門・係合計
                    secKakariTotal(g, i, z);
                }
                else if (g[ColSz, i].Value.ToString() == KA_TOTAL)
                {
                    // 製造部門・課合計
                    secKaTotal(g, i, z);
                }
                else if (g[ColSz, i].Value.ToString() == SEIZOU_TOTAL)
                {
                    // 製造部門合計
                    secSeizouKansetsuTotal(g, i, z);
                }
                else if (g[ColSz, i].Value.ToString() == BU_TOTAL)
                {
                    // 間接部門・部合計
                    secKaTotal(g, i, z);
                }
                else if (g[ColSz, i].Value.ToString() == KANSETSU_TOTAL)
                {
                    // 間接部門合計
                    secSeizouKansetsuTotal(g, i, z);
                }
                else if (g[ColSz, i].Value.ToString() == ALL_TOTAL)
                {
                    // 全社合計
                    secAllTotal(g, i, z);
                }
                else
                {
                    // 各項目の加算
                    for (int iz = 0; iz < 4; iz++)
                    {
                        z[iz].zNin += Utility.StrtoInt(Utility.NulltoStr(g[ColNin, i].Value));
                        z[iz].zKeikaku += Utility.StrtoDouble(Utility.NulltoStr(g[ColKeikaku, i].Value));
                        z[iz].zZisseki += Utility.StrtoDouble(Utility.NulltoStr(g[ColZisseki, i].Value));
                        z[iz].zRe1 += Utility.StrtoDouble(Utility.NulltoStr(g[Col1, i].Value));
                        z[iz].zRe2 += Utility.StrtoDouble(Utility.NulltoStr(g[Col2, i].Value));
                        z[iz].zRe3 += Utility.StrtoDouble(Utility.NulltoStr(g[Col3, i].Value));
                        z[iz].zRe4 += Utility.StrtoDouble(Utility.NulltoStr(g[Col4, i].Value));
                        z[iz].zRe5 += Utility.StrtoDouble(Utility.NulltoStr(g[Col5, i].Value));
                        z[iz].zRe6 += Utility.StrtoDouble(Utility.NulltoStr(g[Col6, i].Value));
                        z[iz].zRe7 += Utility.StrtoDouble(Utility.NulltoStr(g[Col7, i].Value));
                        z[iz].zRe8 += Utility.StrtoDouble(Utility.NulltoStr(g[Col8, i].Value));
                        z[iz].zRe9 += Utility.StrtoDouble(Utility.NulltoStr(g[Col9, i].Value));
                        z[iz].zRe10 += Utility.StrtoDouble(Utility.NulltoStr(g[Col10, i].Value));
                        z[iz].zByDay += Utility.StrtoDouble(Utility.NulltoStr(g[ColbyDay, i].Value));
                        z[iz].zByMan += Utility.StrtoDouble(Utility.NulltoStr(g[ColbyMan, i].Value));
                        z[iz].zToZan += Utility.StrtoDouble(Utility.NulltoStr(g[ColToZan, i].Value));
                    }
                }
            }
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     製造部門・係合計 </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト</param>
        /// <param name="i">
        ///     データグリッドビュー行</param>
        /// <param name="z">
        ///     zTotal（トータルクラス）配列</param>
        ///-------------------------------------------------------------------
        private void secKakariTotal(DataGridView g, int i, zTotal[] z)
        {
            // 製造部門・係合計
            g[ColNin, i].Value = z[0].zNin;
            g[ColKeikaku, i].Value = z[0].zKeikaku;
            g[ColZisseki, i].Value = z[0].zZisseki;
            g[Col1, i].Value = z[0].zRe1;
            g[Col2, i].Value = z[0].zRe2;
            g[Col3, i].Value = z[0].zRe3;
            g[Col4, i].Value = z[0].zRe4;
            g[Col5, i].Value = z[0].zRe5;
            g[Col6, i].Value = z[0].zRe6;
            g[Col7, i].Value = z[0].zRe7;
            g[Col8, i].Value = z[0].zRe8;
            g[Col9, i].Value = z[0].zRe9;
            g[Col10, i].Value = z[0].zRe10;
            g[ColbyDay, i].Value = z[0].zByDay;
            g[ColbyMan, i].Value = z[0].zZisseki / z[0].zNin;
            g[ColToZan, i].Value = z[0].zToZan;

            // 初期化
            z[0].zNin = 0;
            z[0].zKeikaku = 0;
            z[0].zZisseki = 0;
            z[0].zRe1 = 0;
            z[0].zRe2 = 0;
            z[0].zRe3 = 0;
            z[0].zRe4 = 0;
            z[0].zRe5 = 0;
            z[0].zRe6 = 0;
            z[0].zRe7 = 0;
            z[0].zRe8 = 0;
            z[0].zRe9 = 0;
            z[0].zRe10 = 0;
            z[0].zByDay = 0;
            z[0].zByMan = 0;
            z[0].zToZan = 0;
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     製造部門・課合計 </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト</param>
        /// <param name="i">
        ///     データグリッドビュー行</param>
        /// <param name="z">
        ///     zTotal（トータルクラス）配列</param>
        ///-------------------------------------------------------------------
        private void secKaTotal(DataGridView g, int i, zTotal[] z)
        {
            // 製造部門・課合計
            g[ColNin, i].Value = z[1].zNin;
            g[ColKeikaku, i].Value = z[1].zKeikaku;
            g[ColZisseki, i].Value = z[1].zZisseki;
            g[Col1, i].Value = z[1].zRe1;
            g[Col2, i].Value = z[1].zRe2;
            g[Col3, i].Value = z[1].zRe3;
            g[Col4, i].Value = z[1].zRe4;
            g[Col5, i].Value = z[1].zRe5;
            g[Col6, i].Value = z[1].zRe6;
            g[Col7, i].Value = z[1].zRe7;
            g[Col8, i].Value = z[1].zRe8;
            g[Col9, i].Value = z[1].zRe9;
            g[Col10, i].Value = z[1].zRe10;
            g[ColbyDay, i].Value = z[1].zByDay;
            g[ColbyMan, i].Value = z[1].zZisseki / z[1].zNin;
            g[ColToZan, i].Value = z[1].zToZan;

            // 初期化
            z[1].zNin = 0;
            z[1].zKeikaku = 0;
            z[1].zZisseki = 0;
            z[1].zRe1 = 0;
            z[1].zRe2 = 0;
            z[1].zRe3 = 0;
            z[1].zRe4 = 0;
            z[1].zRe5 = 0;
            z[1].zRe6 = 0;
            z[1].zRe7 = 0;
            z[1].zRe8 = 0;
            z[1].zRe9 = 0;
            z[1].zRe10 = 0;
            z[1].zByDay = 0;
            z[1].zByMan = 0;
            z[1].zToZan = 0;
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     製造／間接部門・合計 </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト</param>
        /// <param name="i">
        ///     データグリッドビュー行</param>
        /// <param name="z">
        ///     zTotal（トータルクラス）配列</param>
        ///-------------------------------------------------------------------
        private void secSeizouKansetsuTotal(DataGridView g, int i, zTotal[] z)
        {
            // 製造/間接部門・合計
            g[ColNin, i].Value = z[2].zNin;
            g[ColKeikaku, i].Value = z[2].zKeikaku;
            g[ColZisseki, i].Value = z[2].zZisseki;
            g[Col1, i].Value = z[2].zRe1;
            g[Col2, i].Value = z[2].zRe2;
            g[Col3, i].Value = z[2].zRe3;
            g[Col4, i].Value = z[2].zRe4;
            g[Col5, i].Value = z[2].zRe5;
            g[Col6, i].Value = z[2].zRe6;
            g[Col7, i].Value = z[2].zRe7;
            g[Col8, i].Value = z[2].zRe8;
            g[Col9, i].Value = z[2].zRe9;
            g[Col10, i].Value = z[2].zRe10;
            g[ColbyDay, i].Value = z[2].zByDay;
            g[ColbyMan, i].Value = z[2].zZisseki / z[2].zNin; ;
            g[ColToZan, i].Value = z[2].zToZan;

            // 初期化
            z[2].zNin = 0;
            z[2].zKeikaku = 0;
            z[2].zZisseki = 0;
            z[2].zRe1 = 0;
            z[2].zRe2 = 0;
            z[2].zRe3 = 0;
            z[2].zRe4 = 0;
            z[2].zRe5 = 0;
            z[2].zRe6 = 0;
            z[2].zRe7 = 0;
            z[2].zRe8 = 0;
            z[2].zRe9 = 0;
            z[2].zRe10 = 0;
            z[2].zByDay = 0;
            z[2].zByMan = 0;
            z[2].zToZan = 0;
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     全社合計 </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト</param>
        /// <param name="i">
        ///     データグリッドビュー行</param>
        /// <param name="z">
        ///     zTotal（トータルクラス）配列</param>
        ///-------------------------------------------------------------------
        private void secAllTotal(DataGridView g, int i, zTotal[] z)
        {
            // 全社合計
            g[ColNin, i].Value = z[3].zNin;
            g[ColKeikaku, i].Value = z[3].zKeikaku;
            g[ColZisseki, i].Value = z[3].zZisseki;
            g[Col1, i].Value = z[3].zRe1;
            g[Col2, i].Value = z[3].zRe2;
            g[Col3, i].Value = z[3].zRe3;
            g[Col4, i].Value = z[3].zRe4;
            g[Col5, i].Value = z[3].zRe5;
            g[Col6, i].Value = z[3].zRe6;
            g[Col7, i].Value = z[3].zRe7;
            g[Col8, i].Value = z[3].zRe8;
            g[Col9, i].Value = z[3].zRe9;
            g[Col10, i].Value = z[3].zRe10;
            g[ColbyDay, i].Value = z[3].zByDay;
            g[ColbyMan, i].Value = z[3].zZisseki / z[3].zNin;
            g[ColToZan, i].Value = z[3].zToZan;

            // 初期化
            z[3].zNin = 0;
            z[3].zKeikaku = 0;
            z[3].zZisseki = 0;
            z[3].zRe1 = 0;
            z[3].zRe2 = 0;
            z[3].zRe3 = 0;
            z[3].zRe4 = 0;
            z[3].zRe5 = 0;
            z[3].zRe6 = 0;
            z[3].zRe7 = 0;
            z[3].zRe8 = 0;
            z[3].zRe9 = 0;
            z[3].zRe10 = 0;
            z[3].zByDay = 0;
            z[3].zByMan = 0;
            z[3].zToZan = 0;
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     残業集計データ表示 </summary>
        /// <param name="gr">
        ///     データグリッドビューオブジェクト</param>
        /// <param name="yy">
        ///     対象年</param>
        /// <param name="mm">
        ///     対象月</param>
        ///-------------------------------------------------------------------
        private void showZangyoTotal(DataGridView gr, int yy, int mm)
        {
            // 残業集計データ読みこみ
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

            if (comboBox1.SelectedIndex == 0)
            {
                // 部署別理由別で残業時間を集計 ※社員所属で集計
                var s = dts.残業集計
                    .GroupBy(a => a.部署コード)
                    .Select(g => new
                    {
                        buCode = g.Key,
                        hhh = g.GroupBy(b => b.残業理由)
                        .Select(h => new
                        {
                            zanRe = h.Key,
                            zanH = h.Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10))
                        }).OrderBy(a => a.zanRe)
                    });

                foreach (var t in s)
                {
                    double zanZisseki = 0;　// 実績時間
                    double kaDays = 0;      // 当月稼働日数
                    int r = 0;
                    double over10 = 0;

                    foreach (var i in t.hhh)
                    {
                        for (int rI = 0; rI < gr.RowCount; rI++)
                        {
                            if (gr[ColSz, rI].Value.ToString() == t.buCode)
                            {
                                double zan = i.zanH / 60;

                                if (i.zanRe == 1) gr[Col1, rI].Value = zan;
                                if (i.zanRe == 2) gr[Col2, rI].Value = zan;
                                if (i.zanRe == 3) gr[Col3, rI].Value = zan;
                                if (i.zanRe == 4) gr[Col4, rI].Value = zan;
                                if (i.zanRe == 5) gr[Col5, rI].Value = zan;
                                if (i.zanRe == 6) gr[Col6, rI].Value = zan;
                                if (i.zanRe == 7) gr[Col7, rI].Value = zan;
                                if (i.zanRe == 8) gr[Col8, rI].Value = zan;
                                if (i.zanRe == 9) gr[Col9, rI].Value = zan;
                                if (i.zanRe >= 10)
                                {
                                    over10 += zan;
                                    gr[Col10, rI].Value = over10;
                                }
                                zanZisseki += zan;
                                r = rI;
                            }
                        }
                    }

                    gr[ColZisseki, r].Value = zanZisseki;
                    gr[ColbyDay, r].Value = zanZisseki * Utility.StrtoInt(lblWdays.Text) / Utility.StrtoInt(lblKDays.Text);
                    gr[ColbyMan, r].Value = zanZisseki / Utility.StrtoInt(Utility.NulltoStr(gr[ColNin, r].Value));

                    // 該当部署の最近の出勤簿日付
                    int maxDay = dts.過去勤務票ヘッダ.Where(a => a.部署コード == t.buCode).Max(a => a.日);

                    // 最近の出勤簿日付の残業
                    double ss = dts.残業集計.Where(a => a.部署コード == t.buCode && a.日 == maxDay)
                        .Sum(a => a.残業時 * 60 + (a.残業分 * 60 / 10));

                    gr[ColToZan, r].Value = ss / 60;
                }
            }
            else
            {
                // 部署別理由別で残業時間を集計　※応援先
                var s = dts.残業集計
                    .GroupBy(a => a.応援先)
                    .Select(g => new
                    {
                        buCode = g.Key,
                        hhh = g.GroupBy(b => b.残業理由)
                        .Select(h => new
                        {
                            zanRe = h.Key,
                            zanH = h.Sum(a => (a.残業時 * 60) + (a.残業分 * 60 / 10))
                        }).OrderBy(a => a.zanRe)
                    });

                foreach (var t in s)
                {
                    double zanZisseki = 0;　// 実績時間
                    double kaDays = 0;      // 当月稼働日数
                    int r = 0;
                    double over10 = 0;

                    foreach (var i in t.hhh)
                    {
                        for (int rI = 0; rI < gr.RowCount; rI++)
                        {
                            if (gr[ColSz, rI].Value.ToString() == t.buCode)
                            {
                                double zan = i.zanH / 60;

                                if (i.zanRe == 1) gr[Col1, rI].Value = zan;
                                if (i.zanRe == 2) gr[Col2, rI].Value = zan;
                                if (i.zanRe == 3) gr[Col3, rI].Value = zan;
                                if (i.zanRe == 4) gr[Col4, rI].Value = zan;
                                if (i.zanRe == 5) gr[Col5, rI].Value = zan;
                                if (i.zanRe == 6) gr[Col6, rI].Value = zan;
                                if (i.zanRe == 7) gr[Col7, rI].Value = zan;
                                if (i.zanRe == 8) gr[Col8, rI].Value = zan;
                                if (i.zanRe == 9) gr[Col9, rI].Value = zan;
                                if (i.zanRe >= 10)
                                {
                                    over10 += zan;
                                    gr[Col10, rI].Value = over10;
                                }
                                zanZisseki += zan;
                                r = rI;
                            }
                        }
                    }

                    gr[ColZisseki, r].Value = zanZisseki;
                    gr[ColbyDay, r].Value = zanZisseki * Utility.StrtoInt(lblWdays.Text) / Utility.StrtoInt(lblKDays.Text);
                    gr[ColbyMan, r].Value = zanZisseki / Utility.StrtoInt(Utility.NulltoStr(gr[ColNin, r].Value));

                    // 該当部署の最近の出勤簿日付
                    int maxDay = dts.過去勤務票ヘッダ.Where(a => a.部署コード == t.buCode).Max(a => a.日);

                    // 最近の出勤簿日付の残業
                    double ss = dts.残業集計.Where(a => a.部署コード == t.buCode && a.日 == maxDay)
                        .Sum(a => a.残業時 * 60 + (a.残業分 * 60 / 10));

                    gr[ColToZan, r].Value = ss / 60;
                }
            }
        }

        private void txtYear_TextChanged(object sender, EventArgs e)
        {
            yymmChanged();
        }

        private void yymmChanged()
        {
            DateTime dt;
            string str = txtYear.Text + "/" + txtMonth.Text + "/1";
            if (!DateTime.TryParse(str, out dt))
            {
                return;
            }

            if (dts.過去勤務票ヘッダ.Any(a => a.年 == Utility.StrtoInt(txtYear.Text) &&
                                                   a.月 == Utility.StrtoInt(txtMonth.Text)))
            {
                // 最新の出勤簿日付を取得・表示
                var s = dts.過去勤務票ヘッダ.Where(a => a.年 == Utility.StrtoInt(txtYear.Text) &&
                                                       a.月 == Utility.StrtoInt(txtMonth.Text))
                                           .Max(a => a.日);

                lblKDays.Text = getKadouDays(Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text)).ToString(); 
                lblDate.Text = txtYear.Text + "/" + txtMonth.Text.PadLeft(2, '0') + "/" + s.ToString().PadLeft(2, '0');
                lblWdays.Text = getWorkDays(DateTime.Parse(lblDate.Text)).ToString();
                comboBox1.Enabled = true;
                comboBox1.SelectedIndex = 0;
                linkLabel1.Enabled = true;
            }
            else
            {
                lblKDays.Text = "--";
                lblDate.Text = "出勤簿なし";
                lblWdays.Text = "--";
                comboBox1.Enabled = false;
                comboBox1.SelectedIndex = -1;
                linkLabel1.Enabled = false;
            }
        }

        private void txtMonth_TextChanged(object sender, EventArgs e)
        {
            yymmChanged();
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     実働日数を求める </summary>
        /// <param name="dt">
        ///     いつまで（日付）</param>
        /// <returns>
        ///     実働日数</returns>
        ///-------------------------------------------------------------
        private int getWorkDays(DateTime dt)
        {
            int rtn = 0;

            //　該当月の該当日までの休日を取得
            int s = dts.休日.Count(a => a.年月日.Year == dt.Year && a.年月日.Month == dt.Month && a.年月日 <= dt);

            // 実働日数
            rtn = dt.Day - s;

            return rtn;
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     稼働日数を求める </summary>
        /// <param name="yy">
        ///     対象年</param>
        /// <param name="mm">
        ///     対象月</param>
        /// <returns>
        ///     稼働日数</returns>
        ///----------------------------------------------------------------
        private int getKadouDays(int yy, int mm)
        {
            int rtn = 0;

            //　該当月の該当日までの休日を取得
            int s = dts.休日.Count(a => a.年月日.Year == yy && a.年月日.Month == mm);

            DateTime dt = new DateTime(yy, mm, 1);
            dt = dt.AddMonths(1);
            dt = dt.AddDays(-1);

            // 稼働日数
            rtn = dt.Day - s;

            return rtn;
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private class zTotal
        {
            public int zNin { get; set; }
            public double zKeikaku { get; set; }
            public double zZisseki { get; set; }
            public double zRe1 { get; set; }
            public double zRe2 { get; set; }
            public double zRe3 { get; set; }
            public double zRe4 { get; set; }
            public double zRe5 { get; set; }
            public double zRe6 { get; set; }
            public double zRe7 { get; set; }
            public double zRe8 { get; set; }
            public double zRe9 { get; set; }
            public double zRe10 { get; set; }
            public double zByDay { get; set; }
            public double zByMan { get; set; }
            public double zToZan { get; set; }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Close();
        }

        private void frmSumZanList_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            Dispose();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
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

            this.Cursor = Cursors.WaitCursor;
            gridTemp(dataGridView1, Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text));
            showZangyoTotal(dataGridView1, Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text));
            setSectionTotal(dataGridView1);
            this.Cursor = Cursors.Default;

            if (dataGridView1.RowCount > 0)
            {
                linkLabel3.Enabled = true;
            }
            else
            {
                linkLabel3.Enabled = false;
            }
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            setExcelData();
        }


        ///---------------------------------------------------------------------
        /// <summary>
        ///     常陽コンピュータサービス向けエクセル給与シート出力を実行する </summary>
        /// <param name="gPath">
        ///     グループフォルダパス</param>
        ///---------------------------------------------------------------------
        private void setExcelData()
        {
            
            // 表示データを多次元配列へセット
            string[,] mArray = null;
            dgvToMlArray(ref mArray, dataGridView1);

            // エクセルシート出力
            string gXlsPath = Properties.Settings.Default.xlsZanRep;  // エクセル給与シートパス
            saveExcelZangyo(mArray, gXlsPath);
        }
        
        ///-------------------------------------------------------------
        /// <summary>
        ///     残業集計データグリッドビューを多次元配列に展開する </summary>
        /// <param name="mAry">
        ///     多次元配列</param>
        /// <param name="g">
        ///     datagridviewオブジェクト</param>
        ///-------------------------------------------------------------
        public void dgvToMlArray(ref string[,] mAry, DataGridView g)
        {
            int r = g.RowCount;
            int c = g.ColumnCount;

            mAry = new string[r, c];

            int rX = 0;

            for (int i = 0; i < g.RowCount; i++)
            {
                mAry[rX, 0] = Utility.NulltoStr(g[ColSz, i].Value);
                mAry[rX, 1] = Utility.NulltoStr(g[ColSznm, i].Value);
                mAry[rX, 2] = Utility.NulltoStr(g[ColNin, i].Value);
                mAry[rX, 3] = Utility.StrtoDouble(Utility.NulltoStr(g[ColKeikaku, i].Value)).ToString("#,##0.0");
                mAry[rX, 4] = Utility.StrtoDouble(Utility.NulltoStr(g[ColZisseki, i].Value)).ToString("#,##0.0");
                mAry[rX, 5] = getNotZeroValue(Utility.StrtoDouble(Utility.NulltoStr(g[Col1, i].Value)).ToString("#,##0.0"));
                mAry[rX, 6] = getNotZeroValue(Utility.StrtoDouble(Utility.NulltoStr(g[Col2, i].Value)).ToString("#,##0.0"));
                mAry[rX, 7] = getNotZeroValue(Utility.StrtoDouble(Utility.NulltoStr(g[Col3, i].Value)).ToString("#,##0.0"));
                mAry[rX, 8] = getNotZeroValue(Utility.StrtoDouble(Utility.NulltoStr(g[Col4, i].Value)).ToString("#,##0.0"));
                mAry[rX, 9] = getNotZeroValue(Utility.StrtoDouble(Utility.NulltoStr(g[Col5, i].Value)).ToString("#,##0.0"));
                mAry[rX, 10] = getNotZeroValue(Utility.StrtoDouble(Utility.NulltoStr(g[Col6, i].Value)).ToString("#,##0.0"));
                mAry[rX, 11] = getNotZeroValue(Utility.StrtoDouble(Utility.NulltoStr(g[Col7, i].Value)).ToString("#,##0.0"));
                mAry[rX, 12] = getNotZeroValue(Utility.StrtoDouble(Utility.NulltoStr(g[Col8, i].Value)).ToString("#,##0.0"));
                mAry[rX, 13] = getNotZeroValue(Utility.StrtoDouble(Utility.NulltoStr(g[Col9, i].Value)).ToString("#,##0.0"));
                mAry[rX, 14] = getNotZeroValue(Utility.StrtoDouble(Utility.NulltoStr(g[Col10, i].Value)).ToString("#,##0.0"));
                mAry[rX, 15] = Utility.StrtoDouble(Utility.NulltoStr(g[ColbyDay, i].Value)).ToString("#,##0.0");
                mAry[rX, 16] = Utility.StrtoDouble(Utility.NulltoStr(g[ColbyMan, i].Value)).ToString("#,##0.0");
                mAry[rX, 17] = Utility.StrtoDouble(Utility.NulltoStr(g[ColToZan, i].Value)).ToString("#,##0.0");

                rX++;
            }           
        }

        private string getNotZeroValue(string val)
        {
            string rtn = "";

            if (val != "0.0")
            {
                rtn = val;
            }

            return rtn;
        }
        
        ///------------------------------------------------------------------
        /// <summary>
        ///     残業集計表エクセル出力 </summary>
        /// <param name="xls">
        ///     エクセルシート</param>
        /// <param name="xlsFile">
        ///     残業集計表エクセルシート</param>
        ///------------------------------------------------------------------
        public void saveExcelZangyo(string[,] xls, string xlsFile)
        {
            Excel.Application oXls = new Excel.Application();
            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(xlsFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];
            oxlsSheet.Select(Type.Missing);

            Excel.Range _rng = null;

            Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

            const int sGYO = 4;       //エクセルファイル明細開始行

            try
            {
                Cursor = Cursors.WaitCursor;

                // ウィンドウを非表示にする
                oXls.Visible = false;
                oXls.DisplayAlerts = false;

                //// 前回の書き込みセルを初期化する
                //_rng = oxlsSheet.Range[oxlsSheet.Cells[sGYO, 1], oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 18]];
                //_rng.Value2 = "";
                
                //// シートを追加する
                //oxlsSheet.Copy(Type.Missing, oXlsBook.Sheets[1]);
                //oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[2];

                // 残業集計データ書き込み
                _rng = oxlsSheet.Range[oxlsSheet.Cells[sGYO, 1], oxlsSheet.Cells[(sGYO + xls.GetLength(0) - 1), 18]];
                _rng.Value2 = xls;
                
                // ヘッダ
                DateTime dt = DateTime.Parse(lblDate.Text);
                oxlsSheet.Cells[1, 1] = "集計年月： " + txtYear.Text + "年" + txtMonth.Text + "月";
                oxlsSheet.Cells[1, 3] = dt.Day.ToString() + "日まで";
                oxlsSheet.Cells[1, 14] = "稼働日数 " + lblKDays.Text + "日  　　　集計日数 " + lblWdays.Text + "日";
                
                int endRow = 3 + dataGridView1.RowCount;

                // 罫線
                for (int i = 4; i <= endRow; i++)
                {
                    _rng = (Excel.Range)oxlsSheet.Cells[i, 1];
                    rng[0] = (Excel.Range)oxlsSheet.Cells[i, 1];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[i, 18];

                    if (_rng.Value2.ToString() != KAKARI_TOTAL && _rng.Value2.ToString() != KA_TOTAL && _rng.Value2.ToString() != SEIZOU_TOTAL &&
                        _rng.Value2.ToString() != BU_TOTAL && _rng.Value2.ToString() != KANSETSU_TOTAL && _rng.Value2.ToString() != ALL_TOTAL)
                    {
                        //セル下部へドットヨコ罫線を引く
                        oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDot;                       
                    }
                    else
                    { //セル上部へ実線ヨコ罫線を引く
                        oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        //セル下部へ実線ヨコ罫線を引く
                        oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        oxlsSheet.get_Range(rng[0], rng[1]).Interior.Color = Color.LightGray;
                    }
                }
                
                //セル下部へ実線ヨコ罫線を引く
                rng[0] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 1];
                rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 18];
                oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                
                rng[0] = (Excel.Range)oxlsSheet.Cells[4, 1];

                //表全体に実線縦罫線を引く
                oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                //表全体の左端縦罫線
                oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                //表全体の右端縦罫線
                oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                for (int i = 7; i <= 15; i++)
                {                    
                    rng[0] = (Excel.Range)oxlsSheet.Cells[3, i];
                    rng[1] = (Excel.Range)oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, i];
                    oxlsSheet.get_Range(rng[0], rng[1]).Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDot;
                }

                Cursor = Cursors.Default;

                // ウィンドウを表示にする
                oXls.Visible = true;
                oxlsSheet.PrintPreview();
                //oxlsSheet.PrintOut(1, Type.Missing, 1, false, oXls.ActivePrinter, Type.Missing, Type.Missing, Type.Missing);

                //ダイアログボックスの初期設定
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Title = "残業集計表";
                saveFileDialog1.OverwritePrompt = true;
                saveFileDialog1.RestoreDirectory = true;
                saveFileDialog1.FileName = "残業集計表_" + txtYear.Text + txtMonth.Text.PadLeft(2, '0');
                saveFileDialog1.Filter = "Microsoft Office Excelファイル(*.xlsx)|*.xlsx|全てのファイル(*.*)|*.*";

                //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
                string fileName;
                DialogResult ret = saveFileDialog1.ShowDialog();

                if (ret == System.Windows.Forms.DialogResult.OK)
                {
                    // エクセル保存
                    fileName = saveFileDialog1.FileName;
                    oXlsBook.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "残業集計表エクセル出力", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }
            finally
            {
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

    }
}
