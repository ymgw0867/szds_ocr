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
    public partial class frmKintaiRepNew : Form
    {
        public frmKintaiRepNew(string dbName)
        {
            InitializeComponent();

            _dbName = dbName;
        }

        string _dbName = string.Empty;

        DataSet1 dts = new DataSet1();
        //DataSet1TableAdapters.過去勤務票ヘッダTableAdapter hAdp = new DataSet1TableAdapters.過去勤務票ヘッダTableAdapter();
        //DataSet1TableAdapters.過去勤務票明細TableAdapter iAdp = new DataSet1TableAdapters.過去勤務票明細TableAdapter();
        DataSet1TableAdapters.休日TableAdapter dAdp = new DataSet1TableAdapters.休日TableAdapter();
        //DataSet1TableAdapters.残業集計TableAdapter zAdp = new DataSet1TableAdapters.残業集計TableAdapter();
        //DataSet1TableAdapters.帰宅後勤務TableAdapter kAdp = new DataSet1TableAdapters.帰宅後勤務TableAdapter();

        // 雇用区分
        const int KBN_SHAIN = 1;
        const int KBN_PART = 3;     // パート社員
        const int KBN_PART_2 = 8;   // 契約社員（パート扱い）
        const int KBN_PART_3 = 9;   // アルバイト60（パート扱い）

        // 事由コード
        const string JIYU_NENKYU = "1";
        const string JIYU_ZENHANKYU = "2";
        const string JIYU_KOUHANKYU = "3";
        const string JIYU_TSUMIKYU = "4";
        const string JIYU_TSUMIZENHAN = "5";
        const string JIYU_TSUMIKOUHAN = "6";
        const string JIYU_YUUKOUKYU = "7";
        const string JIYU_YUUKOUZENHAN = "8";
        const string JIYU_YUUKOUKOUHAN = "9";
        const string JIYU_KEKKIN = "10";
        const string JIYU_DAIKYU = "12";
        const string JIYU_FURIKYU = "13";
        const string JIYU_YOBIDASHI = "30";
        const string JIYU_DOYOTOKKYU = "40";    // 2017/11/21
        const string JIYU_KEKKINZENHAN = "17";  // 2017/11/21
        const string JIYU_KEKKINKOUHAN = "18";  // 2017/11/21

        // 勤務体系コード
        const string SHIFT_KYUSHUTSU = "031";
        const string SHIFT_KYUKEI_KYUSHUTSU = "032";    // 2018/02/05
        const string SHIFT_HEIKITAKUGO = "041";
        const string SHIFT_KYUKITAKUGO = "042";

        // 合計欄変数
        double shukkinDays = 0;
        double kyushutsuDays = 0;
        double yuukouDays = 0;
        double nenkyuuDays = 0;
        double tsumikyuuDays = 0;       // 積休 2017/10/04
        double kekkinDays = 0;
        double workTime = 0;
        double zanTime = 0;
        double shinyaTime = 0;
        double kyushutsuTime = 0;
        double kyusuhtsuShinyaTime = 0;
        double koukiTime = 0;
        double kumitateTime = 0;
        double yobidashi = 0;
        double chisouKai = 0;           // 遅刻早退回数 2017/09/22
        double chisouTime = 0;          // 遅刻早退時間 2017/09/22
        double chisouTimeKyuGyo = 0;    // 休業遅刻早退時間 2017/09/23

        // カラム定義
        private string ColChk = "c0";
        private string ColSz = "c1";
        private string ColSznm = "c2";
        private string ColCode = "c3";
        private string ColNin = "c4";
        private string ColID = "c5";

        // 奉行から出力した勤怠データ配列
        string[] workArray = null;

        // 奉行から出力した年休・積休残データ配列
        string[] nenkyuArray = null;

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            main();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Close();
        }

        private void frmKintaiRep_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            //// 部署名コンボボックスのデータソースをセットする
            //Utility.ComboBumon.loadBusho(comboBox2, _dbName);

            // 部署一覧グリッド定義
            gridViewSet(dataGridView1);

            // 部門一覧表示
            //departmentShow();

            // 年月初期値
            txtYear.Text = DateTime.Today.Year.ToString();
            txtMonth.Text = DateTime.Today.Month.ToString();

            linkLabel3.Enabled = false;
            //linkLabel2.Enabled = false;

            linkLblOn.Enabled = false;
            linkLblOff.Enabled = false;

            comboBox1.SelectedIndex = 0;
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        ///-------------------------------------------------------------
        private void gridViewSet(DataGridView dg)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更する

                dg.EnableHeadersVisualStyles = false;
                dg.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                dg.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                // 列ヘッダー表示位置指定
                dg.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                dg.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 10, FontStyle.Regular);

                // データフォント指定
                dg.DefaultCellStyle.Font = new Font("Meiryo UI", 10, FontStyle.Regular);

                // 行の高さ
                dg.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                dg.ColumnHeadersHeight = 22;
                dg.RowTemplate.Height = 22;

                // 全体の高さ
                dg.Height = 332;

                // 奇数行の色
                dg.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;

                // 各列指定
                DataGridViewCheckBoxColumn chk = new DataGridViewCheckBoxColumn();
                chk.Name = ColChk;
                dg.Columns.Add(chk);
                dg.Columns[ColChk].HeaderText = "";

                dg.Columns.Add(ColCode, "コード");
                dg.Columns.Add(ColSznm, "部署名");
                //dg.Columns.Add(ColNin, "人数");
                //dg.Columns.Add(ColID, "ID");

                //dg.Columns[ColID].Visible = false;

                dg.Columns[ColChk].Width = 30;
                dg.Columns[ColCode].Width = 80;
                dg.Columns[ColSznm].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                //dg.Columns[ColNin].Width = 80;

                dg.Columns[ColChk].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dg.Columns[ColCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //dg.Columns[ColNin].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                // 編集可否
                dg.ReadOnly = false;
                foreach (DataGridViewColumn item in dg.Columns)
                {
                    // チェックボックスのみ使用可
                    if (item.Name == ColChk)
                    {
                        dg.Columns[item.Name].ReadOnly = false;
                    }
                    else
                    {
                        dg.Columns[item.Name].ReadOnly = true;
                    }
                }

                // 行ヘッダを表示しない
                dg.RowHeadersVisible = false;

                // 選択モード
                dg.SelectionMode = DataGridViewSelectionMode.CellSelect;
                dg.MultiSelect = false;

                // 追加行表示しない
                dg.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                dg.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                dg.AllowUserToOrderColumns = false;

                // 列サイズ変更禁止
                dg.AllowUserToResizeColumns = true;

                // 行サイズ変更禁止
                dg.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //dg.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                // 罫線
                dg.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                dg.CellBorderStyle = DataGridViewCellBorderStyle.None;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void departmentShow()
        {
            // 接続文字列取得
            string sc = sqlControl.obcConnectSting.get(_dbName);

            // 部門一覧表示
            //gridViewShowData(sc, dataGridView1);
            gridViewShowBusho_Obc(sc, dataGridView1);
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     グリッドビューへ部門情報を表示する </summary>
        /// <param name="dg">
        ///     DataGridViewオブジェクト名</param>
        ///---------------------------------------------------------------------
        private void gridViewShowData(string sConnect, DataGridView dg)
        {
            sqlControl.DataControl sdCon = new Common.sqlControl.DataControl(sConnect);
            sqlControl.DataControl sdCon2 = new Common.sqlControl.DataControl(sConnect);
            SqlDataReader dR;

            string dt = DateTime.Today.ToShortDateString();

            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT DepartmentID, DepartmentCode, DepartmentName ");
            sb.Append("FROM tbDepartment ");
            sb.Append("where EstablishDate <= '").Append(dt).Append("'");
            sb.Append(" and AbolitionDate >= '").Append(dt).Append("'");
            sb.Append(" and ValidDate <= '").Append(dt).Append("'");
            sb.Append(" and InValidDate >= '").Append(dt).Append("'");
            sb.Append(" order by DepartmentCode");

            dR = sdCon.free_dsReader(sb.ToString());

            try
            {
                //グリッドビューに表示する
                int iX = 0;
                dg.RowCount = 0;

                while (dR.Read())
                {
                    // 所属社員がいないときはネグる
                    int nin = getBumonEmpCount(sdCon2, dR["DepartmentCode"].ToString(), DateTime.Today);
                    if (nin == global.flgOff)
                    {
                        continue;
                    }

                    //データグリッドにデータを表示する
                    dg.Rows.Add();

                    dg[ColChk, iX].Value = false;

                    string bCode = string.Empty;

                    if (dR["DepartmentCode"].ToString().Trim().Length > 5)
                    {
                        bCode = dR["DepartmentCode"].ToString().Substring(15 - 5, 5);
                    }
                    else
                    {
                        bCode = dR["DepartmentCode"].ToString().Trim().PadLeft(5, '0');
                    }

                    dg[ColCode, iX].Value = bCode;
                    dg[ColSznm, iX].Value = dR["DepartmentName"].ToString().Trim();
                    dg[ColNin, iX].Value = nin.ToString();
                    dg[ColID, iX].Value = dR["DepartmentID"].ToString();

                    iX++;
                }

                dg.Sort(dg.Columns[ColCode], ListSortDirection.Ascending);

                dg.CurrentCell = null;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            }
            finally
            {
                dR.Close();
                sdCon.Close();
                sdCon2.Close();
            }

            // 部門情報がないとき
            if (dg.RowCount == 0)
            {
                MessageBox.Show("就業奉行に部門情報が存在しません", "部署取得", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                Environment.Exit(0);
            }
        }

        ///------------------------------------------------------------------------
        /// <summary>
        ///     就業奉行から部署一覧を表示する : 2017/09/23 </summary>
        /// <param name="sConnect">
        ///     接続文字列</param>
        /// <param name="dg">
        ///     データグリッドビューオブジェクト</param>
        ///------------------------------------------------------------------------
        private void gridViewShowBusho_Obc(string sConnect, DataGridView dg)
        {
            sqlControl.DataControl sdCon = new Common.sqlControl.DataControl(sConnect);
            sqlControl.DataControl sdCon2 = new Common.sqlControl.DataControl(sConnect);
            SqlDataReader dR = null;

            try
            {
                string b = string.Empty;
                string dt = DateTime.Today.ToShortDateString();

                StringBuilder sb = new StringBuilder();
                sb.Append("select DepartmentID, right(rtrim(DepartmentCode), 5) as DepartmentCode, DepartmentName ");
                sb.Append("FROM tbDepartment ");
                sb.Append("where EstablishDate <= '").Append(dt).Append("'");
                sb.Append(" and AbolitionDate >= '").Append(dt).Append("'");
                sb.Append(" and ValidDate <= '").Append(dt).Append("'");
                sb.Append(" and InValidDate >= '").Append(dt).Append("'");
                sb.Append(" order by DepartmentCode");

                dR = sdCon.free_dsReader(sb.ToString());

                int iX = 0;
                dg.RowCount = 0;

                while (dR.Read())
                {
                    // 検索用部署コード
                    if (Utility.StrtoInt(dR["DepartmentCode"].ToString()) != global.flgOff)
                    {
                        b = dR["DepartmentCode"].ToString().Trim().PadLeft(15, '0');
                    }
                    else
                    {
                        b = dR["DepartmentCode"].ToString().Trim().PadRight(15, ' ');
                    }

                    // 所属社員がいないときはネグる
                    if (getBumonEmpCount(sdCon2, b, DateTime.Today) == global.flgOff)
                    {
                        continue;
                    }
                    
                    //データグリッドにデータを表示する
                    dg.Rows.Add();

                    dg[ColChk, iX].Value = false;
                    dg[ColCode, iX].Value = dR["DepartmentCode"].ToString();
                    dg[ColSznm, iX].Value = dR["DepartmentName"].ToString();

                    iX++;
                }

                dg.Sort(dg.Columns[ColCode], ListSortDirection.Ascending);

                dg.CurrentCell = null;

                linkLabel3.Enabled = true;
                linkLblOn.Enabled = true;
                linkLblOff.Enabled = true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            }
            finally
            {
                dR.Close();
                sdCon.Close();
                sdCon2.Close();
            }
        }


        private void linkLblOn_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("全ての部署を印刷対象とします。よろしいですか。", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                dataGridView1[ColChk, i].Value = true;
            }
        }

        private void linkLblOff_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("全ての部署を印刷対象外とします。よろしいですか。", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                dataGridView1[ColChk, i].Value = false;
            }
        }

        ///-----------------------------------------------------------------
        /// <summary>
        ///     任意の部門の社員数を取得する </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl </param>
        /// <param name="strCode">
        ///     SQL文字列</param>
        /// <param name="sDt">
        ///     基準年月日 : 2017/09/28</param>
        /// <returns>
        ///     人数</returns>
        ///-----------------------------------------------------------------
        private int getBumonEmpCount(sqlControl.DataControl sdCon, string strCode, DateTime sDt)
        {
            int nin = 0;
            SqlDataReader cDr = sdCon.free_dsReader(Utility.getEmployeeCount(strCode, sDt));
            while (cDr.Read())
            {
                nin = Utility.StrtoInt(cDr["cnt"].ToString());
                break;
            }

            cDr.Close();

            return nin;
        }

        private void main()
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

            // 2017/09/21
            if (comboBox1.SelectedIndex == 0 && label6.Text == string.Empty)
            {
                MessageBox.Show("当月勤怠実績データを選択してください", "参照ファイル未設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                button2.Focus();
                return;
            }

            // 2017/10/16
            //if (comboBox1.SelectedIndex == 0 && label1.Text == string.Empty)
            if (label1.Text == string.Empty) // 2018/01/18
            {
                MessageBox.Show("年休・積休怠データを選択してください", "参照ファイル未設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                button3.Focus(); 
                return;
            }

            int pCnt = 0;

            // 選択部署
            foreach (DataGridViewRow r in dataGridView1.Rows)
            {
                // チェックされている部署を対象とする
                if (dataGridView1[ColChk, r.Index].Value.ToString() == "True")
                {
                    pCnt++;
                }
            }

            if (pCnt == 0)
            {
                MessageBox.Show("印刷する部署を選択してください", "印刷部署", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            string prnName = string.Empty;
            if (comboBox1.SelectedIndex == 1)
            {
                prnName = "白紙の勤怠表";
            }
            else
            {
                prnName = "勤務実績が印字された勤怠表";
            }

            if (MessageBox.Show(prnName + "を発行します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            this.Cursor = Cursors.WaitCursor;

            //// 実績入り勤怠表を印字するとき
            //if (!checkBox1.Checked)
            //{
            //    // 勤怠データ配列読み込み
            //    workArray = System.IO.File.ReadAllLines(label6.Text, Encoding.Default);
            //}
            
            // 奉行データベース接続
            string sc = null;
            sqlControl.DataControl sdCon = null;

            try
            {
                int yy = Utility.StrtoInt(txtYear.Text);
                int mm = Utility.StrtoInt(txtMonth.Text);

                dAdp.Fill(dts.休日);

                // 奉行マスター接続
                sc = sqlControl.obcConnectSting.get(_dbName);
                sdCon = new Common.sqlControl.DataControl(sc);

                // 勤怠表作成
                if (comboBox1.SelectedIndex == 0)
                {
                    // 実績印字勤怠表
                    kintaiSumNew(sdCon);    // 2017/10/05
                }
                else
                {
                    // 白紙勤怠表
                    kintaiSumTemplate(sdCon);    // 2017/10/05
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                // 奉行マスター接続解除
                if (sdCon != null)
                {
                    if (sdCon.Cn.State == ConnectionState.Open)
                    {
                        sdCon.Close();
                    }
                }

                Cursor = Cursors.Default;
            }
        }

        ///-------------------------------------------------------------------------
        /// <summary>
        ///     勤怠表白紙作成 : 2017/10/05 </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl </param>
        ///-------------------------------------------------------------------------
        private void kintaiSumTemplate(sqlControl.DataControl sdCon)
        {
            // エクセルオブジェクト
            Excel.Application oXls = new Excel.Application();
            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(Properties.Settings.Default.xlsKintai,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];
            Excel.Worksheet oxlsMsSheet = (Excel.Worksheet)oXlsBook.Sheets[1]; // テンプレートシート
            oxlsSheet.Select(Type.Missing);

            Excel.Range rng = null;
            object[,] rtnArray = null;

            //当月用勤怠表シート

            // 休日背景色を取得
            rng = oxlsMsSheet.Cells[6, 1];
            double dd = rng.Interior.Color;

            // タイトル
            oxlsMsSheet.Cells[1, 1] = txtYear.Text + "年" + txtMonth.Text.PadLeft(2, ' ') + "月度　勤　怠　表";

            // シートの日付領域を配列に読み込む
            object[,] shDays = null;
            rng = oxlsMsSheet.Range[oxlsMsSheet.Cells[6, 4], oxlsMsSheet.Cells[6, 34]];
            shDays = (object[,])rng.Value2;

            DateTime dt = new DateTime(Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text), 1);
            int fDay = dt.AddMonths(1).AddDays(-1).Day; // 月末日

            // 当月日付を配列にセットする
            for (int i = 1; i <= shDays.Length; i++)
            {
                if (i <= fDay)
                {
                    shDays[1, i] = i.ToString();

                    // 休日の背景色を設定
                    DateTime dt2 = new DateTime(Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text), i);

                    if (dts.休日.Any(a => a.年月日 == dt2))
                    {
                        rng = oxlsMsSheet.Range[oxlsMsSheet.Cells[6, 3 + i], oxlsMsSheet.Cells[41, 3 + i]];
                        rng.Interior.Color = dd;
                    }
                }
                else
                {
                    shDays[1, i] = string.Empty;
                }
            }

            // 日付配列をシートの日付領域に一括貼り付け
            rng = oxlsMsSheet.Range[oxlsMsSheet.Cells[6, 4], oxlsMsSheet.Cells[6, 34]];
            rng.Value2 = shDays;

            int xR = 0;
            int pCnt = 1;

            // 雇用区分
            int kKbn = 0;

            string sBushoNum = string.Empty;    // 部署コード
            string sShainNum = string.Empty;    // 社員番号

            try
            {
                int iX = 0;     // 行カウンタ
                int bCnt = 0;   // 部署別カウンタ

                // 部署ループ
                foreach (DataGridViewRow r in dataGridView1.Rows)
                {
                    // チェックされている部署を対象とする
                    if (dataGridView1[ColChk, r.Index].Value.ToString() == "False")
                    {
                        continue;
                    }

                    // 部署コード取得
                    string bCode = dataGridView1[ColCode, r.Index].Value.ToString().Trim();

                    // DepartmentCode取得
                    string strCode = Utility.getDepartmentCode(bCode);

                    // 所属・役職・ライン№の基準日を取得 2018/01/18
                    //DateTime bDt;
                    DateTime rDt;
                    if (!DateTime.TryParse(txtYear.Text + "/" + txtMonth.Text + "/01", out rDt))
                    {
                        rDt = DateTime.Today;
                        //bDt = DateTime.Today;
                    }
                    else
                    {
                        // 対象年月の末日
                        //bDt = rDt.AddMonths(1).AddDays(-1);
                    }

                    // 該当部署の社員情報データリーダーを取得する
                    //SqlDataReader dr = sdCon.free_dsReader(Utility.getEmployeeOrder(strCode, DateTime.Today)); // 2018/01/18

                    // 退職基準日引数を追加 2018/02/19
                    //SqlDataReader dr = sdCon.free_dsReader(Utility.getEmployeeKintaiRep(strCode, bDt, rDt)); // 2018/01/18
                    SqlDataReader dr = sdCon.free_dsReader(Utility.getEmployeeKintaiRep(strCode, DateTime.Today, rDt)); // 2018/03/08

                    string[] saArray = null;

                    int iSa = 0;

                    // 並び替えた社員番号から配列を作成する
                    while (dr.Read())
                    {
                        // 社員番号
                        string sCode = Utility.StrtoInt(dr["EmployeeNo"].ToString()).ToString().PadLeft(6, '0');

                        Array.Resize(ref saArray, iSa + 1);
                        saArray[iSa] = sCode;
                        iSa++;
                    }

                    dr.Close();

                    // 社員が存在しないとき
                    if (saArray == null)
                    {
                        continue;
                    }

                    // 社員番号配列を順に読む
                    for (int i = 0; i < saArray.Length; i++)
                    {
                        // 社員番号でブレーク
                        if (sShainNum != saArray[i])
                        {
                            if (sShainNum != string.Empty)
                            {
                                // 配列からシートへ一括して出力する
                                //setXlsData(ref rtnArray, kKbn, sShainNum);
                                rng = oxlsSheet.Range[oxlsSheet.Cells[xR, 4], oxlsSheet.Cells[xR + 3, oxlsMsSheet.UsedRange.Columns.Count]];
                                rng.Value2 = rtnArray;

                                iX++;   // 行カウンタ加算
                            }

                            // 初期処理、改ページ時または部署コードが変わったら
                            if (iX > 6 || sBushoNum != bCode)
                            {
                                // 部署内カウンタ
                                if (iX > 6)
                                {
                                    bCnt++; // 同じ部署で次のページ
                                }
                                else
                                {
                                    bCnt = 0;   // 次の部署
                                }

                                // テンプレートシートを追加する
                                pCnt++;
                                oxlsMsSheet.Copy(Type.Missing, oXlsBook.Sheets[pCnt - 1]);
                                oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt];

                                // 部署名
                                string buSho = getDepartmentName(bCode, sdCon);
                                oxlsSheet.Cells[3, 2] = bCode + "  " + buSho;

                                // シート名
                                string sheetName = bCode.PadLeft(5, '0') + " " + buSho;
                                if (bCnt > 0)
                                {
                                    sheetName += "(" + bCnt.ToString() + ")";
                                }

                                oxlsSheet.Name = sheetName;

                                // 行カウンタ初期化
                                iX = 0;
                            }

                            // 合計欄初期化
                            totalClear();

                            // シートのセルを一括して配列に取得します
                            xR = iX * 5 + 7;
                            rng = oxlsMsSheet.Range[oxlsMsSheet.Cells[xR, 4], oxlsMsSheet.Cells[xR + 3, oxlsMsSheet.UsedRange.Columns.Count]];
                            rtnArray = (object[,])rng.Value2;

                            // 氏名と雇用区分を取得
                            kKbn = 0;
                            string sName = string.Empty;
                            getEmployee(sdCon, saArray[i], ref sName, ref kKbn);

                            //// 社員番号と氏名をシートに貼り付ける
                            //oxlsSheet.Cells[xR, 1] = saArray[i];
                            //oxlsSheet.Cells[xR + 1, 1] = sName;
                            //oxlsSheet.Cells[xR + 3, 1] = "年休残数：";
                            //oxlsSheet.Cells[xR + 4, 1] = "積立残数：";
                            
                            // 2018/01/18
                            // 年休残・積休残を取得する
                            decimal nenZan = 0;
                            decimal tsumiZan = 0;
                            getNenkyuData(Utility.StrtoInt(saArray[i]), out nenZan, out tsumiZan);

                            // 社員番号と氏名をシートに貼り付ける
                            oxlsSheet.Cells[xR, 1] = saArray[i];
                            oxlsSheet.Cells[xR + 1, 1] = sName;
                            oxlsSheet.Cells[xR + 3, 1] = "年休残数：" + nenZan;
                            oxlsSheet.Cells[xR + 4, 1] = "積立残数：" + tsumiZan;
                        }
                            
                        sBushoNum = bCode;
                        sShainNum = saArray[i];
                    }
                }

                if (rtnArray != null)
                {
                    // 配列からシートへ一括して出力する
                    //setXlsData(ref rtnArray, kKbn, sShainNum);
                    rng = oxlsSheet.Range[oxlsSheet.Cells[xR, 4], oxlsSheet.Cells[xR + 3, oxlsMsSheet.UsedRange.Columns.Count]];
                    rng.Value2 = rtnArray;

                    // 1枚目はテンプレートシートなので印刷時には削除する
                    oXls.DisplayAlerts = false;
                    oXlsBook.Sheets[1].Delete();

                    // 1枚目のシートが表示されるようにする
                    oxlsSheet = oXlsBook.Sheets[1];
                    oxlsSheet.Select();

                    //マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    // 確認のためExcelのウィンドウを表示する
                    oXls.Visible = true;

                    // 印刷
                    oXlsBook.PrintOutEx(Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    //oXlsBook.PrintOutEx();
                    //oXlsBook.PrintPreview(true);

                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    //保存処理
                    oXls.DisplayAlerts = false;

                    DialogResult ret;

                    //ダイアログボックスの初期設定
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.Title = "勤怠表";
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    saveFileDialog1.FileName = txtYear.Text + "年" + txtMonth.Text.PadLeft(2, ' ') + "月 勤怠表";
                    saveFileDialog1.Filter = "Microsoft Office Excelファイル(*.xlsx)|*.xlsx|全てのファイル(*.*)|*.*";

                    //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
                    string fileName;
                    ret = saveFileDialog1.ShowDialog();

                    if (ret == System.Windows.Forms.DialogResult.OK)
                    {
                        fileName = saveFileDialog1.FileName;
                        oXlsBook.SaveAs(fileName, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing,
                                        Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }

                    MessageBox.Show("終了しました", "勤怠表作成");
                }
                else
                {
                    //マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    //保存処理
                    oXls.DisplayAlerts = false;

                    MessageBox.Show("勤怠データはありませんでした", "勤怠表作成");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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

        ///-------------------------------------------------------------------------
        /// <summary>
        ///     勤怠表作成 : 2017/10/05 </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl </param>
        ///-------------------------------------------------------------------------
        private void kintaiSumNew(sqlControl.DataControl sdCon)
        {
            // エクセルオブジェクト
            Excel.Application oXls = new Excel.Application();
            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(Properties.Settings.Default.xlsKintai,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];
            Excel.Worksheet oxlsMsSheet = (Excel.Worksheet)oXlsBook.Sheets[1]; // テンプレートシート
            oxlsSheet.Select(Type.Missing);

            Excel.Range rng = null;
            object[,] rtnArray = null;

            //当月用勤怠表シート

            // 休日背景色を取得
            rng = oxlsMsSheet.Cells[6, 1];
            double dd = rng.Interior.Color;

            // タイトル
            oxlsMsSheet.Cells[1, 1] = txtYear.Text + "年" + txtMonth.Text.PadLeft(2, ' ') + "月度　勤　怠　表";

            // シートの日付領域を配列に読み込む
            object[,] shDays = null;
            rng = oxlsMsSheet.Range[oxlsMsSheet.Cells[6, 4], oxlsMsSheet.Cells[6, 34]];
            shDays = (object[,])rng.Value2;

            DateTime dt = new DateTime(Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text), 1);
            int fDay = dt.AddMonths(1).AddDays(-1).Day; // 月末日

            // 当月日付を配列にセットする
            for (int i = 1; i <= shDays.Length; i++)
            {
                if (i <= fDay)
                {
                    shDays[1, i] = i.ToString();

                    // 休日の背景色を設定
                    DateTime dt2 = new DateTime(Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text), i);

                    if (dts.休日.Any(a => a.年月日 == dt2))
                    {
                        rng = oxlsMsSheet.Range[oxlsMsSheet.Cells[6, 3 + i], oxlsMsSheet.Cells[41, 3 + i]];
                        rng.Interior.Color = dd;
                    }
                }
                else
                {
                    shDays[1, i] = string.Empty;
                }
            }

            // 日付配列をシートの日付領域に一括貼り付け
            rng = oxlsMsSheet.Range[oxlsMsSheet.Cells[6, 4], oxlsMsSheet.Cells[6, 34]];
            rng.Value2 = shDays;

            int xR = 0;
            int pCnt = 1;

            // 雇用区分
            int kKbn = 0;

            string sBushoNum = string.Empty;    // 部署コード
            string sShainNum = string.Empty;    // 社員番号

            try
            {
                int iX = 0;     // 行カウンタ
                int bCnt = 0;   // 部署別カウンタ

                // 部署ループ
                foreach (DataGridViewRow r in dataGridView1.Rows)
                {
                    // チェックされている部署を対象とする
                    if (dataGridView1[ColChk, r.Index].Value.ToString() == "False")
                    {
                        continue;
                    }

                    // 部署コード取得
                    string bCode = dataGridView1[ColCode, r.Index].Value.ToString().Trim();

                    // DepartmentCode取得
                    string strCode = Utility.getDepartmentCode(bCode);

                    // 所属・役職・ライン№の基準日を取得 2018/01/18
                    //DateTime bDt;   // 基準日
                    DateTime rDt;   // 退職基準日
                    if (!DateTime.TryParse(txtYear.Text + "/" + txtMonth.Text + "/01", out rDt))
                    {
                        rDt = DateTime.Today;
                        //bDt = DateTime.Today; // 2018/03/02コメント化
                    }
                    else
                    {
                        // 対象年月の末日
                        //bDt = rDt.AddMonths(1).AddDays(-1); // 2018/03/02コメント化
                    }

                    // 該当部署の社員情報データリーダーを取得する : 対象年月末日時点(2018/01/17)
                    //SqlDataReader dr = sdCon.free_dsReader(Utility.getEmployeeOrder(strCode, DateTime.Today));　// 2018/01/17

                    // 退職基準日引数を追加 2018/02/19
                    // 2018/03/02　基準日は当日にまた戻す
                    SqlDataReader dr = sdCon.free_dsReader(Utility.getEmployeeKintaiRep(strCode, DateTime.Today, rDt));

                    string[] saArray = null;

                    int iSa = 0;

                    // 並び替えた社員番号から配列を作成する
                    while (dr.Read())
                    {
                        // 社員番号
                        string sCode = Utility.StrtoInt(dr["EmployeeNo"].ToString()).ToString().PadLeft(6, '0');

                        Array.Resize(ref saArray, iSa + 1);
                        saArray[iSa] = sCode;
                        iSa++;
                    }

                    dr.Close();

                    // 社員が存在しないとき
                    if (saArray == null)
                    {
                        continue;
                    }

                    // 2018/02/05
                    string wkDate = string.Empty;

                    // 社員番号配列を順に読む
                    for (int i = 0; i < saArray.Length; i++)
                    {
                        // 奉行勤怠データを読む
                        foreach (var s in workArray)
                        {
                            // カンマ区切りで分割して配列に格納する
                            string[] t = s.Split(',');

                            // 該当社員番号か？
                            if (saArray[i] != t[0].Replace("\"", ""))
                            {
                                continue;
                            }

                            // 日付情報を取得する：2017/09/22
                            string strDate = t[2].Replace("\"", "");
                            string f = "ggyy年MM月dd日";
                            //string f = "yyyy/MM/dd";
                            System.Globalization.CultureInfo ci = new System.Globalization.CultureInfo("ja-JP");
                            //DateTime iDt = DateTime.ParseExact(strDate, "ggyy年MM月dd日", ci);
                            DateTime iDt = DateTime.Parse(strDate, ci, System.Globalization.DateTimeStyles.AssumeLocal);

                            // 社員番号でブレーク
                            if (sShainNum != t[0].Replace("\"", ""))
                            {
                                if (sShainNum != string.Empty)
                                {
                                    // 配列からシートへ一括して出力する
                                    setXlsData(ref rtnArray, kKbn, sShainNum);
                                    rng = oxlsSheet.Range[oxlsSheet.Cells[xR, 4], oxlsSheet.Cells[xR + 3, oxlsMsSheet.UsedRange.Columns.Count]];
                                    rng.Value2 = rtnArray;

                                    oxlsSheet.Cells[xR + 4, 38] = shinyaTime.ToString("##0.0");     // 深夜残業
                                    oxlsSheet.Cells[xR + 4, 36] = tsumikyuuDays.ToString("##0.0");  // 積休使用数
                                    //oxlsSheet.Cells[xR + 4, 40] = koukiTime.ToString("##0.0");      // 工機夜時間 : 2017/09/23

                                    iX++;   // 行カウンタ加算
                                }

                                // 初期処理、改ページ時または部署コードが変わったら
                                if (iX > 6 || sBushoNum != bCode)
                                {
                                    // 部署内カウンタ
                                    if (iX > 6)
                                    {
                                        bCnt++; // 同じ部署で次のページ
                                    }
                                    else
                                    {
                                        bCnt = 0;   // 次の部署
                                    }

                                    // テンプレートシートを追加する
                                    pCnt++;
                                    oxlsMsSheet.Copy(Type.Missing, oXlsBook.Sheets[pCnt - 1]);
                                    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt];

                                    // 部署名
                                    string buSho = getDepartmentName(bCode, sdCon);
                                    oxlsSheet.Cells[3, 2] = bCode + "  " + buSho;

                                    // シート名
                                    string sheetName = bCode.PadLeft(5, '0') + " " + buSho;
                                    if (bCnt > 0)
                                    {
                                        sheetName += "(" + bCnt.ToString() + ")";
                                    }

                                    oxlsSheet.Name = sheetName;

                                    // 行カウンタ初期化
                                    iX = 0;
                                }

                                // 合計欄初期化
                                totalClear();

                                // シートのセルを一括して配列に取得します
                                xR = iX * 5 + 7;
                                rng = oxlsMsSheet.Range[oxlsMsSheet.Cells[xR, 4], oxlsMsSheet.Cells[xR + 3, oxlsMsSheet.UsedRange.Columns.Count]];
                                rtnArray = (object[,])rng.Value2;

                                // 氏名と雇用区分を取得
                                kKbn = 0;
                                string sName = string.Empty;
                                getEmployee(sdCon, t[0].Replace("\"", ""), ref sName, ref kKbn);

                                // 年休残・積休残を取得する
                                decimal nenZan = 0;
                                decimal tsumiZan = 0;
                                getNenkyuData(Utility.StrtoInt(t[0].Replace("\"", "")), out nenZan, out tsumiZan);

                                // 社員番号と氏名をシートに貼り付ける
                                oxlsSheet.Cells[xR, 1] = t[0].Replace("\"", "");
                                oxlsSheet.Cells[xR + 1, 1] = sName;
                                oxlsSheet.Cells[xR + 3, 1] = "年休残数：" + nenZan;
                                oxlsSheet.Cells[xR + 4, 1] = "積立残数：" + tsumiZan;
                            }

                            // 日付毎のデータを配列にセットする

                            // CSVデータからシフトコードを取得する 2017/09/22
                            string sftCode = t[3].Replace("\"", "");
                            sftCode = sftCode.Trim().PadLeft(3, '0');

                            // 呼出回数：事由[30]カウント 2018/05/30
                            if (Utility.StrtoInt(t[5].Replace("\"", "")).ToString() == JIYU_YOBIDASHI ||
                                     Utility.StrtoInt(t[7].Replace("\"", "")).ToString() == JIYU_YOBIDASHI ||
                                     Utility.StrtoInt(t[9].Replace("\"", "")).ToString() == JIYU_YOBIDASHI) // 2018/05/28
                            {
                                // 事由コードが呼出記号[30]のとき呼出回数に加算 2018/05/28
                                yobidashi++;
                            }

                            //----------------------------------------------------
                            //  出勤欄記号
                            //----------------------------------------------------
                            // 休日出勤のとき : 休出・休憩ありを条件に追加 2018/02/05
                            if (sftCode == SHIFT_KYUSHUTSU || sftCode == SHIFT_KYUKEI_KYUSHUTSU)
                            {
                                kyushutsuDays++;    //　日数カウント : 2017/10/05  2018/09/18 有効化
                                rtnArray[1, iDt.Day] = "休出";    // 2017/09/22
                            }
                            //else if (sftCode == SHIFT_HEIKITAKUGO)　// 2018/05/28 コメント化
                            //{
                            //    // 平日帰宅後回数カウント
                            //    yobidashi++;
                            //}
                            //else if (sftCode == SHIFT_KYUKITAKUGO)　// 2018/05/28 コメント化
                            //{
                            //    // 休日帰宅後回数カウント
                            //    yobidashi++;
                            //}
                            else
                            {
                                // 事由配列 : 2017/09/22 CSVデータから取得
                                //string[] jArray = { Utility.StrtoInt(t[5].Replace("\"", "")).ToString(), Utility.StrtoInt(t[7].Replace("\"", "")).ToString(), Utility.StrtoInt(t[9].Replace("\"", "")).ToString() };

                                // 2017/11/21
                                string[] jArray = { Utility.StrtoInt(t[5].Replace("\"", "")).ToString() + ":" + t[6].Replace("\"", "").ToString(),
                                                    Utility.StrtoInt(t[7].Replace("\"", "")).ToString() + ":" + t[8].Replace("\"", "").ToString(),
                                                    Utility.StrtoInt(t[9].Replace("\"", "")).ToString()  + ":" + t[10].Replace("\"", "").ToString()};
                                
                                // 事由から出勤欄の記号を求める
                                int jSt = 0;
                                rtnArray[1, iDt.Day] = getShukkinKigou(jArray, ref jSt);    // 2017/09/22

                                // 事由なしのとき）
                                if (jSt == 0)
                                {
                                    // 出勤欄表示文字列取得
                                    rtnArray[1, iDt.Day] = getWorkMark(kKbn, t);     // 2017/09/22
                                }
                            }

                            //--------------------------------------------------------
                            //  パート社員のとき出勤時間を取得する
                            //--------------------------------------------------------
                            if (kKbn == KBN_PART || kKbn == KBN_PART_2 || kKbn == KBN_PART_3)
                            {
                                double wt = Utility.StrtoDouble(t[15].Replace("\"", ""));
                                workTime += wt;     // 加算
                            }

                            //--------------------------------------------------------------
                            //  普通残業時間を取得・加算 : 2018/02/05 同日残業時間は加算
                            //--------------------------------------------------------------
                            double dZan = Utility.StrtoDouble(t[36].Replace("\"", ""));
                            dZan += Utility.StrtoDouble(t[33].Replace("\"", ""));   // 早出残業を加算　2017/09/22

                            if (dZan > 0)
                            {
                                double zz = Utility.StrtoDouble(Utility.NulltoStr(rtnArray[2, iDt.Day]));
                                rtnArray[2, iDt.Day] = (zz + dZan).ToString("#,##0.0");    // 2017/09/22
                            }

                            zanTime += dZan;

                            //-------------------------------------------------------
                            //  深夜残業時間を取得・加算 : 2018/02/05 同日深夜残業時間は加算
                            //-------------------------------------------------------
                            double dSZan = Utility.StrtoDouble(t[39].Replace("\"", ""));

                            if (dSZan > 0)
                            {
                                double zz = Utility.StrtoDouble(Utility.NulltoStr(rtnArray[3, iDt.Day]));
                                rtnArray[3, iDt.Day] = (zz + dSZan).ToString("#,##0.0");   // 2017/09/22
                            }

                            shinyaTime += dSZan;

                            //-------------------------------------------------------
                            //  遅早時間を取得・加算
                            //-------------------------------------------------------
                            // 通常遅早時間
                            double dChisou = Utility.StrtoDouble(t[24].Replace("\"", "")) + Utility.StrtoDouble(t[27].Replace("\"", "")) + Utility.StrtoDouble(t[48].Replace("\"", ""));

                            if (dChisou > 0)
                            {
                                rtnArray[4, iDt.Day] = dChisou.ToString("#,##0.0");
                            }

                            chisouTime += Utility.StrtoDouble(t[24].Replace("\"", "")) + Utility.StrtoDouble(t[27].Replace("\"", ""));

                            // 休業遅早時間
                            chisouTimeKyuGyo += Utility.StrtoDouble(t[48].Replace("\"", "")); 

                            //-------------------------------------------------------------------
                            //  休日残業時間を取得・加算 : 2018/02/05 同日休日残業時間は加算
                            //-------------------------------------------------------------------
                            double zan = Utility.StrtoDouble(t[42].Replace("\"", ""));

                            if (zan > 0)
                            {
                                double zz = Utility.StrtoDouble(Utility.NulltoStr(rtnArray[2, iDt.Day]));
                                rtnArray[2, iDt.Day] = (zz + zan).ToString("#,##0.0");
                                kyushutsuTime += zan;
                            }

                            //---------------------------------------------------------------------
                            //  休日深夜残業時間を取得・加算 : 2018/02/05 同日休日残業時間は加算
                            //---------------------------------------------------------------------
                            dSZan = Utility.StrtoDouble(t[45].Replace("\"", ""));

                            if (dSZan > 0)
                            {
                                double zz = Utility.StrtoDouble(Utility.NulltoStr(rtnArray[3, iDt.Day]));
                                rtnArray[3, iDt.Day] = (zz + dSZan).ToString("#,##0.0");
                                kyusuhtsuShinyaTime += dSZan; 
                            }

                            sBushoNum = bCode;
                            sShainNum = t[0].Replace("\"", "");
                        }
                    }
                }

                if (rtnArray != null)
                {
                    // 配列からシートへ一括して出力する
                    setXlsData(ref rtnArray, kKbn, sShainNum);
                    rng = oxlsSheet.Range[oxlsSheet.Cells[xR, 4], oxlsSheet.Cells[xR + 3, oxlsMsSheet.UsedRange.Columns.Count]];
                    rng.Value2 = rtnArray;

                    oxlsSheet.Cells[xR + 4, 38] = shinyaTime.ToString("##0.0");     // 深夜残業 : 2017/10/04
                    oxlsSheet.Cells[xR + 4, 36] = tsumikyuuDays.ToString("##0.0");  // 積休使用数：2017/10/04
                    //oxlsSheet.Cells[xR + 4, 40] = shinyaTime.ToString("##0.0");     // 工機夜時間 : 2017/09/23

                    // 1枚目はテンプレートシートなので印刷時には削除する
                    oXls.DisplayAlerts = false;
                    oXlsBook.Sheets[1].Delete();

                    // 1枚目のシートが表示されるようにする
                    oxlsSheet = oXlsBook.Sheets[1];
                    oxlsSheet.Select();

                    //マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    // 確認のためExcelのウィンドウを表示する
                    oXls.Visible = true;

                    // 印刷
                    oXlsBook.PrintOutEx(Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    //oXlsBook.PrintOutEx();
                    //oXlsBook.PrintPreview(true);

                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    //保存処理
                    oXls.DisplayAlerts = false;

                    DialogResult ret;

                    //ダイアログボックスの初期設定
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.Title = "勤怠表";
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    saveFileDialog1.FileName = txtYear.Text + "年" + txtMonth.Text.PadLeft(2, ' ') + "月 勤怠表";
                    saveFileDialog1.Filter = "Microsoft Office Excelファイル(*.xlsx)|*.xlsx|全てのファイル(*.*)|*.*";

                    //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
                    string fileName;
                    ret = saveFileDialog1.ShowDialog();

                    if (ret == System.Windows.Forms.DialogResult.OK)
                    {
                        fileName = saveFileDialog1.FileName;
                        oXlsBook.SaveAs(fileName, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing,
                                        Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }

                    MessageBox.Show("終了しました", "勤怠表作成");
                }
                else
                {
                    //マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    //保存処理
                    oXls.DisplayAlerts = false;

                    MessageBox.Show("勤怠データはありませんでした", "勤怠表作成");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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

        ///-------------------------------------------------------------------------
        /// <summary>
        ///     勤怠表作成 </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl </param>
        ///-------------------------------------------------------------------------
        private void kintaiSum(sqlControl.DataControl sdCon)
        {
            // エクセルオブジェクト
            Excel.Application oXls = new Excel.Application();
            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(Properties.Settings.Default.xlsKintai,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];
            Excel.Worksheet oxlsMsSheet = (Excel.Worksheet)oXlsBook.Sheets[1]; // テンプレートシート
            oxlsSheet.Select(Type.Missing);

            Excel.Range rng = null;
            object[,] rtnArray = null;

             //当月用勤怠表シート

            // 休日背景色を取得
            rng = oxlsMsSheet.Cells[6, 1];
            double dd = rng.Interior.Color;

            // タイトル
            oxlsMsSheet.Cells[1, 1] = txtYear.Text + "年" + txtMonth.Text.PadLeft(2, ' ') + "月度　勤　怠　表";
            
            // シートの日付領域を配列に読み込む
            object[,] shDays = null;
            rng = oxlsMsSheet.Range[oxlsMsSheet.Cells[6, 4], oxlsMsSheet.Cells[6, 34]];
            shDays = (object[,])rng.Value2;

            DateTime dt = new DateTime(Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text), 1);
            int fDay = dt.AddMonths(1).AddDays(-1).Day; // 月末日
            
            // 当月日付を配列にセットする
            for (int i = 1; i <= shDays.Length; i++)
            {
                if (i <= fDay)
                {
                    shDays[1, i] = i.ToString();

                    // 休日の背景色を設定
                    DateTime dt2 = new DateTime(Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text), i);

                    if (dts.休日.Any(a => a.年月日 == dt2))
                    {
                        rng = oxlsMsSheet.Range[oxlsMsSheet.Cells[6, 3 + i], oxlsMsSheet.Cells[41, 3 + i]];
                        rng.Interior.Color = dd;
                    }
                }
                else
                {
                    shDays[1, i] = string.Empty;
                }
            }

            // 日付配列をシートの日付領域に一括貼り付け
            rng = oxlsMsSheet.Range[oxlsMsSheet.Cells[6, 4], oxlsMsSheet.Cells[6, 34]];
            rng.Value2 = shDays;
            
            int xR = 0;
            int pCnt = 1;

            // 雇用区分
            int kKbn = 0;

            string sBushoNum = string.Empty;    // 部署コード
            string sShainNum = string.Empty;    // 社員番号

            try
            {
                int iX = 0;     // 行カウンタ
                int bCnt = 0;   // 部署別カウンタ
              
                // 部署ループ
                foreach (DataGridViewRow r in dataGridView1.Rows)
                {
                    // チェックされている部署を対象とする
                    if (dataGridView1[ColChk, r.Index].Value.ToString() == "False")
                    {
                        continue;
                    }

                    // 部署コード取得
                    string bCode = dataGridView1[ColCode, r.Index].Value.ToString().Trim();

                    // DepartmentCode取得
                    string strCode = Utility.getDepartmentCode(bCode);

                    // 該当部署の社員情報データリーダーを取得する
                    SqlDataReader dr = sdCon.free_dsReader(Utility.getEmployeeOrder(strCode, DateTime.Today));

                    string[] saArray = null;

                    int iSa = 0;

                    // 並び替えた社員番号から配列を作成する
                    while (dr.Read())
                    {
                        // 社員番号
                        string sCode = Utility.StrtoInt(dr["EmployeeNo"].ToString()).ToString().PadLeft(6, '0');

                        Array.Resize(ref saArray, iSa + 1);
                        saArray[iSa] = sCode;
                        iSa++;
                    }

                    dr.Close();

                    // 社員が存在しないとき
                    if (saArray == null)
                    {
                        continue;
                    }

                    // 社員番号配列を順に読む
                    for (int i = 0; i < saArray.Length; i++)
                    {
                        // 奉行勤怠データを読む
                        foreach (var s in workArray)
                        {
                            // カンマ区切りで分割して配列に格納する
                            string[] t = s.Split(',');

                            // 該当社員番号か？
                            if (saArray[i] != t[0].Replace("\"", ""))
                            {
                                continue;
                            }

                            // 日付情報を取得する：2017/09/22
                            string strDate = t[2].Replace("\"", "");
                            string f = "ggyy年MM月dd日";
                            //string f = "yyyy/MM/dd";
                            System.Globalization.CultureInfo ci = new System.Globalization.CultureInfo("ja-JP");
                            //DateTime iDt = DateTime.ParseExact(strDate, "ggyy年MM月dd日", ci);
                            DateTime iDt = DateTime.Parse(strDate, ci, System.Globalization.DateTimeStyles.AssumeLocal);
                            
                            // 社員番号でブレーク
                            if (sShainNum != t[0].Replace("\"", ""))
                            {
                                if (sShainNum != string.Empty)
                                {
                                    // 配列からシートへ一括して出力する
                                    setXlsData(ref rtnArray, kKbn, sShainNum);
                                    rng = oxlsSheet.Range[oxlsSheet.Cells[xR, 4], oxlsSheet.Cells[xR + 3, oxlsMsSheet.UsedRange.Columns.Count]];
                                    rng.Value2 = rtnArray;

                                    oxlsSheet.Cells[xR + 4, 38] = shinyaTime.ToString("##0.0");     // 深夜残業
                                    oxlsSheet.Cells[xR + 4, 36] = tsumikyuuDays.ToString("##0.0");  // 積休使用数
                                    //oxlsSheet.Cells[xR + 4, 40] = koukiTime.ToString("##0.0");      // 工機夜時間 : 2017/09/23

                                    iX++;   // 行カウンタ加算
                                }

                                // 初期処理、改ページ時または部署コードが変わったら
                                if (iX > 6 || sBushoNum != bCode)
                                {
                                    // 部署内カウンタ
                                    if (iX > 6)
                                    {
                                        bCnt++; // 同じ部署で次のページ
                                    }
                                    else
                                    {
                                        bCnt = 0;   // 次の部署
                                    }

                                    // テンプレートシートを追加する
                                    pCnt++;
                                    oxlsMsSheet.Copy(Type.Missing, oXlsBook.Sheets[pCnt - 1]);
                                    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt];

                                    // 部署名
                                    string buSho = getDepartmentName(bCode, sdCon);
                                    oxlsSheet.Cells[3, 2] = bCode + "  " + buSho;

                                    // シート名
                                    string sheetName = bCode.PadLeft(5, '0') + " " + buSho;
                                    if (bCnt > 0)
                                    {
                                        sheetName += "(" + bCnt.ToString() + ")";
                                    }

                                    oxlsSheet.Name = sheetName;

                                    // 行カウンタ初期化
                                    iX = 0;
                                }

                                // 合計欄初期化
                                totalClear();

                                // シートのセルを一括して配列に取得します
                                xR = iX * 5 + 7;
                                rng = oxlsMsSheet.Range[oxlsMsSheet.Cells[xR, 4], oxlsMsSheet.Cells[xR + 3, oxlsMsSheet.UsedRange.Columns.Count]];
                                rtnArray = (object[,])rng.Value2;

                                // 氏名と雇用区分を取得
                                kKbn = 0;
                                string sName = string.Empty;
                                getEmployee(sdCon, t[0].Replace("\"", ""), ref sName, ref kKbn);

                                // 社員番号と氏名をシートに貼り付ける
                                oxlsSheet.Cells[xR, 1] = t[0].Replace("\"", "");
                                oxlsSheet.Cells[xR + 1, 1] = sName;
                                oxlsSheet.Cells[xR + 3, 1] = "年休残数：";
                                oxlsSheet.Cells[xR + 4, 1] = "積立残数：";
                            }

                            // 日付毎のデータを配列にセットする

                            // CSVデータからシフトコードを取得する 2017/09/22
                            string sftCode = t[3].Replace("\"", "");
                            sftCode = sftCode.Trim().PadLeft(3, '0');

                            // 休日出勤のとき : 休出・休憩ありを条件に追加 2018/02/05
                            if (sftCode == SHIFT_KYUSHUTSU || sftCode == SHIFT_KYUKEI_KYUSHUTSU)
                            {
                                //kyushutsuDays++;    //　日数カウント : 2017/10/05

                                // 2017/09/22 : CSVデータから休日残業時間を取得
                                double zan = Utility.StrtoDouble(t[42].Replace("\"", ""));

                                if (zan > 0)
                                {
                                    rtnArray[2, iDt.Day] = zan.ToString("#,##0.0"); // 2017/09/22
                                    kyushutsuTime += zan;   // 休出残業時間加算
                                }

                                rtnArray[1, iDt.Day] = "休出";    // 2017/09/22

                                // 休出深夜残業 : 2017/09/22 CSVデータから取得
                                double dSZan = Utility.StrtoDouble(t[45].Replace("\"", ""));
                                if (dSZan > 0)
                                {
                                    rtnArray[3, iDt.Day] = dSZan.ToString("#,##0.0");   // 2017/09/22
                                    kyusuhtsuShinyaTime += dSZan;   // 休出深夜残業加算
                                }
                            }
                            else if (sftCode == SHIFT_HEIKITAKUGO)
                            {
                                // 平日帰宅後回数カウント
                                yobidashi++;
                            }
                            else if (sftCode == SHIFT_KYUKITAKUGO)
                            {
                                // 休日帰宅後回数カウント
                                yobidashi++;

                                // 休出時間加算：CSVデータから取得 2017/09/22
                                kyushutsuTime += Utility.StrtoDouble(t[42].Replace("\"", ""));

                                // 休出深夜残業加算：CSVデータから取得 2017/09/22
                                kyusuhtsuShinyaTime += Utility.StrtoDouble(t[45].Replace("\"", "")); 
                            }
                            else
                            {
                                // 事由配列 : 2017/09/22 CSVデータから取得
                                string[] jArray = { Utility.StrtoInt(t[5].Replace("\"", "")).ToString() + ":" + t[6].Replace("\"", "").ToString(),
                                                    Utility.StrtoInt(t[7].Replace("\"", "")).ToString() + ":" + t[8].Replace("\"", "").ToString(),
                                                    Utility.StrtoInt(t[9].Replace("\"", "")).ToString()  + ":" + t[10].Replace("\"", "").ToString()};

                                // 事由から出勤欄の記号を求める
                                int jSt = 0;
                                rtnArray[1, iDt.Day] = getShukkinKigou(jArray, ref jSt);    // 2017/09/22

                                // 事由なしのとき）
                                if (jSt == 0)
                                {
                                    // 出勤欄表示文字列取得
                                    rtnArray[1, iDt.Day] = getWorkMark(kKbn, t);     // 2017/09/22
                                }
                                
                                // パート社員のとき出勤時間を取得する
                                if (kKbn == KBN_PART)
                                {
                                    double wt = Utility.StrtoDouble(t[15].Replace("\"", ""));
                                    workTime += wt;     // 加算
                                }

                                // 普通残業 : CSVデータから取得 2017/09/22
                                double dZan = Utility.StrtoDouble(t[36].Replace("\"", ""));

                                // 早出残業を加算　2017/09/22
                                dZan += Utility.StrtoDouble(t[33].Replace("\"", ""));

                                // 普通残業時間を表示 : 2017/09/22
                                if (dZan > 0)
                                {
                                    rtnArray[2, iDt.Day] = dZan.ToString("#,##0.0");    // 2017/09/22
                                }

                                zanTime += dZan;

                                // 深夜残業 : CSVデータから取得 2017/09/22
                                double dSZan = Utility.StrtoDouble(t[39].Replace("\"", ""));

                                // 深夜残業時間を表示 : 2017/09/22
                                if (dSZan > 0)
                                {
                                    rtnArray[3, iDt.Day] = dSZan.ToString("#,##0.0");   // 2017/09/22
                                }

                                shinyaTime += dSZan;

                                // 遅早時間 2017/09/22
                                double dChisou = Utility.StrtoDouble(t[24].Replace("\"", "")) + Utility.StrtoDouble(t[27].Replace("\"", "")) + Utility.StrtoDouble(t[48].Replace("\"", ""));

                                if (dChisou > 0)
                                {
                                    rtnArray[4, iDt.Day] = dChisou.ToString("#,##0.0");
                                }

                                chisouTime += Utility.StrtoDouble(t[24].Replace("\"", "")) + Utility.StrtoDouble(t[27].Replace("\"", ""));     // 通常遅早時間
                                chisouTimeKyuGyo += Utility.StrtoDouble(t[48].Replace("\"", ""));     // 休業遅早時間
                            }

                            sBushoNum = bCode;
                            sShainNum = t[0].Replace("\"","");
                        }
                    }
                }

                if (rtnArray != null)
                {
                    // 配列からシートへ一括して出力する
                    setXlsData(ref rtnArray, kKbn, sShainNum);
                    rng = oxlsSheet.Range[oxlsSheet.Cells[xR, 4], oxlsSheet.Cells[xR + 3, oxlsMsSheet.UsedRange.Columns.Count]];
                    rng.Value2 = rtnArray;

                    oxlsSheet.Cells[xR + 4, 38] = shinyaTime.ToString("##0.0");     // 深夜残業 : 2017/10/04
                    oxlsSheet.Cells[xR + 4, 36] = tsumikyuuDays.ToString("##0.0");  // 積休使用数：2017/10/04
                    //oxlsSheet.Cells[xR + 4, 40] = shinyaTime.ToString("##0.0");     // 工機夜時間 : 2017/09/23

                    // 1枚目はテンプレートシートなので印刷時には削除する
                    oXls.DisplayAlerts = false;
                    oXlsBook.Sheets[1].Delete();

                    // 1枚目のシートが表示されるようにする
                    oxlsSheet = oXlsBook.Sheets[1];
                    oxlsSheet.Select();

                    //マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    // 確認のためExcelのウィンドウを表示する
                    oXls.Visible = true;

                    // 印刷
                    oXlsBook.PrintOutEx(Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    //oXlsBook.PrintOutEx();
                    //oXlsBook.PrintPreview(true);

                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    //保存処理
                    oXls.DisplayAlerts = false;

                    DialogResult ret;

                    //ダイアログボックスの初期設定
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.Title = "勤怠表";
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    saveFileDialog1.FileName = txtYear.Text + "年" + txtMonth.Text.PadLeft(2, ' ') + "月 勤怠表";
                    saveFileDialog1.Filter = "Microsoft Office Excelファイル(*.xlsx)|*.xlsx|全てのファイル(*.*)|*.*";

                    //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
                    string fileName;
                    ret = saveFileDialog1.ShowDialog();

                    if (ret == System.Windows.Forms.DialogResult.OK)
                    {
                        fileName = saveFileDialog1.FileName;
                        oXlsBook.SaveAs(fileName, Type.Missing, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing,
                                        Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }

                    MessageBox.Show("終了しました", "勤怠表作成");
                }
                else
                {
                    //マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    //保存処理
                    oXls.DisplayAlerts = false;

                    MessageBox.Show("勤怠データはありませんでした", "勤怠表作成");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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

        ///-------------------------------------------------------------------------
        /// <summary>
        ///     勤怠表作成 </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl </param>
        ///-------------------------------------------------------------------------
        private void kintaiSumOrg(sqlControl.DataControl sdCon)
        {
            // エクセルオブジェクト
            Excel.Application oXls = new Excel.Application();
            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(Properties.Settings.Default.xlsKintai,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];
            Excel.Worksheet oxlsMsSheet = (Excel.Worksheet)oXlsBook.Sheets[1]; // テンプレートシート
            oxlsSheet.Select(Type.Missing);

            Excel.Range rng = null;
            object[,] rtnArray = null;

            //当月用勤怠表シート

            // 休日背景色を取得
            rng = oxlsMsSheet.Cells[6, 1];
            double dd = rng.Interior.Color;

            // タイトル
            oxlsMsSheet.Cells[1, 1] = txtYear.Text + "年" + txtMonth.Text.PadLeft(2, ' ') + "月度　勤　怠　表";

            // シートの日付領域を配列に読み込む
            object[,] shDays = null;
            rng = oxlsMsSheet.Range[oxlsMsSheet.Cells[6, 4], oxlsMsSheet.Cells[6, 34]];
            shDays = (object[,])rng.Value2;

            DateTime dt = new DateTime(Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text), 1);
            int fDay = dt.AddMonths(1).AddDays(-1).Day; // 月末日

            // 当月日付を配列にセットする
            for (int i = 1; i <= shDays.Length; i++)
            {
                if (i <= fDay)
                {
                    shDays[1, i] = i.ToString();

                    // 休日の背景色を設定
                    DateTime dt2 = new DateTime(Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text), i);

                    if (dts.休日.Any(a => a.年月日 == dt2))
                    {
                        rng = oxlsMsSheet.Range[oxlsMsSheet.Cells[6, 3 + i], oxlsMsSheet.Cells[41, 3 + i]];
                        rng.Interior.Color = dd;
                    }
                }
                else
                {
                    shDays[1, i] = string.Empty;
                }
            }

            // 日付配列をシートの日付領域に一括貼り付け
            rng = oxlsMsSheet.Range[oxlsMsSheet.Cells[6, 4], oxlsMsSheet.Cells[6, 34]];
            rng.Value2 = shDays;
            
            int xR = 0;
            int pCnt = 1;

            // 雇用区分
            int kKbn = 0;

            try
            {
                ////// 応援移動票と帰宅後勤務も必要。勤務票と合算した専用のデータテーブルが必要　////////////////////////////
                var s = dts.過去勤務票明細.Where(a => a.過去勤務票ヘッダRow != null).OrderBy(a => a.過去勤務票ヘッダRow.部署コード).ThenBy(a => a.社員番号).ThenBy(a => a.過去勤務票ヘッダRow.日);

                //if (comboBox2.SelectedIndex != -1)
                //{
                //    Utility.ComboBumon cmb = (Utility.ComboBumon)comboBox2.SelectedItem;
                //    s = s.Where(a => a.過去勤務票ヘッダRow.部署コード == cmb.code).OrderBy(a => a.過去勤務票ヘッダRow.部署コード).ThenBy(a => a.社員番号).ThenBy(a => a.過去勤務票ヘッダRow.日);
                //}

                string sBushoNum = string.Empty;    // 部署コード
                string sShainNum = string.Empty;    // 社員番号

                int iX = 0;     // 行カウンタ
                int bCnt = 0;   // 部署別カウンタ

                foreach (var t in s)
                {
                    // 社員番号でブレーク
                    if (sShainNum != t.社員番号)
                    {
                        if (sShainNum != string.Empty)
                        {
                            // 配列からシートへ一括して出力する
                            setXlsData(ref rtnArray, kKbn, sShainNum);
                            rng = oxlsSheet.Range[oxlsSheet.Cells[xR, 4], oxlsSheet.Cells[xR + 3, oxlsMsSheet.UsedRange.Columns.Count]];
                            rng.Value2 = rtnArray;

                            oxlsSheet.Cells[xR + 4, 38] = shinyaTime.ToString("##0.0");     // 深夜残業
                            oxlsSheet.Cells[xR + 4, 40] = koukiTime.ToString("##0.0");      // 工機夜時間

                            iX++;   // 行カウンタ加算
                        }

                        // 初期処理、改ページ時または部署コードが変わったら
                        if (iX > 6 || sBushoNum != t.過去勤務票ヘッダRow.部署コード)
                        {
                            // 部署内カウンタ
                            if (iX > 6)
                            {
                                bCnt++; // 同じ部署で次のページ
                            }
                            else
                            {
                                bCnt = 0;   // 次の部署
                            }

                            // テンプレートシートを追加する
                            pCnt++;
                            oxlsMsSheet.Copy(Type.Missing, oXlsBook.Sheets[pCnt - 1]);
                            oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt];

                            // 部署名
                            string buSho = getDepartmentName(t.過去勤務票ヘッダRow.部署コード, sdCon);
                            oxlsSheet.Cells[3, 2] = t.過去勤務票ヘッダRow.部署コード + "  " + buSho;

                            // シート名
                            string sheetName = t.過去勤務票ヘッダRow.部署コード.PadLeft(5, '0') + " " + buSho;
                            if (bCnt > 0)
                            {
                                sheetName += "(" + bCnt.ToString() + ")";
                            }

                            oxlsSheet.Name = sheetName;

                            // 行カウンタ初期化
                            iX = 0;
                        }

                        // 合計欄初期化
                        totalClear();

                        // シートのセルを一括して配列に取得します
                        xR = iX * 5 + 7;
                        rng = oxlsMsSheet.Range[oxlsMsSheet.Cells[xR, 4], oxlsMsSheet.Cells[xR + 3, oxlsMsSheet.UsedRange.Columns.Count]];
                        rtnArray = (object[,])rng.Value2;

                        // 氏名と雇用区分を取得
                        kKbn = 0;
                        string sName = string.Empty;
                        getEmployee(sdCon, t.社員番号, ref sName, ref kKbn);

                        // 社員番号と氏名をシートに貼り付ける
                        oxlsSheet.Cells[xR, 1] = t.社員番号;
                        oxlsSheet.Cells[xR + 1, 1] = sName;
                        oxlsSheet.Cells[xR + 3, 1] = "年休残数：";
                        oxlsSheet.Cells[xR + 4, 1] = "積立残数：";
                    }

                    // 日付毎のデータを配列にセットする

                    // シフトコードを取得する
                    string sftCode = string.Empty;
                    if (t.シフトコード != string.Empty)
                    {
                        sftCode = t.シフトコード.PadLeft(3, '0');
                    }
                    else
                    {
                        sftCode = t.過去勤務票ヘッダRow.シフトコード.ToString().PadLeft(3, '0');
                    }

                    // 休日出勤のとき : 休出・休憩ありを条件に追加 2018/02/05
                    if (sftCode == SHIFT_KYUSHUTSU || sftCode == SHIFT_KYUKEI_KYUSHUTSU)
                    {
                        kyushutsuDays++;    //　日数カウント

                        //double zan = getZanTime(t);

                        double zan = getKumiKoukiTime(sdCon, t, sftCode, "007");
                        if (zan > 0)
                        {
                            rtnArray[2, t.過去勤務票ヘッダRow.日] = zan.ToString("#,##0.0");
                            kyushutsuTime += zan;   // 休出時間加算
                        }

                        rtnArray[1, t.過去勤務票ヘッダRow.日] = "休出";

                        // 休出深夜残業
                        double dSZan = getShinyaTime(sdCon, t, "008");
                        if (dSZan > 0)
                        {
                            rtnArray[3, t.過去勤務票ヘッダRow.日] = dSZan.ToString("#,##0.0");
                            kyusuhtsuShinyaTime += dSZan;   // 休出深夜残業加算
                        }
                    }
                    else if (sftCode == SHIFT_HEIKITAKUGO)
                    {
                        // 平日帰宅後回数カウント
                        yobidashi++;
                    }
                    else if (sftCode == SHIFT_KYUKITAKUGO)
                    {
                        // 休日帰宅後回数カウント
                        yobidashi++;

                        double zan = Utility.StrtoDouble(t.残業時1) * 60 + (Utility.StrtoDouble(t.残業分1) * 60 / 10);
                        zan += Utility.StrtoDouble(t.残業時2) * 60 + (Utility.StrtoDouble(t.残業分2) * 60 / 10);
                        zan = zan / 60;

                        kyushutsuTime += zan;   // 休出時間加算
                    }
                    else
                    {
                        // 事由配列 
                        string[] jArray = { Utility.StrtoInt(t.事由1).ToString(), Utility.StrtoInt(t.事由2).ToString(), Utility.StrtoInt(t.事由3).ToString() };

                        // 事由から出勤欄の記号を求める
                        int jSt = 0;
                        rtnArray[1, t.過去勤務票ヘッダRow.日] = getShukkinKigou(jArray, ref jSt);

                        // 出勤のとき
                        if (jSt == 0)
                        {
                            // 出勤欄表示文字列取得
                            rtnArray[1, t.過去勤務票ヘッダRow.日] = getWorkTime(kKbn, sdCon, t);
                        }

                        // 事由なし、または半休のとき
                        if (jSt == 0 || jSt == 2)
                        {
                            double km = 0;   // 工機、組立夜勤

                            // 組立夜勤１
                            km = getKumiKoukiTime(sdCon, t, sftCode, "010");

                            if (km > 0)
                            {
                                if (jSt == 0)
                                {
                                    rtnArray[1, t.過去勤務票ヘッダRow.日] = "●";
                                }

                                kumitateTime += km;
                            }

                            // 組立夜勤２
                            km = getKumiKoukiTime(sdCon, t, sftCode, "020");

                            if (km > 0)
                            {
                                if (jSt == 0)
                                {
                                    rtnArray[1, t.過去勤務票ヘッダRow.日] = "●";
                                }

                                kumitateTime += km;
                            }

                            // 工機夜勤
                            km = getKumiKoukiTime(sdCon, t, sftCode, "009");

                            if (km > 0)
                            {
                                if (jSt == 0)
                                {
                                    rtnArray[1, t.過去勤務票ヘッダRow.日] = "●";
                                }

                                koukiTime += km;
                            }
                        }

                        // 普通残業
                        double dZan = getKumiKoukiTime(sdCon, t, sftCode, "005");
                        dZan += getKumiKoukiTime(sdCon, t, sftCode, "004");     // 早出残業を加算
                        zanTime += dZan;

                        // 深夜残業
                        double dSZan = getShinyaTime(sdCon, t, "006");
                        shinyaTime += dSZan;

                        // 帰宅後勤務時間
                        foreach (var j in dts.帰宅後勤務.Where(a => a.社員番号 == t.社員番号 && a.年 == t.過去勤務票ヘッダRow.年 && a.月 == t.過去勤務票ヘッダRow.月 && a.日 == t.過去勤務票ヘッダRow.日))
                        {
                            // 帰宅後勤務時間を取得
                            DateTime cTm = DateTime.Today;
                            DateTime sTime = new DateTime(cTm.Year, cTm.Month, cTm.Day, Utility.StrtoInt(j.出勤時), Utility.StrtoInt(j.出勤分), 0);
                            DateTime eTime = new DateTime(cTm.Year, cTm.Month, cTm.Day, Utility.StrtoInt(j.退勤時), Utility.StrtoInt(j.退勤分), 0);

                            double wt = Utility.GetTimeSpan(sTime, eTime).TotalMinutes;
                            wt = (wt - (wt % 30));  // 計算値を30分単位に丸める
                            wt = wt / 60;

                            // 帰宅後深夜残業時間を取得
                            double ktgShinyaZan = getKitakugoTime(sdCon, j, sftCode, "006");
                            dSZan += ktgShinyaZan;
                            shinyaTime += ktgShinyaZan;

                            // 普通残業時間
                            double fZan = wt - ktgShinyaZan;
                            dZan += fZan;
                            zanTime += fZan;
                        }

                        // 普通残業時間を表示
                        if (dZan > 0)
                        {
                            rtnArray[2, t.過去勤務票ヘッダRow.日] = dZan.ToString("#,##0.0");
                        }

                        // 深夜残業時間を表示
                        if (dSZan > 0)
                        {
                            rtnArray[3, t.過去勤務票ヘッダRow.日] = dSZan.ToString("#,##0.0");
                        }

                    }

                    sBushoNum = t.過去勤務票ヘッダRow.部署コード;
                    sShainNum = t.社員番号;
                }

                // 配列からシートへ一括して出力する
                setXlsData(ref rtnArray, kKbn, sShainNum);
                rng = oxlsSheet.Range[oxlsSheet.Cells[xR, 4], oxlsSheet.Cells[xR + 3, oxlsMsSheet.UsedRange.Columns.Count]];
                rng.Value2 = rtnArray;

                oxlsSheet.Cells[xR + 4, 38] = shinyaTime.ToString("##0.0");     // 深夜残業
                oxlsSheet.Cells[xR + 4, 40] = shinyaTime.ToString("##0.0");     // 工機夜時間

                // 1枚目はテンプレートシートなので印刷時には削除する
                oXls.DisplayAlerts = false;
                oXlsBook.Sheets[1].Delete();

                // 1枚目のシートが表示されるようにする
                oxlsSheet = oXlsBook.Sheets[1];
                oxlsSheet.Select();

                //マウスポインタを元に戻す
                this.Cursor = Cursors.Default;

                // 確認のためExcelのウィンドウを表示する
                //oXls.Visible = true;

                // 印刷
                oXlsBook.PrintOutEx();
                //oXlsBook.PrintPreview(true);

                // ウィンドウを非表示にする
                oXls.Visible = false;

                //保存処理
                oXls.DisplayAlerts = false;

                DialogResult ret;

                //ダイアログボックスの初期設定
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Title = "勤怠表";
                saveFileDialog1.OverwritePrompt = true;
                saveFileDialog1.RestoreDirectory = true;
                saveFileDialog1.FileName = txtYear.Text + "年" + txtMonth.Text.PadLeft(2, ' ') + "月 勤怠表";
                saveFileDialog1.Filter = "Microsoft Office Excelファイル(*.xlsx)|*.xlsx|全てのファイル(*.*)|*.*";

                //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
                string fileName;
                ret = saveFileDialog1.ShowDialog();

                if (ret == System.Windows.Forms.DialogResult.OK)
                {
                    fileName = saveFileDialog1.FileName;
                    oXlsBook.SaveAs(fileName, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing,
                                    Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }

                MessageBox.Show("終了しました", "勤怠表作成");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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

        private double getKumiKoukiTime(sqlControl.DataControl sdCon, DataSet1.過去勤務票明細Row t, string sftCode, string itemCode)
        {
            double dSZan = 0;

            // 出勤時間計算
            DateTime cTm = DateTime.Now;
            DateTime szSTime = DateTime.Now;
            DateTime szETime = DateTime.Now;
            DateTime sTime =  DateTime.Now;
            DateTime eTime = DateTime.Now;

            if (t.出勤時 == string.Empty)
            {
                // シフトコードの開始終了時刻を取得する
                getSftTime(sdCon, sftCode.PadLeft(4, '0'), "", ref sTime, ref eTime, ref szSTime, ref szETime);
            }
            else
            {
                sTime = new DateTime(cTm.Year, cTm.Month, cTm.Day, Utility.StrtoInt(t.出勤時), Utility.StrtoInt(t.出勤分), 0);
                eTime = new DateTime(cTm.Year, cTm.Month, cTm.Day, Utility.StrtoInt(t.退勤時), Utility.StrtoInt(t.退勤分), 0);
            }

            // 該当するシフトコードの勤怠時間項目の開始終了時刻を取得する
            if (getKumiKoukiSpan(sdCon, sftCode.PadLeft(4, '0'), itemCode, ref szSTime, ref szETime))
            {
                datetimeAdjust(ref sTime, ref eTime, ref szSTime, ref szETime);
                dSZan = getShinzanTime(sTime, eTime, szSTime, szETime);
            }

            return dSZan;
        }

        private double getKumiKoukiTime(sqlControl.DataControl sdCon, string[] t, string sftCode, string itemCode)
        {
            double dSZan = 0;

            // 出勤時間計算
            DateTime cTm = DateTime.Now;
            DateTime szSTime = DateTime.Now;
            DateTime szETime = DateTime.Now;
            DateTime sTime = DateTime.Now;
            DateTime eTime = DateTime.Now;

            if (t[11].Replace("\"", "") == string.Empty)
            {
                return 0;
            }

            string[] stm = t[11].Replace("\"", "").Split(':');
            string[] etm = t[12].Replace("\"", "").Replace("翌", "").Split(':');

            sTime = new DateTime(cTm.Year, cTm.Month, cTm.Day, Utility.StrtoInt(stm[0]), Utility.StrtoInt(stm[1]), 0);
            eTime = new DateTime(cTm.Year, cTm.Month, cTm.Day, Utility.StrtoInt(etm[0]), Utility.StrtoInt(etm[1]), 0);
            
            // 該当するシフトコードの勤怠時間項目の開始終了時刻を取得する
            if (getKumiKoukiSpan(sdCon, sftCode.PadLeft(4, '0'), itemCode, ref szSTime, ref szETime))
            {
                datetimeAdjust(ref sTime, ref eTime, ref szSTime, ref szETime);
                dSZan = getShinzanTime(sTime, eTime, szSTime, szETime);
            }

            return dSZan;
        }

        private double getKitakugoTime(sqlControl.DataControl sdCon, DataSet1.帰宅後勤務Row t, string sftCode, string itemCode)
        {
            double dSZan = 0;

            // 出勤時間計算
            DateTime cTm = DateTime.Now;
            DateTime szSTime = DateTime.Now;
            DateTime szETime = DateTime.Now;
            DateTime sTime = DateTime.Now;
            DateTime eTime = DateTime.Now;

            sTime = new DateTime(cTm.Year, cTm.Month, cTm.Day, Utility.StrtoInt(t.出勤時), Utility.StrtoInt(t.出勤分), 0);
            eTime = new DateTime(cTm.Year, cTm.Month, cTm.Day, Utility.StrtoInt(t.退勤時), Utility.StrtoInt(t.退勤分), 0);

            // 該当するシフトコードの勤怠時間項目の開始終了時刻を取得する
            if (getKumiKoukiSpan(sdCon, sftCode.PadLeft(4, '0'), itemCode, ref szSTime, ref szETime))
            {
                datetimeAdjust(ref sTime, ref eTime, ref szSTime, ref szETime);
                dSZan = getShinzanTime(sTime, eTime, szSTime, szETime);
            }

            return dSZan;
        }



        private void setXlsData(ref object[,] rtnArray, int kbn, string sShainNum)
        {
            // 配列の合計欄に書き込む
            rtnArray[1, 33] = shukkinDays.ToString("#0.0");
            rtnArray[2, 33] = kyushutsuDays.ToString("#0.0");     // 2017/09/23 コメント化  // 2018/09/18 有効化
            rtnArray[3, 33] = yuukouDays.ToString("#0.0");
            rtnArray[4, 33] = nenkyuuDays.ToString("#0.0");
            //rtnArray[5, 33] = tsumikyuuDays.ToString("#0.0");       // 積休使用数 2017/10/04

            rtnArray[1, 35] = kekkinDays.ToString();

            if (kbn == KBN_PART || kbn == KBN_PART_2 || kbn == KBN_PART_3)
            {
                rtnArray[3, 35] = workTime.ToString("##0.0");
            }

            rtnArray[4, 35] = zanTime.ToString("##0.0");
            //rtnArray[5, 35] = shinyaTime.ToString("##0.0");

            rtnArray[1, 37] = kyushutsuTime.ToString("##0.0");
            rtnArray[2, 37] = kyusuhtsuShinyaTime.ToString("##0.0");
            rtnArray[3, 37] = chisouTime.ToString("##0.0");         // 2017/09/27
            rtnArray[4, 37] = chisouTimeKyuGyo.ToString("##0.0"); // 2017/09/27

            //rtnArray[1, 39] = kumitateTime.ToString("##0.0"); // 2017/09/23
            rtnArray[4, 39] = yobidashi;
        }

        ///---------------------------------------------
        /// <summary>
        ///     合計欄初期化 </summary>
        ///---------------------------------------------
        private void totalClear()
        {
            // 合計欄変数
            shukkinDays = 0;
            kyushutsuDays = 0;
            yuukouDays = 0;
            nenkyuuDays = 0;
            tsumikyuuDays = 0;  // 2017/10/04
            kekkinDays = 0;
            workTime = 0;
            zanTime = 0;
            shinyaTime = 0;
            kyushutsuTime = 0;
            kyusuhtsuShinyaTime = 0;
            koukiTime = 0;
            kumitateTime = 0;
            yobidashi = 0;

            chisouKai = 0;      // 2017/09/23
            chisouTime = 0;     // 2017/09/23
            chisouTimeKyuGyo = 0;     // 2017/09/23
        }


        private double getShinyaTime(sqlControl.DataControl sdCon, DataSet1.過去勤務票明細Row t, string itemCode)
        {
            double dSZan = 0;

            // 出勤時間計算
            DateTime cTm = DateTime.Now;
            DateTime szSTime = DateTime.Now;
            DateTime szETime = DateTime.Now;
            DateTime sTime = new DateTime(cTm.Year, cTm.Month, cTm.Day, Utility.StrtoInt(t.出勤時), Utility.StrtoInt(t.出勤分), 0);
            DateTime eTime = new DateTime(cTm.Year, cTm.Month, cTm.Day, Utility.StrtoInt(t.退勤時), Utility.StrtoInt(t.退勤分), 0);

            // 該当するシフトコードの深夜開始終了時刻を取得する
            getSftShinzanSpan(sdCon, t.過去勤務票ヘッダRow.シフトコード.ToString(), t.シフトコード, itemCode, ref szSTime, ref szETime);
            datetimeAdjust(ref sTime, ref eTime, ref szSTime, ref szETime);
            dSZan = getShinzanTime(sTime, eTime, szSTime, szETime);

            return dSZan;
        }



        ///----------------------------------------------------------------------------
        /// <summary>
        ///     日毎の残業時間を求める </summary>
        /// <param name="t">
        ///     DataSet1.過去勤務票明細Row</param>
        /// <returns>
        ///     残業欄</returns>
        ///----------------------------------------------------------------------------
        private double getZanTime(DataSet1.過去勤務票明細Row t)
        {
            double zan = dts.残業集計.Where(a => a.社員番号 == Utility.StrtoDouble(t.社員番号) &&
                                            a.日 == t.過去勤務票ヘッダRow.日)
                                .Sum(a => a.残業時 * 60 + (a.残業分 * 60 / 10));

            zan = zan / 60;

            return zan;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     出勤日の出勤欄表示文字列を求める </summary>
        /// <param name="kKbn">
        ///     雇用区分</param>
        /// <param name="rtnArray">
        ///     出力配列</param>
        /// <param name="sdCon">
        ///     sqlControl.DataControl</param>
        /// <param name="t">
        ///     DataSet1.過去勤務票明細Row</param>
        /// <returns>
        /// 　　出勤欄印字文字列</returns>
        ///------------------------------------------------------------------------------------
        private string getWorkTime(int kKbn, sqlControl.DataControl sdCon, DataSet1.過去勤務票明細Row t)
        {
            // 出勤のとき
            if (kKbn != KBN_PART && kKbn != KBN_PART_2 && kKbn != KBN_PART_3)
            {
                // パート以外の雇用区分
                if (t.シフト通り == global.FLGON || t.出勤時.Trim() != string.Empty || t.シフトコード != string.Empty)
                {
                    shukkinDays++;  // 出勤日数に加算
                    return "◯";
                }
                else
                {
                    return string.Empty;
                }
            }
            else 
            {
                // 以下、パートタイマー
                // 出勤時間計算
                DateTime cTm = DateTime.Now;
                DateTime sTime = DateTime.Now;
                DateTime eTime = DateTime.Now;
                DateTime restSTime = DateTime.Now;
                DateTime restETime = DateTime.Now;
                double wt = 0;

                // パートタイマー
                if (t.出勤時.Trim() != string.Empty)
                {
                    // シフト通りではないとき（時刻記入のとき）
                    if (DateTime.TryParse(t.出勤時 + ":" + t.出勤分, out cTm))
                    {
                        sTime = cTm;

                        if (DateTime.TryParse(t.退勤時 + ":" + t.退勤分, out cTm))
                        {
                            eTime = cTm;

                            // シフトコードから休憩時刻を求める
                            getSftRestTime(sdCon, t.過去勤務票ヘッダRow.シフトコード.ToString(), t.シフトコード, ref restSTime, ref restETime);
                            
                            // 出勤時間を求める
                            datetimeAdjust(ref sTime, ref eTime, ref restSTime, ref restETime);
                            wt = workTimePart(sTime, eTime, restSTime, restETime);
                            workTime += wt;     // 出勤時間に加算
                            shukkinDays++;      // 出勤日数に加算
                            return wt.ToString("##0.0");
                        }
                        else
                        {
                            return string.Empty;
                        }
                    }
                    else
                    {
                        return string.Empty;
                    }
                }
                else
                {
                    // シフトコードから求める
                    getSftTime(sdCon, t.過去勤務票ヘッダRow.シフトコード.ToString(), t.シフトコード, ref sTime, ref eTime, ref restSTime, ref restETime);
                    wt = workTimePart(sTime, eTime, restSTime, restETime);
                    workTime += wt;     // 出勤時間に加算
                    shukkinDays++;      // 出勤日数に加算
                    return wt.ToString("##0.0");
                }
            }
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     出勤日の出勤欄表示文字列を求める : 2017/09/22</summary>
        /// <param name="kKbn">
        ///     雇用区分</param>
        /// <param name="t">
        ///     CSVデータ配列</param>
        /// <returns>
        /// 　　出勤時間文字列</returns>
        ///------------------------------------------------------------------------------------
        private string getWorkMark(int kKbn, string [] t)
        {
            string rtn = string.Empty;
            string sTM = t[11].Replace("\"", "");    // 開始時刻
            string eTM = t[12].Replace("\"", "");    // 退勤時刻

            // 出勤は雇用区分問わず全員「◯」とする 2017/10/04
            if (sTM != string.Empty || eTM != string.Empty)
            {
                shukkinDays++;  // 出勤日数に加算

                // 工機・組立夜勤のとき：2017/10/04
                if (Utility.StrtoDouble(t[18].Replace("\"", "")) != 0 || Utility.StrtoDouble(t[21].Replace("\"", "")) != 0)
                {
                    rtn = "●";
                }
                else
                {
                    rtn = "◯";
                }
            }
            else
            {
                rtn = string.Empty;
            }

            return rtn;


            // 2017/10/04
            //// 出勤のとき
            //if (kKbn != KBN_PART)
            //{
            //    // パート以外の雇用区分
            //    if (sTM != string.Empty || eTM != string.Empty)
            //    {
            //        shukkinDays++;  // 出勤日数に加算
            //        return "◯";
            //    }
            //    else
            //    {
            //        return string.Empty;
            //    }
            //}
            //else
            //{
            //    // パートタイマー
            //    double wt = Utility.StrtoDouble(t[15].Replace("\"",""));
            //    workTime += wt;     // 出勤時間に加算
            //    shukkinDays++;      // 出勤日数に加算
            //    return wt.ToString("##0.0");
            //}
        }


        ///-------------------------------------------------------------------------
        /// <summary>
        ///     事由から出勤記号を取得する </summary>
        /// <param name="jArray">
        ///     事由配列</param>
        /// <param name="jSt">
        ///     該当ステータス</param>
        /// <returns>
        ///     出勤記号</returns>
        ///-------------------------------------------------------------------------
        private string getShukkinKigou(string[] jArray, ref int jSt)
        {
            string rtn = string.Empty;
            string kk = string.Empty;
            int m = 0;

            for (int i = 0; i < jArray.Length; i++)
            {
                string[] jj = jArray[i].Split(':');

                // 有効事由をカウントします
                if (Utility.StrtoInt(jj[0]) != global.flgOff)
                {
                    m++;
                }

                // 半日事由略語を組み合わせ
                string jName = jj[1].Trim();

                if (jName.Contains('欠'))
                {
                    kk += "欠";
                }
                else if (jName.Contains('半'))
                {
                    kk += "半";
                }
            }

            // 当日出勤日数：2017/11/22
            double toShukkin = 0;

            // 事由記入があるとき : 2017/11/22
            if (m > 0)
            {
                toShukkin = 1;

                for (int i = 0; i < jArray.Length; i++)
                {
                    string[] jj = jArray[i].Split(':');

                    switch (jj[0])
                    {
                        case JIYU_NENKYU:
                            rtn = "年";
                            jSt = 1;
                            nenkyuuDays++;
                            toShukkin = 0;  // 当日出勤日数を０にする 2017/11/22
                            break;

                        case JIYU_ZENHANKYU:
                            rtn = "半休";
                            jSt = 2;
                            nenkyuuDays += 0.5;
                            //shukkinDays += 0.5;
                            toShukkin -= 0.5;   // 当日出勤日数から0.5減らす 2017/11/22
                            break;

                        case JIYU_KOUHANKYU:
                            rtn = "半休";
                            jSt = 2;
                            nenkyuuDays += 0.5;
                            //shukkinDays += 0.5;
                            toShukkin -= 0.5;   // 当日出勤日数から0.5減らす 2017/11/22
                            break;

                        case JIYU_TSUMIKYU:
                            rtn = "積休";
                            jSt = 1;
                            //nenkyuuDays++;
                            tsumikyuuDays++;
                            toShukkin = 0;  // 当日出勤日数を０にする 2017/11/22
                            break;

                        case JIYU_TSUMIZENHAN:
                            rtn = "積半休";
                            jSt = 2;
                            //nenkyuuDays += 0.5;
                            tsumikyuuDays += 0.5;
                            //shukkinDays += 0.5;
                            toShukkin -= 0.5;   // 当日出勤日数から0.5減らす 2017/11/22
                            break;

                        case JIYU_TSUMIKOUHAN:
                            rtn = "積半休";
                            jSt = 2;
                            //nenkyuuDays += 0.5;
                            tsumikyuuDays += 0.5;
                            //shukkinDays += 0.5;
                            toShukkin -= 0.5;   // 当日出勤日数から0.5減らす 2017/11/22
                            break;

                        case JIYU_YUUKOUKYU:
                            rtn = "有公";
                            jSt = 1;
                            yuukouDays++;
                            toShukkin = 0;  // 当日出勤日数を０にする 2017/11/22
                            break;

                        case JIYU_YUUKOUZENHAN:
                            rtn = "有半休";
                            jSt = 2;
                            yuukouDays += 0.5;
                            //shukkinDays += 0.5;
                            toShukkin -= 0.5;   // 当日出勤日数から0.5減らす 2017/11/22
                            break;

                        case JIYU_YUUKOUKOUHAN:
                            rtn = "有半休";
                            jSt = 2;
                            yuukouDays += 0.5;
                            //shukkinDays += 0.5;
                            toShukkin -= 0.5;   // 当日出勤日数から0.5減らす 2017/11/22
                            break;

                        case JIYU_KEKKIN:
                            rtn = "欠勤";
                            jSt = 1;
                            kekkinDays++;
                            toShukkin = 0;  // 当日出勤日数を０にする 2017/11/22
                            break;

                        case JIYU_DAIKYU:
                            rtn = "代休";
                            jSt = 1;
                            toShukkin = 0;  // 当日出勤日数を０にする 2017/11/22
                            break;

                        case JIYU_FURIKYU:
                            rtn = "振休";
                            jSt = 1;
                            toShukkin = 0;  // 当日出勤日数を０にする 2017/11/22
                            break;

                        case JIYU_DOYOTOKKYU:   // 2017/11/21
                            rtn = "土特休";
                            jSt = 1;
                            toShukkin = 0;  // 当日出勤日数を０にする 2017/11/22
                            break;

                        case JIYU_KEKKINZENHAN:   // 2017/11/21
                            rtn = "半欠勤";
                            jSt = 2;
                            kekkinDays += 0.5;
                            //shukkinDays += 0.5;
                            toShukkin -= 0.5;   // 当日出勤日数から0.5減らす 2017/11/22
                            break;

                        case JIYU_KEKKINKOUHAN:   // 2017/11/21
                            rtn = "半欠勤";
                            jSt = 2;
                            kekkinDays += 0.5;
                            //shukkinDays += 0.5;
                            toShukkin -= 0.5;   // 当日出勤日数から0.5減らす 2017/11/22
                            break;

                        case JIYU_YOBIDASHI:   // 呼出は出勤にカウントしない 2018/07/04
                            rtn = "";
                            jSt = 2;
                            //kekkinDays += 0.5;
                            //shukkinDays += 0.5;
                            toShukkin = 0;   // 当日出勤日数を０にする 2018/07/04
                            break;

                        default:
                            break;
                    }
                }

                // 出勤日数を加算 : 2017/11/22
                shukkinDays += toShukkin;
            }

            // 出勤欄の表示文字　2017/11/21
            if (m > 1)
            {
                // 複数事由のとき組み合わせを採用
                rtn = kk;
            }

            return rtn;
        }


        ///-------------------------------------------------------------------
        /// <summary>
        ///     経過時間を求める </summary>
        /// <param name="sTime">
        ///     出勤時刻</param>
        /// <param name="eTime">
        ///     退勤時刻</param>
        /// <returns>
        ///     出勤時間</returns>
        /// <param name="restSTime">
        ///     休憩開始時刻</param>
        /// <param name="restETime">
        ///     休憩終了時刻</param>
        ///-------------------------------------------------------------------
        private double workTimePart(DateTime sTime, DateTime eTime, DateTime restSTime, DateTime restETime )
        {
            double wt = Utility.GetTimeSpan(sTime, eTime).TotalMinutes;
            double restTime = Utility.GetTimeSpan(restSTime, restETime).TotalMinutes;

            // 勤務時間中に休憩時刻があるとき
            if (sTime < restSTime && restETime < eTime)
            {
                // 休憩時間を差し引く
                wt = wt - restTime;
            }

            wt = (wt - (wt % 30));  // 計算値を30分単位に丸める
            wt = wt / 60;

            return wt;
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     深夜勤務時間を求める </summary>
        /// <param name="sTime">
        ///     出勤時刻</param>
        /// <param name="eTime">
        ///     退勤時刻</param>
        /// <param name="szSTime">
        ///     深夜開始時刻</param>
        /// <param name="szETime">
        ///     深夜終了時刻</param>
        /// <returns>
        ///     深夜勤務時間</returns>
        ///-------------------------------------------------------------------
        private double getShinzanTime(DateTime sTime, DateTime eTime, DateTime szSTime, DateTime szETime)
        {
            double wt = 0;
            DateTime fromTm = szSTime;
            DateTime toTm = szETime;

            //// 深夜残業開始前に終了したとき
            //if (eTime <= szSTime)
            //{
            //    return 0;
            //}

            //// 深夜勤務時間を算出
            //if (szSTime < eTime)
            //{
            //    if (eTime <= szETime)
            //    {
            //        /* 勤務終了時刻が深夜残業終了時刻以前のとき
            //         * 22:00～05:00 で勤務が08:00～23:00のときなど */
            //        wt = Utility.GetTimeSpan(szSTime, eTime).TotalMinutes;
            //    }
            //    else
            //    {
            //        /* 勤務終了時刻が深夜残業終了時刻以降のとき
            //         * 22:00～05:00 で勤務が08:00～翌日5:30のときなど */
            //        wt = Utility.GetTimeSpan(szSTime, szETime).TotalMinutes;
            //    }
            //}


            // 時間帯開始前に終了または時間帯終了後に開始のとき
            if (eTime <= szSTime || sTime >= szETime)
            {
                return 0;
            }

            // 開始時刻 8:00 22:00 22:30 22:00
            if (sTime < szSTime)
            {
                fromTm = szSTime;
            }
            else
            {
                fromTm = sTime;
            }

            // 終了時刻
            if (eTime <= szETime)
            {
                toTm = eTime;
            }
            else
            {
                toTm = szETime;
            }
            
            wt = Utility.GetTimeSpan(fromTm, toTm).TotalMinutes;
            
            //if (szSTime < eTime)
            //{
            //    if (eTime <= szETime)
            //    {
            //        /* 勤務終了時刻が深夜残業終了時刻以前のとき
            //         * 22:00～05:00 で勤務が08:00～23:00のときなど */
            //        wt = Utility.GetTimeSpan(szSTime, eTime).TotalMinutes;
            //    }
            //    else
            //    {
            //        /* 勤務終了時刻が深夜残業終了時刻以降のとき
            //         * 22:00～05:00 で勤務が08:00～翌日5:30のときなど */
            //        wt = Utility.GetTimeSpan(szSTime, szETime).TotalMinutes;
            //    }
            //}

            wt = (wt - (wt % 30));  // 計算値を30分単位に丸める
            wt = wt / 60;

            return wt;
        }

        ///--------------------------------------------------------------------------
        /// <summary>
        ///     開始終了時間に当日、翌日の日付情報を与える </summary>
        /// <param name="sTime">
        ///     開始時刻</param>
        /// <param name="eTime">
        ///     終了時刻</param>
        /// <param name="restSTime">
        ///     時間帯開始時刻</param>
        /// <param name="restETime">
        ///     時間帯終了時刻</param>
        ///--------------------------------------------------------------------------
        private void datetimeAdjust(ref DateTime sTime, ref DateTime eTime, ref DateTime restSTime, ref DateTime restETime)
        {
            int st = sTime.Hour * 100 + sTime.Minute;
            int et = eTime.Hour * 100 + eTime.Minute;
            int rSt = restSTime.Hour * 100 + restSTime.Minute;
            int rEt = restETime.Hour * 100 + restETime.Minute;

            DateTime nDt = DateTime.Now;
            sTime = new DateTime(nDt.Year, nDt.Month, nDt.Day, sTime.Hour, sTime.Minute, 0); // 今日の日付とする

            // 開始時刻より終了時刻が小さいときは翌日
            if (st > et)
            {
                eTime = new DateTime(nDt.AddDays(1).Year, nDt.AddDays(1).Month, nDt.AddDays(1).Day, eTime.Hour, eTime.Minute, 0);
            }
            else
            {
                eTime = new DateTime(nDt.Year, nDt.Month, nDt.Day, eTime.Hour, eTime.Minute, 0);
            }

            //// 開始時刻より時間帯開始時刻が小さいときは翌日
            //if (st > rSt)
            //{
            //    restSTime = new DateTime(nDt.AddDays(1).Year, nDt.AddDays(1).Month, nDt.AddDays(1).Day, restSTime.Hour, restSTime.Minute, 0);
            //}
            //else
            //{
            //    restSTime = new DateTime(nDt.Year, nDt.Month, nDt.Day, restSTime.Hour, restSTime.Minute, 0);
            //}

            //// 休憩開始時刻より時間帯終了時刻が小さいときは翌日
            //if (rSt > rEt)
            //{
            //    restETime = new DateTime(restSTime.AddDays(1).Year, restSTime.AddDays(1).Month, restSTime.AddDays(1).Day, restETime.Hour, restETime.Minute, 0);
            //}
            //else
            //{
            //    restETime = new DateTime(restSTime.Year, restSTime.Month, restSTime.Day, restETime.Hour, restETime.Minute, 0);
            //}

            // 奉行マスタから取得した日付情報の日付が[2]当日、[3]翌日
            if (restSTime.Day == 3)
            {
                restSTime = new DateTime(nDt.AddDays(1).Year, nDt.AddDays(1).Month, nDt.AddDays(1).Day, restSTime.Hour, restSTime.Minute, 0);
            }
            else
            {
                restSTime = new DateTime(nDt.Year, nDt.Month, nDt.Day, restSTime.Hour, restSTime.Minute, 0);
            }
            
            if (restETime.Day == 3)
            {
                restETime = new DateTime(nDt.AddDays(1).Year, nDt.AddDays(1).Month, nDt.AddDays(1).Day, restETime.Hour, restETime.Minute, 0);
            }
            else
            {
                restETime = new DateTime(nDt.Year, nDt.Month, nDt.Day, restETime.Hour, restETime.Minute, 0);
            }
        }
        

        ///-----------------------------------------------------------------------------
        /// <summary>
        ///     奉行の社員マスターより社員名と雇用区分を取得する </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl</param>
        /// <param name="sNum">
        ///     社員番号</param>
        /// <param name="sName">
        ///     社員名</param>
        /// <param name="sKbn">
        ///     雇用区分</param>
        ///-----------------------------------------------------------------------------
        private void getEmployee(sqlControl.DataControl sdCon, string sNum, ref string sName, ref int sKbn)
        {
            SqlDataReader dR = null;

            try
            {
                // 該当者の雇用区分を取得する
                StringBuilder sb = new StringBuilder();
                sb.Clear();

                //sb.Append("select EmployeeNo,Name,EmploymentDivisionID from tbEmployeeBase ");
                //sb.Append("where EmployeeNo = '" + sNum.PadLeft(10, '0') + "'");

                // 退職者の氏名も取得する：2017/10/05
                //sb.Append(" and BeOnTheRegisterDivisionID != 9");
                
                sb.Append("select EmployeeNo,Name, tbHR_DivisionCategory.CategoryCode ");
                sb.Append("from tbEmployeeBase inner join tbHR_DivisionCategory ");
                sb.Append("on tbEmployeeBase.EmploymentDivisionID = tbHR_DivisionCategory.CategoryID ");
                sb.Append("where EmployeeNo = '" + sNum.PadLeft(10, '0') + "'");
                
                dR = sdCon.free_dsReader(sb.ToString());

                while (dR.Read())
                {
                    sName = dR["Name"].ToString();
                    sKbn = Utility.StrtoInt(dR["CategoryCode"].ToString());
                    break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                dR.Close();
            }
        }

        ///--------------------------------------------------------------------
        /// <summary>
        ///     奉行マスターから部署名を取得する </summary>
        /// <param name="_dbName">
        ///     データベース名</param>
        /// <param name="dCode">
        ///     部署コード</param>
        /// <param name="sdCon">
        ///     sqlControl.DataControl</param>
        /// <returns>
        ///     部署名 </returns>
        ///--------------------------------------------------------------------
        private string getDepartmentName(string dCode, sqlControl.DataControl sdCon)
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

            return dName;
        }

        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     対象のシフトコードの開始時刻・終了時刻を取得する </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl</param>
        /// <param name="sSftCode">
        ///     標準シフトコード</param>
        /// <param name="hSftCode">
        ///     変更シフトコード</param>
        /// <param name="dtStart">
        ///     開始時刻</param>
        /// <param name="dtEnd">
        ///     終了時刻</param>
        /// <param name="restDtStart">
        ///     休憩開始時刻</param>
        /// <param name="restDtEnd">
        ///     休憩終了時刻</param>
        ///----------------------------------------------------------------------------------
        private void getSftTime(sqlControl.DataControl sdCon, string sSftCode, string hSftCode, ref DateTime dtStart, ref DateTime dtEnd, ref DateTime restDtStart, ref DateTime restDtEnd)
        {
            bool bn = false;
            DateTime dtChange = DateTime.Now;       // 日替わり時刻

            // 対象のシフトコードを取得する
            string sftCode = string.Empty;

            if (hSftCode != string.Empty)
            {
                // 変更シフトコードあり
                sftCode = hSftCode.PadLeft(4, '0');
            }
            else if (sSftCode.ToString() != string.Empty)
            {
                // 標準シフトコード
                sftCode = sSftCode.ToString().PadLeft(4, '0');
            }

            // 有効なシフトコードが存在しないとき
            if (sftCode == string.Empty)
            {
                return;
            }

            // 開始時刻、終了時刻を取得する
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select tbLaborSystem.LaborSystemCode, LaborSystemName, tbLaborSystem.LatterHalfStartTime, tbLaborSystem.FirstHalfEndTime,");
            sb.Append("tbLaborSystem.DayChangeTime,a.StartTime,a.EndTime ");
            sb.Append("from tbLaborSystem left join ");
            sb.Append("(select * from tbLaborTimeSpanRule where LaborTimeItemID = 1) as a ");
            sb.Append("on tbLaborSystem.LaborSystemID = a.LaborSystemID ");
            sb.Append("where tbLaborSystem.LaborSystemCode = '").Append(sftCode).Append("'");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                if (!(dR["StartTime"] is DBNull))
                {
                    bn = true;
                    dtStart = (DateTime)dR["StartTime"];
                }

                if (!(dR["EndTime"] is DBNull))
                {
                    bn = true;
                    dtEnd = (DateTime)dR["EndTime"];
                    dtChange = (DateTime)dR["DayChangeTime"];
                }
            }

            dR.Close();

            // 休憩開始時刻、休憩終了時刻を求める
            sb.Clear();            
            sb.Append("SELECT tbLaborSystem.LaborSystemID,LaborSystemCode,tbRestTimeSpanRule.StartTime,");
            sb.Append("tbRestTimeSpanRule.EndTime ");
            sb.Append("from tbLaborSystem inner join tbRestTimeSpanRule ");
            sb.Append("on tbLaborSystem.LaborSystemID = tbRestTimeSpanRule.LaborSystemID ");
            sb.Append("where LaborSystemCode = '").Append(sftCode).Append("'");

            dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                restDtStart = (DateTime)dR["StartTime"];
                restDtEnd = (DateTime)dR["EndTime"];
            }

            dR.Close();
        }

        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     対象のシフトコードの休憩開始時刻・終了時刻を取得する </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl</param>
        /// <param name="sSftCode">
        ///     標準シフトコード</param>
        /// <param name="hSftCode">
        ///     変更シフトコード</param>
        /// <param name="restDtStart">
        ///     休憩開始時刻</param>
        /// <param name="restDtEnd">
        ///     休憩終了時刻</param>
        ///----------------------------------------------------------------------------------
        private void getSftRestTime(sqlControl.DataControl sdCon, string sSftCode, string hSftCode, ref DateTime restDtStart, ref DateTime restDtEnd)
        {
            bool bn = false;
            DateTime dtChange = DateTime.Now;       // 日替わり時刻

            // 対象のシフトコードを取得する
            string sftCode = string.Empty;

            if (hSftCode != string.Empty)
            {
                // 変更シフトコードあり
                sftCode = hSftCode.PadLeft(4, '0');
            }
            else if (sSftCode.ToString() != string.Empty)
            {
                // 標準シフトコード
                sftCode = sSftCode.ToString().PadLeft(4, '0');
            }

            // 有効なシフトコードが存在しないとき
            if (sftCode == string.Empty)
            {
                return;
            }

            StringBuilder sb = new StringBuilder();

            // 休憩開始時刻、休憩終了時刻を求める
            sb.Clear();
            sb.Append("SELECT tbLaborSystem.LaborSystemID,LaborSystemCode,tbRestTimeSpanRule.StartTime,");
            sb.Append("tbRestTimeSpanRule.EndTime ");
            sb.Append("from tbLaborSystem inner join tbRestTimeSpanRule ");
            sb.Append("on tbLaborSystem.LaborSystemID = tbRestTimeSpanRule.LaborSystemID ");
            sb.Append("where LaborSystemCode = '").Append(sftCode).Append("'");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                restDtStart = (DateTime)dR["StartTime"];
                restDtEnd = (DateTime)dR["EndTime"];
            }

            dR.Close();
        }


        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     対象のシフトコードの組立夜勤1の開始時刻・終了時刻を取得する </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl</param>
        /// <param name="sSftCode">
        ///     シフトコード</param>
        /// <param name="szStart">
        ///     深夜残業開始時刻</param>
        /// <param name="szEnd">
        ///     深夜残業終了時刻</param>
        /// <returns>
        ///     true:時間帯あり、false:時間帯なし</returns>
        ///----------------------------------------------------------------------------------
        private bool getKumiKoukiSpan(sqlControl.DataControl sdCon, string sftCode, string itemCode, ref DateTime szStart, ref DateTime szEnd)
        {
            bool rtn = false;

            StringBuilder sb = new StringBuilder();

            // 開始時刻、終了時刻を求める
            sb.Clear();
            sb.Append("select tbLaborSystem.LaborSystemCode, LaborSystemName,a.StartTime,a.EndTime ");
            sb.Append("from tbLaborSystem left join ");
            sb.Append("(select tbLaborTimeSpanRule.* from tbLaborTimeSpanRule inner join tbLaborTimeItem ");
            sb.Append("on tbLaborTimeSpanRule.LaborTimeItemID = tbLaborTimeItem.LaborTimeItemID ");
            sb.Append("where tbLaborTimeItem.LaborTimeItemCode = '").Append(itemCode).Append("') as a ");
            sb.Append("on tbLaborSystem.LaborSystemID = a.LaborSystemID ");
            sb.Append("where tbLaborSystem.LaborSystemCode = '").Append(sftCode).Append("'");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                if (!(dR["StartTime"] is DBNull))
                {
                    szStart = (DateTime)dR["StartTime"];
                    szEnd = (DateTime)dR["EndTime"];
                    rtn = true;
                }
            }

            dR.Close();

            return rtn;
        }

        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     対象のシフトコードの深夜残業の開始時刻・終了時刻を取得する </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl</param>
        /// <param name="sSftCode">
        ///     標準シフトコード</param>
        /// <param name="hSftCode">
        ///     変更シフトコード</param>
        /// <param name="itemCode">
        ///     勤怠時間項目コード</param>
        /// <param name="szStart">
        ///     深夜残業開始時刻</param>
        /// <param name="szEnd">
        ///     深夜残業終了時刻</param>
        ///----------------------------------------------------------------------------------
        private void getSftShinzanSpan(sqlControl.DataControl sdCon, string sSftCode, string hSftCode, string itemCode, ref DateTime szStart, ref DateTime szEnd)
        {          
            // 対象のシフトコードを取得する
            string sftCode = string.Empty;

            if (hSftCode != string.Empty)
            {
                // 変更シフトコードあり
                sftCode = hSftCode.PadLeft(4, '0');
            }
            else if (sSftCode.ToString() != string.Empty)
            {
                // 標準シフトコード
                sftCode = sSftCode.ToString().PadLeft(4, '0');
            }

            // 有効なシフトコードが存在しないとき
            if (sftCode == string.Empty)
            {
                return;
            }

            StringBuilder sb = new StringBuilder();

            // 深夜残業開始時刻、深夜残業終了時刻を求める
            sb.Clear();
            sb.Append("select tbLaborSystem.LaborSystemCode, LaborSystemName,a.StartTime,a.EndTime ");
            sb.Append("from tbLaborSystem left join ");
            sb.Append("(select tbLaborTimeSpanRule.* from tbLaborTimeSpanRule inner join tbLaborTimeItem ");
            sb.Append("on tbLaborTimeSpanRule.LaborTimeItemID = tbLaborTimeItem.LaborTimeItemID ");
            sb.Append("where tbLaborTimeItem.LaborTimeItemCode = '").Append(itemCode).Append("') as a ");
            sb.Append("on tbLaborSystem.LaborSystemID = a.LaborSystemID ");
            sb.Append("where tbLaborSystem.LaborSystemCode = '").Append(sftCode).Append("'");
            
            //sb.Clear();
            //sb.Append("select tbLaborSystem.LaborSystemCode, LaborSystemName, ");
            //sb.Append("a.StartTime,a.EndTime ");
            //sb.Append("from tbLaborSystem left join ");
            //sb.Append("(select * from tbLaborTimeSpanRule where LaborTimeItemID = 6) as a ");
            //sb.Append("on tbLaborSystem.LaborSystemID = a.LaborSystemID ");
            //sb.Append("where tbLaborSystem.LaborSystemCode = '").Append(sftCode).Append("'");
            
            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                if (!(dR["StartTime"] is DBNull))
                {
                    szStart = (DateTime)dR["StartTime"];
                    szEnd = (DateTime)dR["EndTime"];
                }
            }

            dR.Close();
        }

        private void txtMonth_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void frmKintaiRep_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            departmentShow();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "勤怠データ選択";
            openFileDialog1.FileName = string.Empty;
            openFileDialog1.Filter = "ＣＳＶファイル(*.csv)|*.csv|全てのファイル(*.*)|*.*";

            //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
            string fileName;
            DialogResult ret = openFileDialog1.ShowDialog();

            if (ret == System.Windows.Forms.DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
                label6.Text = openFileDialog1.FileName;

                // 勤怠データ配列読み込み
                workArray = System.IO.File.ReadAllLines(label6.Text, Encoding.Default);

                int yy = 0, mm = 0, days = 0;
                getWorkDateDays(out yy, out mm, out days);

                txtYear.Text = yy.ToString();
                txtMonth.Text = mm.ToString();
            }
            else
            {
                fileName = string.Empty;
            }
        }

        private void impWorkData(string files)
        {
            // CSVファイルインポート
            var s = System.IO.File.ReadAllLines(files, Encoding.Default);

            foreach (var stBuffer in s)
            {
                // カンマ区切りで分割して配列に格納する
                string[] stCSV = stBuffer.Split(',');

            }
        }

        ///-----------------------------------------------------------------
        /// <summary>
        ///     CSVデータの対象年月と最大日付を取得する </summary>
        /// <param name="_yy">
        ///     年</param>
        /// <param name="_mm">
        ///     月</param>
        /// <param name="_days">
        ///     最大日付</param>
        /// <returns>
        ///     true, false</returns>
        ///-----------------------------------------------------------------
        private bool getWorkDateDays(out int _yy, out int _mm, out int _days)
        {
            bool rtn = false;
            _yy = 0;
            _mm = 0;
            _days = 0;

            string sNum = string.Empty;

            foreach (var item in workArray)
            {
                string[] t = item.Split(',');

                // 社員番号取得
                string strSnum = t[0].Replace("\"", "");

                // 1行目見出し行は読み飛ばす
                if (strSnum == "社員番号")
                {
                    continue;
                }

                if (sNum != string.Empty && sNum != strSnum)
                {
                    rtn = true;
                    break;
                }

                sNum = strSnum;

                // 日付情報を取得する
                string strDate = t[2].Replace("\"", "");
                //string f = "ggyy年MM月dd日";
                System.Globalization.CultureInfo ci = new System.Globalization.CultureInfo("ja-JP");
                DateTime iDt = DateTime.Parse(strDate, ci, System.Globalization.DateTimeStyles.AssumeLocal);

                _yy = iDt.Year;
                _mm = iDt.Month;
                _days = iDt.Day;
            }

            return rtn;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "年休・積休残データ選択";
            openFileDialog1.FileName = string.Empty;
            openFileDialog1.Filter = "ＣＳＶファイル(*.csv)|*.csv|全てのファイル(*.*)|*.*";

            //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
            string fileName;
            DialogResult ret = openFileDialog1.ShowDialog();

            if (ret == System.Windows.Forms.DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
                label1.Text = openFileDialog1.FileName;

                // 勤怠データ配列読み込み
                nenkyuArray = System.IO.File.ReadAllLines(label1.Text, Encoding.Default);
            }
            else
            {
                fileName = string.Empty;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
            {
                button2.Enabled = true;
                button3.Enabled = true;
            }
            else
            {
                button2.Enabled = false;
                button3.Enabled = true;
            }
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     年休残・積休残日数をCSVデータから取得する </summary>
        /// <param name="sNum">
        ///     社員番号</param>
        /// <param name="nen">
        ///     年休残</param>
        /// <param name="tsumi">
        ///     積休残</param>
        ///---------------------------------------------------------------------
        private void getNenkyuData(int sNum, out decimal nen, out decimal tsumi)
        {
            nen = 0;
            tsumi = 0;

            for (int i = 1; i < nenkyuArray.Length; i++)
            {
                string[] t = nenkyuArray[i].Split(',');

                if (sNum != Utility.StrtoInt(t[0].Replace("\"", "")))
                {
                    continue;
                }

                tsumi = (decimal)Utility.StrtoDouble(t[3].Replace("\"", ""));
                nen = (decimal)Utility.StrtoDouble(t[4].Replace("\"", ""));
                break;
            }
        }
    }
}
