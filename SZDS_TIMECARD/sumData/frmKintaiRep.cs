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
    public partial class frmKintaiRep : Form
    {
        public frmKintaiRep(string dbName)
        {
            InitializeComponent();

            _dbName = dbName;
        }

        string _dbName = string.Empty;

        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.過去勤務票ヘッダTableAdapter hAdp = new DataSet1TableAdapters.過去勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.過去勤務票明細TableAdapter iAdp = new DataSet1TableAdapters.過去勤務票明細TableAdapter();
        DataSet1TableAdapters.休日TableAdapter dAdp = new DataSet1TableAdapters.休日TableAdapter();
        DataSet1TableAdapters.残業集計TableAdapter zAdp = new DataSet1TableAdapters.残業集計TableAdapter();
        DataSet1TableAdapters.帰宅後勤務TableAdapter kAdp = new DataSet1TableAdapters.帰宅後勤務TableAdapter();

        // 雇用区分
        const int KBN_SHAIN = 2;
        const int KBN_PART = 6;

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

        // 勤務体系コード
        const string SHIFT_KYUSHUTSU = "031";
        const string SHIFT_KYUKEI_KYUSHUTSU = "032";
        const string SHIFT_HEIKITAKUGO = "041";
        const string SHIFT_KYUKITAKUGO = "042";

        // 合計欄変数
        double shukkinDays = 0;
        double kyushutsuDays = 0;
        double yuukouDays = 0;
        double nenkyuuDays = 0;
        double kekkinDays = 0;
        double workTime = 0;
        double zanTime = 0;
        double shinyaTime = 0;
        double kyushutsuTime = 0;
        double kyusuhtsuShinyaTime = 0;
        double koukiTime = 0;
        double kumitateTime = 0;
        double yobidashi = 0;

        // カラム定義
        private string ColChk = "c0";
        private string ColSz = "c1";
        private string ColSznm = "c2";
        private string ColCode = "c3";
        private string ColNin = "c4";
        private string ColID = "c5";

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
            gridViewShowBusho(sc, dataGridView1);
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

        private void gridViewShowBusho(string sConnect, DataGridView dg)
        {
            sqlControl.DataControl sdCon = new Common.sqlControl.DataControl(sConnect);

            hAdp.FillByYYMM(dts.過去勤務票ヘッダ, Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text));

            if (dts.過去勤務票ヘッダ.Count == 0)
            {
                MessageBox.Show("該当年月に勤怠データが存在しません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                linkLabel3.Enabled = false;
                linkLblOn.Enabled = false;
                linkLblOff.Enabled = false;
                dg.RowCount = 0;
                return;
            }
                            
            try
            {
                var s = dts.過去勤務票ヘッダ
                    //.Where(a => a.年 == Utility.StrtoInt(txtYear.Text) && a.月 == Utility.StrtoInt(txtMonth.Text))
                    .Select(a => new
                {
                    y = a.部署コード,
                }).Distinct()
                .OrderBy(a => a.y);

                int iX = 0;
                dg.RowCount = 0;

                foreach (var t in s)
                {
                    //データグリッドにデータを表示する
                    dg.Rows.Add();

                    dg[ColChk, iX].Value = false;
                    dg[ColCode, iX].Value = t.y;
                    dg[ColSznm, iX].Value = getDepartmentName(t.y, sdCon);

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
                sdCon.Close();
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
            //int nin = 0;
            //SqlDataReader cDr = sdCon.free_dsReader(strCode);
            //while (cDr.Read())
            //{
            //    nin = Utility.StrtoInt(cDr["cnt"].ToString());
            //    break;
            //}

            int nin = 0;
            SqlDataReader cDr = sdCon.free_dsReader(Utility.getEmployeeCount(strCode, sDt)); // 基準年月日 2017/09/28
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

            if (MessageBox.Show("勤怠表を発行します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            this.Cursor = Cursors.WaitCursor;

            // 奉行データベース接続
            string sc = null;
            sqlControl.DataControl sdCon = null;

            try
            {
                int yy = Utility.StrtoInt(txtYear.Text);
                int mm = Utility.StrtoInt(txtMonth.Text);
                
                hAdp.FillByYYMM(dts.過去勤務票ヘッダ, Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text));
                
                if (dts.過去勤務票ヘッダ.Count == 0)
                {
                    MessageBox.Show("対象データがありません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                                
                iAdp.Fill(dts.過去勤務票明細);
                dAdp.Fill(dts.休日);
                zAdp.Fill(dts.残業集計, yy, mm, yy, mm, yy, mm, yy, mm, yy, mm, yy, mm);
                kAdp.FillByYYMM(dts.帰宅後勤務, yy, mm);

                // 奉行マスター接続
                sc = sqlControl.obcConnectSting.get(_dbName);
                sdCon = new Common.sqlControl.DataControl(sc);

                // 勤怠表作成
                kintaiSum(sdCon);
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

                    // 部署・社員情報データリーダー取得
                    SqlDataReader dr = sdCon.free_dsReader(Utility.getEmployeeOrder(strCode, DateTime.Today));

                    string[] saArray = null;

                    int iSa = 0;

                    // 並び替えた社員番号から配列を作成する
                    while (dr.Read())
                    {
                        // 社員番号
                        string sCode = Utility.StrtoInt(dr["EmployeeNo"].ToString()).ToString();

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
                        var s = dts.過去勤務票明細.Where(a => a.過去勤務票ヘッダRow != null && a.社員番号 == saArray[i]).OrderBy(a => a.過去勤務票ヘッダRow.部署コード).ThenBy(a => a.過去勤務票ヘッダRow.日);

                        // 勤怠データがないときはネグる
                        if (s.Count() == 0)
                        {
                            continue;
                        }

                        // 過去勤務票明細データを日付順に読む
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
                                //if (iX > 6 || sBushoNum != t.過去勤務票ヘッダRow.部署コード)
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
                                    //string buSho = getDepartmentName(t.過去勤務票ヘッダRow.部署コード, sdCon);
                                    //oxlsSheet.Cells[3, 2] = t.過去勤務票ヘッダRow.部署コード + "  " + buSho;
                                    string buSho = getDepartmentName(bCode, sdCon);
                                    oxlsSheet.Cells[3, 2] = bCode + "  " + buSho;

                                    // シート名
                                    //string sheetName = t.過去勤務票ヘッダRow.部署コード.PadLeft(5, '0') + " " + buSho;
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

                            //sBushoNum = t.過去勤務票ヘッダRow.部署コード;
                            sBushoNum = bCode;
                            sShainNum = t.社員番号;
                        }
                    }
                }

                if (rtnArray != null)
                {
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
            rtnArray[2, 33] = kyushutsuDays.ToString("#0.0");
            rtnArray[3, 33] = yuukouDays.ToString("#0.0");
            rtnArray[4, 33] = nenkyuuDays.ToString("#0.0");

            rtnArray[1, 35] = kekkinDays.ToString();

            if (kbn == KBN_PART)
            {
                rtnArray[3, 35] = workTime.ToString("##0.0");
            }

            rtnArray[4, 35] = zanTime.ToString("##0.0");
            //rtnArray[5, 35] = shinyaTime.ToString("##0.0");

            rtnArray[1, 37] = kyushutsuTime.ToString("##0.0");
            rtnArray[2, 37] = kyusuhtsuShinyaTime.ToString("##0.0");

            rtnArray[1, 39] = kumitateTime.ToString("##0.0");
            rtnArray[4, 39] = yobidashi + dts.帰宅後勤務.Count(a => a.社員番号 == sShainNum);
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
            kekkinDays = 0;
            workTime = 0;
            zanTime = 0;
            shinyaTime = 0;
            kyushutsuTime = 0;
            kyusuhtsuShinyaTime = 0;
            koukiTime = 0;
            kumitateTime = 0;
            yobidashi = 0;
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
            if (kKbn != KBN_PART)
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

            for (int i = 0; i < jArray.Length; i++)
            {
                switch (jArray[i])
                {
                    case JIYU_NENKYU:
                        rtn= "年";
                        jSt = 1;
                        nenkyuuDays++;
                        break;

                    case JIYU_ZENHANKYU:
                        rtn = "半休";
                        jSt = 2;
                        nenkyuuDays += 0.5;
                        shukkinDays += 0.5;
                        break;

                    case JIYU_KOUHANKYU:
                        rtn = "半休";
                        jSt = 2;
                        nenkyuuDays += 0.5;
                        shukkinDays += 0.5;
                        break;

                    case JIYU_TSUMIKYU:
                        rtn = "積休";
                        jSt = 1;
                        nenkyuuDays++;
                        break;

                    case JIYU_TSUMIZENHAN:
                        rtn = "半休";
                        jSt = 2;
                        nenkyuuDays += 0.5;
                        shukkinDays += 0.5;
                        break;

                    case JIYU_TSUMIKOUHAN:
                        rtn = "半休";
                        jSt = 2;
                        nenkyuuDays += 0.5;
                        shukkinDays += 0.5;
                        break;

                    case JIYU_YUUKOUKYU:
                        rtn = "有公";
                        jSt = 1;
                        yuukouDays++;
                        break;

                    case JIYU_YUUKOUZENHAN:
                        rtn = "半休";
                        jSt = 2;
                        nenkyuuDays += 0.5;
                        shukkinDays += 0.5;
                        break;

                    case JIYU_YUUKOUKOUHAN:
                        rtn = "半休";
                        jSt = 2;
                        nenkyuuDays += 0.5;
                        shukkinDays += 0.5;
                        break;

                    case JIYU_KEKKIN:
                        rtn= "欠勤";
                        jSt = 1;
                        kekkinDays++;
                        break;

                    case JIYU_DAIKYU:
                        rtn = "代休";
                        jSt = 1;
                        break;

                    case JIYU_FURIKYU:
                        rtn = "振休";
                        jSt = 1;
                        break;

                    default:
                        break;
                }
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
                sb.Append("select EmployeeNo,Name,EmploymentDivisionID from tbEmployeeBase ");
                sb.Append("where EmployeeNo = '" + sNum.PadLeft(10, '0') + "'");
                sb.Append(" and BeOnTheRegisterDivisionID != 9");

                dR = sdCon.free_dsReader(sb.ToString());

                while (dR.Read())
                {
                    sName = dR["Name"].ToString();
                    sKbn = Utility.StrtoInt(dR["EmploymentDivisionID"].ToString());
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

    }
}
