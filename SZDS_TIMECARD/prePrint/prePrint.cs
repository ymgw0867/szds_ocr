﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SZDS_TIMECARD.Common;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;

namespace SZDS_TIMECARD.prePrint
{
    public partial class prePrint : Form
    {
        public prePrint(string dbName, string comName)
        {
            InitializeComponent();

            _dbName = dbName;
            _comMane = comName;

            adp.Fill(dts.休日);
        }

        string _dbName;     // 会社領域データベース名
        string _comMane;    // 選択した会社名

        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.休日TableAdapter adp = new DataSet1TableAdapters.休日TableAdapter();

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 閉じる
            Close();
        }

        private void prePrint_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            Dispose();
        }

        private void prePrint_Load(object sender, EventArgs e)
        {
            // フォームの最大・最小サイズの指定
            Utility.WindowsMaxSize(this, this.Width, this.Height);
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // 部署一覧グリッド定義
            gridViewSet(dataGridView1);

            // 部門一覧表示
            departmentShow();

            //// 雇用区分：社員
            //cmbKoyou.SelectedIndex = 0;

            // 帳票名
            cmbPrnName.SelectedIndex = 0;

            label3.Visible = false;

            toolStripProgressBar1.Visible = false;

            // 勤務体系テーブル選択 2017/08/30
            label6.Text = string.Empty;
            button1.Enabled = true;
        }

        // カラム定義
        private string ColChk = "c0";
        private string ColSz = "c1";
        private string ColSznm = "c2";
        private string ColCode = "c3";
        private string ColNin = "c4";
        private string ColID = "c5";

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
                dg.Height = 288;

                // 奇数行の色
                dg.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;

                // 各列指定
                DataGridViewCheckBoxColumn chk = new DataGridViewCheckBoxColumn();
                chk.Name = ColChk;
                dg.Columns.Add(chk);
                dg.Columns[ColChk].HeaderText = "";

                dg.Columns.Add(ColCode, "コード");
                dg.Columns.Add(ColSznm, "部署名");
                dg.Columns.Add(ColNin, "人数");
                dg.Columns.Add(ColID, "ID");

                dg.Columns[ColID].Visible = false;

                dg.Columns[ColChk].Width = 30;
                dg.Columns[ColCode].Width = 80;
                dg.Columns[ColSznm].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dg.Columns[ColNin].Width = 80;
                
                dg.Columns[ColChk].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dg.Columns[ColCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dg.Columns[ColNin].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

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
            gridViewShowData(sc, dataGridView1, dateTimePicker1.Value);
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     グリッドビューへ部門情報を表示する </summary>
        /// <param name="sConnect">
        ///     データベース接続文字列</param>
        /// <param name="dg">
        ///     DataGridViewオブジェクト名</param>
        /// <param name="sDt">
        ///     基準年月日</param>
        ///---------------------------------------------------------------------
        private void gridViewShowData(string sConnect, DataGridView dg, DateTime sDt)
        {
            sqlControl.DataControl sdCon = new Common.sqlControl.DataControl(sConnect);
            sqlControl.DataControl sdCon2 = new Common.sqlControl.DataControl(sConnect);
            SqlDataReader dR;

            //string dt = DateTime.Today.ToShortDateString();

            // 2017/09/28
            string dt = sDt.ToShortDateString();

            /* 以下の条件で絞込 
             * 設立年月日 <= , 廃止年月日 >=
             * 有効期間（開始）<=, 有効期間（終了）>= 
             */

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
                    int nin = getBumonEmpCount(sdCon2, dR["DepartmentCode"].ToString(), sDt);
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

        private void linkLblOn_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("全ての部署を印刷対象とします。よろしいですか。","",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == DialogResult.No)
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

        private void linkPrn_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            int pCnt = 0;

            if (cmbPrnName.SelectedIndex == -1)
            {
                MessageBox.Show("印刷する帳票を選択してください", "帳票名", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // 2017/08/30
            if (cmbPrnName.SelectedIndex == 0 && label6.Text == string.Empty)
            {
                MessageBox.Show("Excel部署別勤務体系ファイルを選択してください", "参照ファイル未設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            
            if (DateTime.Parse(dateTimePicker1.Value.ToShortDateString()).CompareTo(DateTime.Parse(dateTimePicker2.Value.ToShortDateString())) == 1)
            {
                MessageBox.Show("印刷日付範囲が正しくありません", "印刷日範囲", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // 選択部署
            foreach (DataGridViewRow r in dataGridView1.Rows)
            {
                // チェックされている部署を対象とする
                if (dataGridView1[ColChk, r.Index].Value.ToString() == "True")
                {
                    // 勤怠データＩ／Ｐ票のとき
                    if (cmbPrnName.SelectedIndex == 0)
                    {
                        if (!chkWhite.Checked)
                        {
                            // 社員情報印字のとき
                            pCnt += getTotalPages(Utility.StrtoInt(dataGridView1[ColNin, r.Index].Value.ToString()));
                        }
                        else
                        {
                            // 白紙印刷のとき
                            pCnt++;
                        }
                    }
                    else if (cmbPrnName.SelectedIndex == 1) // 応援移動票のとき
                    {
                        pCnt++;
                    }
                }
            }
            
            if (pCnt == 0)
            {
                MessageBox.Show("印刷する部署を選択してください", "印刷部署", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // 勤怠データＩ／Ｐ票のとき
            if (cmbPrnName.SelectedIndex == 0)
            {
                // 白紙印刷モードのとき
                if (chkWhite.Checked)
                {
                    if (MessageBox.Show("白紙発行モードです。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                    {
                        return;
                    }
                }
            }

            // 日数分
            TimeSpan span = DateTime.Parse(dateTimePicker2.Value.ToShortDateString()) - DateTime.Parse(dateTimePicker1.Value.ToShortDateString());
            pCnt = pCnt * (span.Days + 1) + 1;

            // 確認メッセージ
            if (MessageBox.Show(cmbPrnName.Text + "を発行します。よろしいですか？","確認",MessageBoxButtons.YesNo,MessageBoxIcon.Information) == DialogResult.No)
            {
                return;
            }

            xlsData xls = new xlsData();

            // 勤務体系シートデータ配列取得
            object[,] sArray = null;
            if (cmbPrnName.SelectedIndex == 0)
            {
                sArray = xls.getShiftCode(label6.Text);
            }

            // 部署別残業理由シートデータ配列取得
            object[,] zArray = xls.getZanReason();

            // 印刷ダイアログ表示
            PrintDialog pd = new PrintDialog();
            pd.PrinterSettings = new System.Drawing.Printing.PrinterSettings();

            if (pd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string printerName = pd.PrinterSettings.PrinterName; // プリンター名
                int copies = pd.PrinterSettings.Copies; // 印刷部数
                bool ptof = pd.PrinterSettings.PrintToFile; // printToFile

                // 印刷処理
                if (cmbPrnName.SelectedIndex == 0)
                {
                    // 勤怠データＩ／Ｐ票発行
                    // 印刷プリンター、部数を指定可能とした 2017/08/08
                    prnSheet(sArray, zArray, pCnt, chkWhite.Checked, printerName, copies);
                }
                else if (cmbPrnName.SelectedIndex == 1)
                {
                    // 応援移動票発行
                    // 印刷プリンター、部数を指定可能とした 2017/08/08
                    //prnOuenSheet(zArray, pCnt);
                    prnOuenSheetNonDate(zArray, pCnt, printerName, copies);
                }
            }
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     勤怠データＩ／Ｐ票印刷処理 </summary>
        /// <param name="sArray">
        ///     部署別勤務体系配列 </param>
        /// <param name="zArray">
        ///     部署別残業理由配列 </param>
        /// <param name="pC">
        ///     印刷シート数</param>
        /// <param name="wStatus">
        ///     true:白紙印刷, false:通常印刷</param>
        /// <param name="prnName">
        ///     印刷するプリンタ名</param>
        /// <param name="copies">
        ///     部数</param>
        ///----------------------------------------------------------------------
        private void prnSheet(object [,] sArray, object [,] zArray, int pC, bool wStatus, string prnName, int copies)
        {
            DateTime[] holiday = new DateTime[1];

            bool hol = false;   // 休日ステータス

            // 休日配列初期化
            holiday[0] = DateTime.Parse("1900/01/01");

            DateTime sDt = DateTime.Parse(dateTimePicker1.Value.ToShortDateString());
            DateTime eDt = DateTime.Parse(dateTimePicker2.Value.ToShortDateString());
            DateTime nDt;

            int iH = 1;

            // 休日データを配列に読み込む
            foreach (var t in dts.休日.Where(a => a.年月日 >= sDt && a.年月日 <= eDt))
            {
                Array.Resize(ref holiday, iH + 1);
                holiday[iH] = t.年月日;
                iH++;
            }

            // ライン・部門・製品群コード配列取得 
            string[] hArray = getCategoryArray();

            // シーケンス番号：2017/08/08
            int seqNum = 0;

            //マウスポインタを待機にする
            this.Cursor = Cursors.WaitCursor;

            // Excel起動
            string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

            Excel.Application oXls = new Excel.Application();

            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(Properties.Settings.Default.xlsIPSheet, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];
            Excel.Worksheet oxlsMsSheet = (Excel.Worksheet)oXlsBook.Sheets[1]; // テンプレートシート
            oxlsSheet.Select(Type.Missing);

            Excel.Range rng = null;
        
            int pCnt = 1;   // ページカウント
            //int bCount = 0; // progressBar部署カウント
            object[,] rtnArray = null;

            try
            {
                // progressBar
                toolStripProgressBar1.Maximum = 100;
                toolStripProgressBar1.Minimum = 0;
                toolStripProgressBar1.Visible = true;

                // 部署ループ
                foreach (DataGridViewRow r in dataGridView1.Rows)
                {
                    // チェックされている部署を対象とする
                    if (dataGridView1[ColChk, r.Index].Value.ToString() == "False")
                    {
                        continue;
                    }

                    int iX = 1;

                    // 日付初期化
                    nDt = DateTime.Parse(dateTimePicker1.Value.ToShortDateString());

                    // 部署コード取得
                    string bCode = dataGridView1[ColCode, r.Index].Value.ToString().Trim();

                    // 日付指定範囲でループさせる
                    while (true)
                    {
                        // 休日の判定
                        hol = isHoliday(nDt, holiday);

                        // テンプレートシートを追加する
                        pCnt++;                        
                        oxlsMsSheet.Copy(Type.Missing, oXlsBook.Sheets[pCnt - 1]);
                        oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt];

                        // シートのセルを一括して配列に取得します
                        rng = oxlsMsSheet.Range[oxlsMsSheet.Cells[1, 1], oxlsMsSheet.Cells[oxlsMsSheet.UsedRange.Rows.Count, 100]];
                        rtnArray = (object[,])rng.Value2;

                        // ページ
                        int sp = 1;
                        rtnArray[2, 88] = sp.ToString();

                        // 総ページ数
                        if (wStatus)
                        {
                            // 白紙印刷のとき
                            rtnArray[2, 97] = "1";
                        }
                        else
                        {
                            // 社員情報印字のとき
                            rtnArray[2, 97] = getTotalPages(Utility.StrtoInt(dataGridView1[ColNin, r.Index].Value.ToString()));
                        }

                        // 部署別残業理由をシート配列にセットする
                        setZangyoReasonList(ref rtnArray, zArray, bCode);

                        // 事由コード一覧をシート配列にセット
                        setJiyuCodeList(ref rtnArray);
                        
                        // 部署別シフト（勤務体系）コード一覧をシート配列にセットする
                        // 2017/12/14 自動印刷処理撤廃
                        //setShiftCodeList(ref rtnArray, sArray, bCode);

                        // 年
                        rtnArray[2, 2] = nDt.Year.ToString().Substring(0, 1);
                        rtnArray[2, 5] = nDt.Year.ToString().Substring(1, 1);
                        rtnArray[2, 8] = nDt.Year.ToString().Substring(2, 1);
                        rtnArray[2, 11] = nDt.Year.ToString().Substring(3, 1);

                        // 月
                        rtnArray[2, 17] = nDt.Month.ToString().PadLeft(2, '0').Substring(0, 1);
                        rtnArray[2, 20] = nDt.Month.ToString().PadLeft(2, '0').Substring(1, 1);

                        // 日
                        rtnArray[2, 26] = nDt.Day.ToString().PadLeft(2, '0').Substring(0, 1);
                        rtnArray[2, 29] = nDt.Day.ToString().PadLeft(2, '0').Substring(1, 1);

                        // 部署コード
                        rtnArray[4, 7] = bCode.Substring(0, 1);
                        rtnArray[4, 10] = bCode.Substring(1, 1);
                        rtnArray[4, 13] = bCode.Substring(2, 1);
                        rtnArray[4, 16] = bCode.Substring(3, 1);
                        rtnArray[4, 19] = bCode.Substring(4, 1);

                        // 部署名
                        rtnArray[7, 2] = dataGridView1[ColSznm, r.Index].Value.ToString().Trim();

                        // 勤務体系（シフト）コードをシート配列にセットする
                        setShiftCode(ref rtnArray, sArray, bCode, hol);

                        // シーケンス番号：2017/08/08
                        seqNum++;
                        rtnArray[80, 2] = seqNum.ToString("D3") + "/" + (pC - 1).ToString("D3");

                        // 白紙印刷ではないとき社員情報印字
                        if (!wStatus)
                        {
                            // 接続文字列取得
                            string sc = sqlControl.obcConnectSting.get(_dbName);
                            sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

                            //// 雇用区分取得
                            //int koyou = getKoyouKbn();

                            // DepartmentCode取得
                            string strCode = Utility.getDepartmentCode(bCode);

                            // 部署・社員情報データリーダー取得 : 基準年月日（開始年月日）2017/09/28
                            SqlDataReader dr = sdCon.free_dsReader(Utility.getEmployeeOrder(strCode, sDt));

                            int rC = 19;

                            // 社員情報を配列にセット
                            while (dr.Read())
                            {
                                // ページいっぱいで次ページの準備
                                if (rC > 67 && pCnt < pC)
                                {
                                    // 配列から現在のシートセルに一括してデータをセットします
                                    rng = oxlsSheet.Range[oxlsSheet.Cells[1, 1], oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 100]];
                                    rng.Value2 = rtnArray;

                                    // 現在のページをコピーする
                                    pCnt++;
                                    oxlsSheet.Copy(Type.Missing, oXlsBook.Sheets[pCnt - 1]);
                                    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt];

                                    // シートのセル情報を一括して配列に取得します
                                    rng = oxlsSheet.Range[oxlsSheet.Cells[1, 1], oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 100]];
                                    rtnArray = (object[,])rng.Value2;

                                    // 前シートの社員情報をクリアする
                                    for (int i = 19; i <= 67; i += 6)
                                    {
                                        rtnArray[i, 12] = string.Empty;   // 氏名

                                        // 社員番号
                                        rtnArray[i, 32] = string.Empty;
                                        rtnArray[i, 35] = string.Empty;
                                        rtnArray[i, 38] = string.Empty;
                                        rtnArray[i, 41] = string.Empty;
                                        rtnArray[i, 44] = string.Empty;
                                        rtnArray[i, 47] = string.Empty;

                                        rtnArray[i, 51] = string.Empty; // ライン
                                        rtnArray[i, 57] = string.Empty; // 部門
                                        rtnArray[i, 63] = string.Empty; // 製品群
                                    }

                                    // ページ数加算
                                    sp++;
                                    rtnArray[2, 88] = sp.ToString();

                                    // シーケンス番号：2017/08/08
                                    seqNum++;
                                    rtnArray[80, 2] = seqNum.ToString("D3") + "/" + (pC - 1).ToString("D3");

                                    rC = 19;
                                }


                                rtnArray[rC, 12] = dr["Name"].ToString();   // 氏名

                                // 社員番号
                                string dCode = string.Empty;
                                int len = dr["EmployeeNo"].ToString().Trim().Length;

                                if (len > 6)
                                {
                                    dCode = dr["EmployeeNo"].ToString().Substring(len - 6, 6);
                                }
                                else
                                {
                                    dCode = dr["EmployeeNo"].ToString().Trim().PadLeft(6, '0');
                                }

                                rtnArray[rC, 32] = dCode.Substring(0, 1);
                                rtnArray[rC, 35] = dCode.Substring(1, 1);
                                rtnArray[rC, 38] = dCode.Substring(2, 1);
                                rtnArray[rC, 41] = dCode.Substring(3, 1);
                                rtnArray[rC, 44] = dCode.Substring(4, 1);
                                rtnArray[rC, 47] = dCode.Substring(5, 1);

                                rtnArray[rC, 51] = getHisCategory(hArray, dr["JobTypeID"].ToString());      // ライン
                                rtnArray[rC, 57] = getHisCategory(hArray, dr["DutyID"].ToString());         // 部門
                                rtnArray[rC, 63] = getHisCategory(hArray, dr["QualificationGradeID"].ToString().Trim());  // 製品群

                                rC += 6;

                                //// ページいっぱいで次ページの準備
                                //if (rC > 67 && pCnt < pC)
                                //{
                                //    // 配列から現在のシートセルに一括してデータをセットします
                                //    rng = oxlsSheet.Range[oxlsSheet.Cells[1, 1], oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 100]];
                                //    rng.Value2 = rtnArray;

                                //    // 現在のページをコピーする
                                //    pCnt++;
                                //    oxlsSheet.Copy(Type.Missing, oXlsBook.Sheets[pCnt - 1]);
                                //    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt];

                                //    // シートのセル情報を一括して配列に取得します
                                //    rng = oxlsSheet.Range[oxlsSheet.Cells[1, 1], oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 100]];
                                //    rtnArray = (object[,])rng.Value2;

                                //    // 前シートの社員情報をクリアする
                                //    for (int i = 19; i <= 67; i += 6)
                                //    {
                                //        rtnArray[i, 12] = string.Empty;   // 氏名

                                //        // 社員番号
                                //        rtnArray[i, 32] = string.Empty;
                                //        rtnArray[i, 35] = string.Empty;
                                //        rtnArray[i, 38] = string.Empty;
                                //        rtnArray[i, 41] = string.Empty;
                                //        rtnArray[i, 44] = string.Empty;
                                //        rtnArray[i, 47] = string.Empty;

                                //        rtnArray[i, 51] = string.Empty; // ライン
                                //        rtnArray[i, 57] = string.Empty; // 部門
                                //        rtnArray[i, 63] = string.Empty; // 製品群
                                //    }

                                //    // ページ数加算
                                //    sp++;
                                //    rtnArray[2, 88] = sp.ToString();

                                //    rC = 19;
                                //}
                            }

                            dr.Close();
                            sdCon.Close();
                        }

                        // 配列からシートセルに一括してデータをセットします
                        rng = oxlsSheet.Range[oxlsSheet.Cells[1, 1], oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 100]];
                        rng.Value2 = rtnArray;

                        // progressBar表示
                        label3.Visible = true;
                        label3.Text = "印刷データ作成中..." + pCnt.ToString() + "/" + pC.ToString();

                        toolStripProgressBar1.Value = pCnt * 100 / pC;

                        // 日付範囲を超えたらループからぬける
                        nDt = sDt.AddDays(iX);
                        bool nxt = true;
                        switch (nDt.CompareTo(eDt))
                        {
                            case -1:    // 期限内
                                nxt = true;
                                break;
                            case 0:     // 期限日と同日
                                nxt = true;
                                break;
                            case 1:     // 期限日超過
                                nxt = false;
                                break;
                        }

                        if(!nxt)
                        {
                            break;
                        }

                        iX++;
                    }
                }

                // 1枚目はテンプレートシートなので印刷時には削除する
                oXls.DisplayAlerts = false;
                oXlsBook.Sheets[1].Delete();

                //System.Threading.Thread.Sleep(1000);

                // 確認のためExcelのウィンドウを表示する
                oXls.Visible = true;

                // 印刷：プリンタ名、部数を指定可能とした 2017/08/08
                //oXlsBook.PrintOutEx(Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                oXlsBook.PrintOutEx(Type.Missing, Type.Missing, copies, true, prnName, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                
                //oXlsBook.PrintPreview(false);
                
                oXls.Visible = false;

                // 終了メッセージ 
                MessageBox.Show("終了しました");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            finally
            {
                // ウィンドウを非表示にする
                oXls.Visible = false;

                // 保存処理
                oXls.DisplayAlerts = false;

                // Bookをクローズ
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                // Excelを終了
                oXls.Quit();

                // COM オブジェクトの参照カウントを解放する 
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsMsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                oXls = null;
                oXlsBook = null;
                oxlsSheet = null;
                oxlsMsSheet = null;

                GC.Collect();

                //マウスポインタを元に戻す
                this.Cursor = Cursors.Default;

                // progreassBar非表示
                label3.Visible = false;
                toolStripProgressBar1.Visible = false;
                toolStripProgressBar1.Value = 0;
            }
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     勤務体系（シフト）コードをシート配列にセットする </summary>
        /// <param name="rtnArray">
        ///     シート配列オブジェクト</param>
        /// <param name="sArray">
        ///     勤務体系（シフト）配列</param>
        /// <param name="bCode">
        ///     部署コード</param>
        /// <param name="hol">
        ///     休日ステータス　true:休日、false:勤務日</param>
        ///------------------------------------------------------------------
        private void setShiftCode(ref object [,] rtnArray, object [,] sArray, string bCode, bool hol)
        {
            for (int i = 1; i <= sArray.GetLength(0); i++)
            {
                if (sArray[i, 1].ToString() == bCode)
                {
                    // 休日のとき
                    if (hol && sArray[i, 4].ToString() != global.FLGON)
                    {
                        continue;
                    }

                    // 勤務日のとき
                    if (!hol && sArray[i, 4].ToString() != global.FLGOFF)
                    {
                        continue;
                    }

                    // シフト（勤務体系）コード
                    string sftCode = sArray[i, 2].ToString().PadLeft(3, '0');
                    rtnArray[4, 24] = sftCode.Substring(0, 1);
                    rtnArray[4, 27] = sftCode.Substring(1, 1);
                    rtnArray[4, 30] = sftCode.Substring(2, 1);

                    // シフト（勤務体系）名
                    rtnArray[7, 44] = sArray[i, 3].ToString();

                    // ループから抜ける
                    break;
                }
            }
        }


        ///---------------------------------------------------------------
        /// <summary>
        ///     任意の日付を休日か判定する </summary>
        /// <param name="nDt">
        ///     日付</param>
        /// <param name="holiday">
        ///     休日配列</param>
        /// <returns>
        ///     true:休日、false:勤務日</returns>
        ///---------------------------------------------------------------
        private bool isHoliday(DateTime nDt, DateTime[] holiday)
        {
            bool rtn = false;

            // 休日の判定
            foreach (var item in holiday)
            {
                // 休日のとき
                if (item == nDt)
                {
                    rtn = true;
                    break;
                }
            }

            return rtn;
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     部署別残業理由一覧をシート配列にセットする </summary>
        /// <param name="rtnArray">
        ///     シート配列オブジェクト</param>
        /// <param name="zArray">
        ///     部署別残業理由配列</param>
        /// <param name="bCode">
        ///     部署コード</param>
        ///----------------------------------------------------------------------
        private void setZangyoReasonList(ref object [,] rtnArray, object [,] zArray, string bCode)
        {
            int iZ = 0;

            // 部署別残業理由コードセット
            for (int i = 1; i <= zArray.GetLength(0); i++)
            {
                if (zArray[i, 1].ToString() != bCode)
                {
                    continue;
                }

                iZ++;

                string nm = Utility.NulltoStr(zArray[i, 3]);
                string cd = Utility.NulltoStr(zArray[i, 2]).PadLeft(2, '0');

                if (iZ == 1)
                {
                    rtnArray[74, 2] = cd;
                    rtnArray[74, 5] = nm;
                }
                else if (iZ == 2)
                {
                    rtnArray[75, 2] = cd;
                    rtnArray[75, 5] = nm;
                }
                else if (iZ == 3)
                {
                    rtnArray[76, 2] = cd;
                    rtnArray[76, 5] = nm;
                }
                else if (iZ == 4)
                {
                    rtnArray[77, 2] = cd;
                    rtnArray[77, 5] = nm;
                }
                else if (iZ == 5)
                {
                    rtnArray[78, 2] = cd;
                    rtnArray[78, 5] = nm;
                }
                else if (iZ == 6)
                {
                    rtnArray[74, 23] = cd;
                    rtnArray[74, 26] = nm;
                }
                else if (iZ == 7)
                {
                    rtnArray[75, 23] = cd;
                    rtnArray[75, 26] = nm;
                }
                else if (iZ == 8)
                {
                    rtnArray[76, 23] = cd;
                    rtnArray[76, 26] = nm;
                }
                else if (iZ == 9)
                {
                    rtnArray[77, 23] = cd;
                    rtnArray[77, 26] = nm;
                }
                else if (iZ == 10)
                {
                    rtnArray[78, 23] = cd;
                    rtnArray[78, 26] = nm;
                }
            }
        }


        private void setZangyoReasonOuen(ref object[,] rtnArray, object[,] zArray, string bCode)
        {
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     部署別勤務体系（シフト）一覧をシート配列にセットする </summary>
        /// <param name="rtnArray">
        ///     シート配列オブジェクト</param>
        /// <param name="sArray">
        ///     部署別勤務体系（シフト）配列</param>
        /// <param name="bCode">
        ///     部署コード</param>
        ///----------------------------------------------------------------------
        private void setShiftCodeList(ref object [,] rtnArray, object [,] sArray, string bCode)
        {
            int iS = 73;

            // 部署別勤務体系配列からシフト（勤務体系）コード一覧を取得する
            for (int i = 1; i <= sArray.GetLength(0); i++)
            {
                if (sArray[i, 1].ToString() != bCode)
                {
                    continue;
                }

                rtnArray[iS, 81] = sArray[i, 2].ToString();
                rtnArray[iS, 85] = sArray[i, 3].ToString();
                iS++;

                if (iS > 80)
                {
                    break;
                }
            }
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     事由コード一覧をシート配列にセット </summary>
        /// <param name="rtnArray">
        ///     シート配列</param>
        ///-------------------------------------------------------------
        private void setJiyuCodeList(ref object [,] rtnArray)
        {
            // 接続文字列取得
            string sc = sqlControl.obcConnectSting.get(_dbName);
            sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT LaborReasonCode,LaborReasonName from tbLaborReason ");
            sb.Append("where IsValid = 1 and LaborReasonCode <> '' ");
            sb.Append("order by LaborReasonCode");

            // 部署・社員情報データリーダー取得
            SqlDataReader dr = sdCon.free_dsReader(sb.ToString());

            int i = 0;
            int r = 0;
            int c = 0;

            int [,] cr = { { 45, 48 }, { 53, 56 }, { 61, 64 }, { 69, 72 } };

            while (dr.Read())
            {
                i++;

                if (i <= 4)
                {
                    r = 74;
                    c = i - 1;
                }
                else if (i <= 8)
                {
                    r = 75;
                    c = i - 5;
                }
                else if (i <= 12)
                {
                    r = 76;
                    c = i - 9;
                }
                else if (i <= 16)
                {
                    r = 77;
                    c = i - 13;
                }
                else if (i <= 20)
                {
                    r = 78;
                    c = i - 17;
                }
                else if (i <= 24)
                {
                    r = 79;
                    c = i - 21;
                }
                else
                {
                    break;
                }

                rtnArray[r, cr[c, 0]] = dr["LaborReasonCode"].ToString();
                rtnArray[r, cr[c, 1]] = dr["LaborReasonName"].ToString();
            }

            dr.Close();
            sdCon.Close();
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     社員情報抽出ＳＱＬ作成 </summary>
        /// <param name="bCode">
        ///     部署コード</param>
        /// <returns>
        ///     ＳＱＬ文字列</returns>
        ///---------------------------------------------------------------
        //private string getEmployeeCount(string bCode)
        //{
        //    string dt = DateTime.Today.ToShortDateString();

        //    // 社員情報抽出ＳＱＬ
        //    StringBuilder sb = new StringBuilder();
        //    sb.Append("SELECT count(tbEmployeeBase.EmployeeID) as cnt ");

        //    sb.Append("from(((tbEmployeeBase inner join ");
        //    sb.Append("(select tbEmployeeMainDutyPersonnelChange.EmployeeID, tbEmployeeMainDutyPersonnelChange.AnnounceDate,");
        //    sb.Append("tbEmployeeMainDutyPersonnelChange.BelongID, tbEmployeeMainDutyPersonnelChange.DutyID,");
        //    sb.Append("tbEmployeeMainDutyPersonnelChange.JobTypeID, tbEmployeeMainDutyPersonnelChange.QualificationGradeID ");

        //    sb.Append("from tbEmployeeMainDutyPersonnelChange inner join ");

        //    sb.Append("(select EmployeeID, max(AnnounceDate) as AnnounceDate from tbEmployeeMainDutyPersonnelChange ");
        //    sb.Append("where AnnounceDate <= '").Append(DateTime.Today.ToShortDateString()).Append("' ");
        //    sb.Append("group by EmployeeID) as a ");
        //    sb.Append("on(tbEmployeeMainDutyPersonnelChange.EmployeeID = a.EmployeeID) and ");
        //    sb.Append("(tbEmployeeMainDutyPersonnelChange.AnnounceDate = a.AnnounceDate) ");
        //    sb.Append(") as d ");
        //    sb.Append("on tbEmployeeBase.EmployeeID = d.EmployeeID) ");

        //    sb.Append("inner join tbDepartment on d.BelongID = tbDepartment.DepartmentID) ");
        //    sb.Append("inner join tbHR_DivisionCategory on tbEmployeeBase.BeOnTheRegisterDivisionID = tbHR_DivisionCategory.CategoryID) ");
        //    sb.Append("where DepartmentCode = '" + bCode + "' and tbHR_DivisionCategory.CategoryCode <> 2"); // 2017/05/08 

        //    return sb.ToString();
        //}

        ///-------------------------------------------------------------------
        /// <summary>
        ///     ライン・部門・製品群コード配列取得   </summary>
        /// <returns>
        ///     ID,コード配列</returns>
        ///-------------------------------------------------------------------
        private string [] getCategoryArray()
        {
            // 接続文字列取得
            string sc = sqlControl.obcConnectSting.get(_dbName);
            sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            StringBuilder sb = new StringBuilder();
            sb.Append("select CategoryID, CategoryCode from tbHistoryDivisionCategory");
            SqlDataReader dr = sdCon.free_dsReader(sb.ToString());

            int iX = 0;
            string[] hArray = new string[1];

            while(dr.Read())
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

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     ライン・部門・製品群コード取得　</summary>
        /// <param name="hArray">
        ///     配列</param>
        /// <param name="sCode">
        ///     CategoryID</param>
        /// <returns>
        ///     CategoryCode</returns>
        ///-----------------------------------------------------------------------
        private string getHisCategory(string [] hArray, string sCode)
        {
            string rtnCode = "";

            foreach (var t in hArray)
            {
                string [] n = t.Split(',');

                if (n[0].ToString() == sCode)
                {
                    rtnCode = n[1];
                    break;
                }
            }

            return rtnCode.Trim();
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
            SqlDataReader cDr = sdCon.free_dsReader(Utility.getEmployeeCount(strCode, sDt));
            while (cDr.Read())
            {
                nin = Utility.StrtoInt(cDr["cnt"].ToString());
                break;
            }

            cDr.Close();

            return nin;
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     部署ごとの総ページ数を取得する </summary>
        /// <param name="n">
        ///     社員数 </param>
        /// <returns>
        ///     総ページ数 </returns>
        ///-------------------------------------------------------------
        private int getTotalPages(int n)
        {
            int tp = 0;

            if ((n % 9 != 0))
            {
                tp = (n / 9) + 1;
            }
            else
            {
                tp = n / 9;
            }

            return tp;
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker2.MinDate = dateTimePicker1.Value;

            // 部門一覧表示
            departmentShow();
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            if (DateTime.Parse(dateTimePicker1.Value.ToShortDateString()).CompareTo(DateTime.Parse(dateTimePicker2.Value.ToShortDateString())) == 1)
            {
                dateTimePicker1.Value = dateTimePicker2.Value;
            }
        }

        private void chkWhite_CheckedChanged(object sender, EventArgs e)
        {
            if (chkWhite.Checked)
            {
                chkWhite.ForeColor = Color.Red;
            }
            else
            {
                chkWhite.ForeColor = SystemColors.ControlText;
            }
        }

        private void cmbPrnName_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbPrnName.SelectedIndex == 0)
            {
                // 勤怠データI/P票印刷
                chkWhite.Visible = true;
                dateTimePicker1.Enabled = true;
                dateTimePicker2.Enabled = true;

                button1.Enabled = true; // 2017/08/30
            }
            else
            {
                // 応援移動票
                chkWhite.Visible = false;
                dateTimePicker1.Value = DateTime.Today;
                dateTimePicker2.Value = DateTime.Today;
                dateTimePicker1.Enabled = false;
                dateTimePicker2.Enabled = false;

                button1.Enabled = false; // 2017/08/30
                label6.Text = string.Empty;
            }
        }
        
        ///----------------------------------------------------------------------
        /// <summary>
        ///     応援移動票印刷処理 </summary>
        /// <param name="zArray">
        ///     部署別残業理由配列 </param>
        /// <param name="pC">
        ///     印刷シート数</param>
        ///----------------------------------------------------------------------
        private void prnOuenSheet(object[,] zArray, int pC)
        {
            DateTime[] holiday = new DateTime[1];

            bool hol = false;   // 休日ステータス

            // 休日配列初期化
            holiday[0] = DateTime.Parse("1900/01/01");

            DateTime sDt = DateTime.Parse(dateTimePicker1.Value.ToShortDateString());
            DateTime eDt = DateTime.Parse(dateTimePicker2.Value.ToShortDateString());
            DateTime nDt;

            int iH = 1;

            // 休日データを配列に読み込む
            foreach (var t in dts.休日.Where(a => a.年月日 >= sDt && a.年月日 <= eDt))
            {
                Array.Resize(ref holiday, iH + 1);
                holiday[iH] = t.年月日;
                iH++;
            }
            
            //マウスポインタを待機にする
            this.Cursor = Cursors.WaitCursor;

            // Excel起動
            string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

            Excel.Application oXls = new Excel.Application();

            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(Properties.Settings.Default.xlsIdouSheet, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];
            Excel.Worksheet oxlsMsSheet = (Excel.Worksheet)oXlsBook.Sheets[1]; // テンプレートシート
            oxlsSheet.Select(Type.Missing);

            Excel.Range rng = null;

            int pCnt = 1;   // ページカウント
            object[,] rtnArray = null;

            try
            {
                // progressBar
                toolStripProgressBar1.Maximum = 100;
                toolStripProgressBar1.Minimum = 0;
                toolStripProgressBar1.Visible = true;

                // 部署ループ
                foreach (DataGridViewRow r in dataGridView1.Rows)
                {
                    // チェックされている部署を対象とする
                    if (dataGridView1[ColChk, r.Index].Value.ToString() == "False")
                    {
                        continue;
                    }

                    int iX = 1;

                    // 日付初期化
                    nDt = DateTime.Parse(dateTimePicker1.Value.ToShortDateString());

                    // 部署コード取得
                    string bCode = dataGridView1[ColCode, r.Index].Value.ToString().Trim();

                    // 日付指定範囲でループさせる
                    while (true)
                    {
                        // 休日の判定
                        hol = isHoliday(nDt, holiday);

                        // テンプレートシートを追加する
                        pCnt++;
                        oxlsMsSheet.Copy(Type.Missing, oXlsBook.Sheets[pCnt - 1]);
                        oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt];

                        // シートのセルを一括して配列に取得します
                        rng = oxlsMsSheet.Range[oxlsMsSheet.Cells[1, 1], oxlsMsSheet.Cells[oxlsMsSheet.UsedRange.Rows.Count, 143]];
                        rtnArray = (object[,])rng.Value2;
                        
                        int sp = 1;
                        rtnArray[4, 27] = sp.ToString();    // ページ
                        rtnArray[4, 36] = "1";  // 総ページ数

                        // 部署別残業理由をシート配列にセットする
                        setZangyoReasonList(ref rtnArray, zArray, bCode);

                        // 年
                        rtnArray[2, 2] = nDt.Year.ToString().Substring(0, 1);
                        rtnArray[2, 5] = nDt.Year.ToString().Substring(1, 1);
                        rtnArray[2, 8] = nDt.Year.ToString().Substring(2, 1);
                        rtnArray[2, 11] = nDt.Year.ToString().Substring(3, 1);

                        // 月
                        rtnArray[2, 17] = nDt.Month.ToString().PadLeft(2, '0').Substring(0, 1);
                        rtnArray[2, 20] = nDt.Month.ToString().PadLeft(2, '0').Substring(1, 1);

                        // 日
                        rtnArray[2, 26] = nDt.Day.ToString().PadLeft(2, '0').Substring(0, 1);
                        rtnArray[2, 29] = nDt.Day.ToString().PadLeft(2, '0').Substring(1, 1);

                        // 部署コード
                        rtnArray[4, 7] = bCode.Substring(0, 1);
                        rtnArray[4, 10] = bCode.Substring(1, 1);
                        rtnArray[4, 13] = bCode.Substring(2, 1);
                        rtnArray[4, 16] = bCode.Substring(3, 1);
                        rtnArray[4, 19] = bCode.Substring(4, 1);

                        // 部署名
                        rtnArray[4, 41] = dataGridView1[ColSznm, r.Index].Value.ToString().Trim();
                        
                        // 配列からシートセルに一括してデータをセットします
                        rng = oxlsSheet.Range[oxlsSheet.Cells[1, 1], oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 100]];
                        rng.Value2 = rtnArray;


                        int iZ = 1;

                        // 部署別残業理由コードセット
                        Excel.TextBox etxt = null;
                        for (int i = 1; i <= zArray.GetLength(0); i++)
                        {
                            if (zArray[i, 1].ToString() != bCode)
                            {
                                continue;
                            }

                            if (iZ == 1)
                            {
                                etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zan1");
                                etxt.Text = zArray[i, 3].ToString();
                                etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zanC1");
                                etxt.Text = zArray[i, 2].ToString().PadLeft(2, '0');
                            }
                            else if (iZ == 2)
                            {
                                etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zan2");
                                etxt.Text = zArray[i, 3].ToString();
                                etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zanC2");
                                etxt.Text = zArray[i, 2].ToString().PadLeft(2, '0');
                            }
                            else if (iZ == 3)
                            {
                                etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zan3");
                                etxt.Text = zArray[i, 3].ToString();
                                etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zanC3");
                                etxt.Text = zArray[i, 2].ToString().PadLeft(2, '0');
                            }
                            else if (iZ == 4)
                            {
                                etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zan4");
                                etxt.Text = zArray[i, 3].ToString();
                                etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zanC4");
                                etxt.Text = zArray[i, 2].ToString().PadLeft(2, '0');
                            }
                            else if (iZ == 5)
                            {
                                etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zan5");
                                etxt.Text = zArray[i, 3].ToString();
                                etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zanC5");
                                etxt.Text = zArray[i, 2].ToString().PadLeft(2, '0');
                            }
                            else if (iZ == 6)
                            {
                                etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zan6");
                                etxt.Text = zArray[i, 3].ToString();
                                etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zanC6");
                                etxt.Text = zArray[i, 2].ToString().PadLeft(2, '0');
                            }
                            else if (iZ == 7)
                            {
                                etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zan7");
                                etxt.Text = zArray[i, 3].ToString();
                                etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zanC7");
                                etxt.Text = zArray[i, 2].ToString().PadLeft(2, '0');
                            }
                            else if (iZ == 8)
                            {
                                etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zan8");
                                etxt.Text = zArray[i, 3].ToString();
                                etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zanC8");
                                etxt.Text = zArray[i, 2].ToString().PadLeft(2, '0');
                            }
                            else if (iZ == 9)
                            {
                                etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zan9");
                                etxt.Text = zArray[i, 3].ToString();
                                etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zanC9");
                                etxt.Text = zArray[i, 2].ToString().PadLeft(2, '0');
                            }
                            else if (iZ == 10)
                            {
                                etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zan10");
                                etxt.Text = zArray[i, 3].ToString();
                                etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zanC10");
                                etxt.Text = zArray[i, 2].ToString().PadLeft(2, '0');
                            }

                            iZ++;
                        }

                        // progressBar表示
                        label3.Visible = true;
                        label3.Text = "印刷データ作成中..." + pCnt.ToString() + "/" + pC.ToString();

                        toolStripProgressBar1.Value = pCnt * 100 / pC;

                        // 日付範囲を超えたらループからぬける
                        nDt = sDt.AddDays(iX);
                        bool nxt = true;
                        switch (nDt.CompareTo(eDt))
                        {
                            case -1:    // 期限内
                                nxt = true;
                                break;
                            case 0:     // 期限日と同日
                                nxt = true;
                                break;
                            case 1:     // 期限日超過
                                nxt = false;
                                break;
                        }

                        if (!nxt)
                        {
                            break;
                        }

                        iX++;
                    }
                }


                // 1枚目はテンプレートシートなので印刷時には削除する
                oXls.DisplayAlerts = false;
                oXlsBook.Sheets[1].Delete();

                //System.Threading.Thread.Sleep(1000);

                // 確認のためExcelのウィンドウを表示する
                oXls.Visible = true;

                // 印刷
                oXlsBook.PrintOutEx(Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //oXlsBook.PrintOut();

                // 確認のためExcelのウィンドウを非表示する
                oXls.Visible = false;

                // 終了メッセージ 
                MessageBox.Show("終了しました");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            finally
            {
                // ウィンドウを非表示にする
                oXls.Visible = false;

                // 保存処理
                oXls.DisplayAlerts = false;

                // Bookをクローズ
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                // Excelを終了
                oXls.Quit();

                // COM オブジェクトの参照カウントを解放する 
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsMsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                oXls = null;
                oXlsBook = null;
                oxlsSheet = null;
                oxlsMsSheet = null;

                GC.Collect();

                //マウスポインタを元に戻す
                this.Cursor = Cursors.Default;

                // progreassBar非表示
                label3.Visible = false;
                toolStripProgressBar1.Visible = false;
                toolStripProgressBar1.Value = 0;
            }
        }


        ///----------------------------------------------------------------------
        /// <summary>
        ///     応援移動票印刷処理 </summary>
        /// <param name="zArray">
        ///     部署別残業理由配列 </param>
        /// <param name="pC">
        ///     印刷シート数</param>
        /// <param name="prnName">
        ///     印刷するプリンタ</param>
        /// <param name="copies">
        ///     部数</param>
        ///----------------------------------------------------------------------
        private void prnOuenSheetNonDate(object[,] zArray, int pC, string prnName, int copies)
        {
            DateTime sDt = DateTime.Parse(dateTimePicker1.Value.ToShortDateString());
            DateTime eDt = DateTime.Parse(dateTimePicker2.Value.ToShortDateString());
            DateTime nDt;

            //マウスポインタを待機にする
            this.Cursor = Cursors.WaitCursor;

            // Excel起動
            string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

            Excel.Application oXls = new Excel.Application();

            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(Properties.Settings.Default.xlsIdouSheet, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                               Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];
            Excel.Worksheet oxlsMsSheet = (Excel.Worksheet)oXlsBook.Sheets[1]; // テンプレートシート
            oxlsSheet.Select(Type.Missing);

            Excel.Range rng = null;

            int pCnt = 1;   // ページカウント
            object[,] rtnArray = null;

            try
            {
                // progressBar
                toolStripProgressBar1.Maximum = 100;
                toolStripProgressBar1.Minimum = 0;
                toolStripProgressBar1.Visible = true;

                // 部署ループ
                foreach (DataGridViewRow r in dataGridView1.Rows)
                {
                    // チェックされている部署を対象とする
                    if (dataGridView1[ColChk, r.Index].Value.ToString() == "False")
                    {
                        continue;
                    }

                    int iX = 1;

                    // 日付初期化
                    nDt = DateTime.Parse(dateTimePicker1.Value.ToShortDateString());

                    // 部署コード取得
                    string bCode = dataGridView1[ColCode, r.Index].Value.ToString().Trim();

                    // テンプレートシートを追加する
                    pCnt++;
                    oxlsMsSheet.Copy(Type.Missing, oXlsBook.Sheets[pCnt - 1]);
                    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt];

                    // シートのセルを一括して配列に取得します
                    rng = oxlsMsSheet.Range[oxlsMsSheet.Cells[1, 1], oxlsMsSheet.Cells[oxlsMsSheet.UsedRange.Rows.Count, 143]];
                    rtnArray = (object[,])rng.Value2;

                    int sp = 1;
                    rtnArray[4, 27] = sp.ToString();    // ページ
                    rtnArray[4, 36] = "1";  // 総ページ数

                    // 部署別残業理由をシート配列にセットする
                    setZangyoReasonList(ref rtnArray, zArray, bCode);

                    //// 年
                    //rtnArray[2, 2] = nDt.Year.ToString().Substring(0, 1);
                    //rtnArray[2, 5] = nDt.Year.ToString().Substring(1, 1);
                    //rtnArray[2, 8] = nDt.Year.ToString().Substring(2, 1);
                    //rtnArray[2, 11] = nDt.Year.ToString().Substring(3, 1);

                    //// 月
                    //rtnArray[2, 17] = nDt.Month.ToString().PadLeft(2, '0').Substring(0, 1);
                    //rtnArray[2, 20] = nDt.Month.ToString().PadLeft(2, '0').Substring(1, 1);

                    //// 日
                    //rtnArray[2, 26] = nDt.Day.ToString().PadLeft(2, '0').Substring(0, 1);
                    //rtnArray[2, 29] = nDt.Day.ToString().PadLeft(2, '0').Substring(1, 1);

                    // 部署コード
                    rtnArray[4, 7] = bCode.Substring(0, 1);
                    rtnArray[4, 10] = bCode.Substring(1, 1);
                    rtnArray[4, 13] = bCode.Substring(2, 1);
                    rtnArray[4, 16] = bCode.Substring(3, 1);
                    rtnArray[4, 19] = bCode.Substring(4, 1);

                    // 部署名
                    rtnArray[4, 41] = dataGridView1[ColSznm, r.Index].Value.ToString().Trim();

                    // 配列からシートセルに一括してデータをセットします
                    rng = oxlsSheet.Range[oxlsSheet.Cells[1, 1], oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 100]];
                    rng.Value2 = rtnArray;
                    
                    int iZ = 1;

                    // 部署別残業理由コードセット
                    Excel.TextBox etxt = null;
                    for (int i = 1; i <= zArray.GetLength(0); i++)
                    {
                        if (zArray[i, 1].ToString() != bCode)
                        {
                            continue;
                        }

                        if (iZ == 1)
                        {
                            etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zan1");
                            etxt.Text = zArray[i, 3].ToString();
                            etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zanC1");
                            etxt.Text = zArray[i, 2].ToString().PadLeft(2, '0');
                        }
                        else if (iZ == 2)
                        {
                            etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zan2");
                            etxt.Text = zArray[i, 3].ToString();
                            etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zanC2");
                            etxt.Text = zArray[i, 2].ToString().PadLeft(2, '0');
                        }
                        else if (iZ == 3)
                        {
                            etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zan3");
                            etxt.Text = zArray[i, 3].ToString();
                            etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zanC3");
                            etxt.Text = zArray[i, 2].ToString().PadLeft(2, '0');
                        }
                        else if (iZ == 4)
                        {
                            etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zan4");
                            etxt.Text = zArray[i, 3].ToString();
                            etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zanC4");
                            etxt.Text = zArray[i, 2].ToString().PadLeft(2, '0');
                        }
                        else if (iZ == 5)
                        {
                            etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zan5");
                            etxt.Text = zArray[i, 3].ToString();
                            etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zanC5");
                            etxt.Text = zArray[i, 2].ToString().PadLeft(2, '0');
                        }
                        else if (iZ == 6)
                        {
                            etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zan6");
                            etxt.Text = zArray[i, 3].ToString();
                            etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zanC6");
                            etxt.Text = zArray[i, 2].ToString().PadLeft(2, '0');
                        }
                        else if (iZ == 7)
                        {
                            etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zan7");
                            etxt.Text = zArray[i, 3].ToString();
                            etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zanC7");
                            etxt.Text = zArray[i, 2].ToString().PadLeft(2, '0');
                        }
                        else if (iZ == 8)
                        {
                            etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zan8");
                            etxt.Text = zArray[i, 3].ToString();
                            etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zanC8");
                            etxt.Text = zArray[i, 2].ToString().PadLeft(2, '0');
                        }
                        else if (iZ == 9)
                        {
                            etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zan9");
                            etxt.Text = zArray[i, 3].ToString();
                            etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zanC9");
                            etxt.Text = zArray[i, 2].ToString().PadLeft(2, '0');
                        }
                        else if (iZ == 10)
                        {
                            etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zan10");
                            etxt.Text = zArray[i, 3].ToString();
                            etxt = (Excel.TextBox)oxlsSheet.TextBoxes("zanC10");
                            etxt.Text = zArray[i, 2].ToString().PadLeft(2, '0');
                        }

                        iZ++;
                    }

                    // progressBar表示
                    label3.Visible = true;
                    label3.Text = "印刷データ作成中..." + pCnt.ToString() + "/" + pC.ToString();

                    toolStripProgressBar1.Value = pCnt * 100 / pC;
                    
                    iX++;
                }


                // 1枚目はテンプレートシートなので印刷時には削除する
                oXls.DisplayAlerts = false;
                oXlsBook.Sheets[1].Delete();

                //System.Threading.Thread.Sleep(1000);

                // 確認のためExcelのウィンドウを表示する
                oXls.Visible = true;

                // 印刷：プリンタ名、部数を指定可能とした 2017/08/08
                //oXlsBook.PrintOutEx(Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                oXlsBook.PrintOutEx(Type.Missing, Type.Missing, copies, true, prnName, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                //// 印刷
                //oXlsBook.PrintOut();

                // 確認のためExcelのウィンドウを非表示にする
                oXls.Visible = false;

                // 終了メッセージ 
                MessageBox.Show("終了しました");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            finally
            {
                // ウィンドウを非表示にする
                oXls.Visible = false;

                // 保存処理
                oXls.DisplayAlerts = false;

                // Bookをクローズ
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                // Excelを終了
                oXls.Quit();

                // COM オブジェクトの参照カウントを解放する 
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsMsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                oXls = null;
                oXlsBook = null;
                oxlsSheet = null;
                oxlsMsSheet = null;

                GC.Collect();

                //マウスポインタを元に戻す
                this.Cursor = Cursors.Default;

                // progreassBar非表示
                label3.Visible = false;
                toolStripProgressBar1.Visible = false;
                toolStripProgressBar1.Value = 0;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "部署別勤務体系テーブル選択";
            openFileDialog1.FileName = string.Empty;
            openFileDialog1.Filter = "Microsoft Office Excelファイル(*.xlsx)|*.xlsx|全てのファイル(*.*)|*.*";

            //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
            string fileName;
            DialogResult ret = openFileDialog1.ShowDialog();

            if (ret == System.Windows.Forms.DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
                label6.Text = openFileDialog1.FileName;
            }
            else
            {
                fileName = string.Empty;
            }
        }

    }
}
