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

namespace SZDS_TIMECARD.sumData
{
    public partial class frmPastDatatoCsv : Form
    {
        public frmPastDatatoCsv(string dbName)
        {
            InitializeComponent();
            _dbName = dbName;
        }

        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.過去勤務票ヘッダTableAdapter hAdp = new DataSet1TableAdapters.過去勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.過去勤務票明細TableAdapter mAdp = new DataSet1TableAdapters.過去勤務票明細TableAdapter();
        DataSet1TableAdapters.過去応援移動票ヘッダTableAdapter ohAdp = new DataSet1TableAdapters.過去応援移動票ヘッダTableAdapter();
        DataSet1TableAdapters.過去応援移動票明細TableAdapter omAdp = new DataSet1TableAdapters.過去応援移動票明細TableAdapter();

        string[] csvArray = null;
        string _dbName = string.Empty;

        // データヘッダ項目
        const string H1 = "年月日";
        const string H2 = "社員番号";
        const string H3 = "社員名";
        const string H4 = "部署コード";
        const string H5 = "シフト";
        const string H6 = "シフト通り";
        const string H7 = "応援";
        const string H8 = "出勤時刻";
        const string H9 = "退勤時刻";
        const string H10 = "残業理由１";
        const string H11 = "残業１";
        const string H12 = "残業理由２";
        const string H13 = "残業２";
        const string H14 = "事由１";
        const string H15 = "事由２";
        const string H16 = "事由３";
        const string H17 = "ライン";
        const string H18 = "部門";
        const string H19 = "製品群";
        const string H20 = "更新年月日";

        private void frmFuriDataSum_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最大サイズ
            Utility.WindowsMaxSize(this, this.Size.Width, this.Size.Height);

            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            // 年月初期値
            txtYear.Text = DateTime.Today.Year.ToString();
            txtMonth.Text = DateTime.Today.Month.ToString();

            txtYearTo.Text = DateTime.Today.Year.ToString();
            txtMonthTo.Text = DateTime.Today.Month.ToString();

            comboBox1.SelectedIndex = 0;
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Close();
        }

        private void frmFuriDataSum_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
        }

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void lnkLblClr_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (txtYear.Text == string.Empty && txtMonth.Text != string.Empty)
            {
                MessageBox.Show("対象年を入力してください", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtYear.Focus();
                return;
            }

            if (txtYear.Text != string.Empty && txtMonth.Text == string.Empty)
            {
                MessageBox.Show("対象月を入力してください", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonth.Focus();
                return;
            }

            if (txtMonth.Text != string.Empty && (Utility.StrtoInt(txtMonth.Text) < 1 || Utility.StrtoInt(txtMonth.Text) > 12))
            {
                MessageBox.Show("対象月を正しく入力してください", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonth.Focus();
                return;
            }

            if (txtYearTo.Text == string.Empty && txtMonthTo.Text != string.Empty)
            {
                MessageBox.Show("対象年を入力してください", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtYearTo.Focus();
                return;
            }

            if (txtYearTo.Text != string.Empty && txtMonthTo.Text == string.Empty)
            {
                MessageBox.Show("対象月を入力してください", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonthTo.Focus();
                return;
            }

            if (txtMonthTo.Text != string.Empty && (Utility.StrtoInt(txtMonthTo.Text) < 1 || Utility.StrtoInt(txtMonthTo.Text) > 12))
            {
                MessageBox.Show("対象月を正しく入力してください", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonthTo.Focus();
                return;
            }

            if (MessageBox.Show("ＣＳＶデータを出力します。よろしいですか。", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }
            
            // メイン処理
            main(Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text), Utility.StrtoInt(txtYearTo.Text), Utility.StrtoInt(txtMonthTo.Text));
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     メイン処理 </summary>
        /// <param name="yy">
        ///     対象年</param>
        /// <param name="mm">
        ///     対象月</param>
        ///---------------------------------------------------------------
        private void main(int yy, int mm, int tyy, int tmm)
        {
            if (comboBox1.SelectedIndex == 0)
            {
                putIPCsvData(yy, mm, tyy, tmm);
            }
            else
            {
                putOuenCsvData(yy, mm, tyy, tmm);
            }

            MessageBox.Show("処理完了しました", "CSVデータ出力", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void putIPCsvData(int fYY, int fMM, int tYY, int tMM)
        {
            Cursor = Cursors.WaitCursor;

            // 奉行SQLServer接続文字列取得
            string sc = sqlControl.obcConnectSting.get(_dbName);
            sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            hAdp.Fill(dts.過去勤務票ヘッダ);
            mAdp.Fill(dts.過去勤務票明細);

            try
            {
                bool pblFirstGyouFlg = true;

                int fYYMM = 0;
                int tYYMM = 999999;

                if (fYY != 0)
                {
                    fYYMM = fYY * 100 + fMM;
                }

                if (tYY != 0)
                {
                    tYYMM = tYY * 100 + tMM;
                }

                var s = dts.過去勤務票明細.Where(a => a.過去勤務票ヘッダRow != null &&
                    (a.過去勤務票ヘッダRow.年 * 100 + a.過去勤務票ヘッダRow.月) >= fYYMM &&
                    (a.過去勤務票ヘッダRow.年 * 100 + a.過去勤務票ヘッダRow.月) <= tYYMM)
                     .OrderBy(a => a.過去勤務票ヘッダRow.年).ThenBy(a => a.過去勤務票ヘッダRow.月).ThenBy(a => a.過去勤務票ヘッダRow.日).ThenBy(a => a.社員番号);

                if (s.Count() == 0)
                {
                    MessageBox.Show("対象年月のデータがありません", "過去勤務データCSV出力");
                    return;
                }

                StringBuilder sb = new StringBuilder();
                int iX = 0;

                foreach (var t in s)
                {
                    // ヘッダファイル出力
                    if (pblFirstGyouFlg)
                    {
                        sb.Clear();
                        sb.Append(H1).Append(",");
                        sb.Append(H2).Append(",");
                        sb.Append(H3).Append(",");
                        sb.Append(H4).Append(",");
                        sb.Append(H5).Append(",");
                        sb.Append(H6).Append(",");
                        sb.Append(H7).Append(",");
                        sb.Append(H8).Append(",");
                        sb.Append(H9).Append(",");
                        sb.Append(H10).Append(",");
                        sb.Append(H11).Append(",");
                        sb.Append(H12).Append(",");
                        sb.Append(H13).Append(",");
                        sb.Append(H14).Append(",");
                        sb.Append(H15).Append(",");
                        sb.Append(H16).Append(",");
                        sb.Append(H17).Append(",");
                        sb.Append(H18).Append(",");
                        sb.Append(H19).Append(",");
                        sb.Append(H20);

                        // 配列にデータを出力
                        Array.Resize(ref csvArray, iX + 1);
                        csvArray[iX] = sb.ToString();

                        iX++;
                        pblFirstGyouFlg = false;
                    }

                    sb.Clear();
                    sb.Append(t.過去勤務票ヘッダRow.年.ToString() + "/" + t.過去勤務票ヘッダRow.月.ToString().PadLeft(2, '0') + "/" + t.過去勤務票ヘッダRow.日.ToString().PadLeft(2, '0')).Append(",");
                    //sb.Append(t.過去勤務票ヘッダRow.月.ToString()).Append(",");
                    //sb.Append(t.過去勤務票ヘッダRow.日.ToString()).Append(",");
                    sb.Append(t.社員番号).Append(",");
                    sb.Append(t.社員名).Append(",");
                    sb.Append(t.過去勤務票ヘッダRow.部署コード).Append(",");

                    string kinmuTaikei = string.Empty;
                    if (t.シフトコード == string.Empty)
                    {
                        kinmuTaikei = t.過去勤務票ヘッダRow.シフトコード.ToString();
                    }
                    else
                    {
                        kinmuTaikei = t.シフトコード;
                    }

                    sb.Append(kinmuTaikei).Append(",");
                    sb.Append(t.シフト通り).Append(",");
                    sb.Append(t.応援).Append(",");

                    //出勤退出時刻
                    if (t.出勤時 == string.Empty && t.出勤分 == string.Empty &&
                        t.退勤時 == string.Empty && t.退勤分 == string.Empty)
                    {
                        // シフト通りのとき
                        string sftSt = string.Empty;
                        string sftEt = string.Empty;

                        GetSftTime(kinmuTaikei.PadLeft(4, '0'), out sftSt, out sftEt, sdCon);

                        // 勤務体系（シフト）コードの開始終了時刻をセットする
                        sb.Append(sftSt).Append(",").Append(sftEt).Append(",");
                    }
                    else
                    {
                        sb.Append(t.出勤時 + ":" + t.出勤分).Append(",");
                        sb.Append(t.退勤時 + ":" + t.退勤分).Append(",");
                    }

                    sb.Append(t.残業理由1).Append(",");
                    sb.Append(t.残業時1.PadLeft(1, '0') + "." + t.残業分1.PadLeft(1, '0')).Append(",");
                    sb.Append(t.残業理由2).Append(",");
                    sb.Append(t.残業時2.PadLeft(1, '0') + "." + t.残業分2.PadLeft(1, '0')).Append(",");
                    sb.Append(t.事由1).Append(",");
                    sb.Append(t.事由2).Append(",");
                    sb.Append(t.事由3).Append(",");

                    if (!t.IsラインNull())
                    {
                        sb.Append(t.ライン).Append(",");
                    }
                    else
                    {
                        sb.Append(",");
                    }

                    if (!t.Is部門Null())
                    {
                        sb.Append(t.部門).Append(",");
                    }
                    else
                    {
                        sb.Append(",");
                    }

                    if (!t.Is製品群Null())
                    {
                        sb.Append(t.製品群).Append(",");
                    }
                    else
                    {
                        sb.Append(",");
                    }

                    sb.Append(t.更新年月日);

                    Array.Resize(ref csvArray, iX + 1);
                    csvArray[iX] = sb.ToString();

                    iX++;
                }

                DialogResult ret;

                //ダイアログボックスの初期設定
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Title = "過去勤怠データ出力";
                saveFileDialog1.OverwritePrompt = true;
                saveFileDialog1.RestoreDirectory = true;
                saveFileDialog1.FileName = "過去勤怠票データ";
                saveFileDialog1.Filter = "CSVファイル(*.CSV)|*.CSV|全てのファイル(*.*)|*.*";

                //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
                string fileName;
                ret = saveFileDialog1.ShowDialog();

                if (ret == System.Windows.Forms.DialogResult.OK)
                {
                    fileName = saveFileDialog1.FileName;

                    // CSVファイル出力
                    csvFileWrite(fileName, csvArray);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (sdCon.Cn.State == ConnectionState.Open)
                {
                    sdCon.Close();
                }

                Cursor = Cursors.Default;
            }
        }


        private void putOuenCsvData(int fYY, int fMM, int tYY, int tMM)
        {
            Cursor = Cursors.WaitCursor;

            // 奉行SQLServer接続文字列取得
            string sc = sqlControl.obcConnectSting.get(_dbName);
            sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            ohAdp.Fill(dts.過去応援移動票ヘッダ);
            omAdp.Fill(dts.過去応援移動票明細);

            try
            {
                bool pblFirstGyouFlg = true;

                int fYYMM = 0;
                int tYYMM = 999999;

                if (fYY != 0)
                {
                    fYYMM = fYY * 100 + fMM;
                }

                if (tYY != 0)
                {
                    tYYMM = tYY * 100 + tMM;
                }

                var s = dts.過去応援移動票明細.Where(a => a.過去応援移動票ヘッダRow != null &&
                    (a.過去応援移動票ヘッダRow.年 * 100 + a.過去応援移動票ヘッダRow.月) >= fYYMM &&
                    (a.過去応援移動票ヘッダRow.年 * 100 + a.過去応援移動票ヘッダRow.月) <= tYYMM)
                     .OrderBy(a => a.過去応援移動票ヘッダRow.年).ThenBy(a => a.過去応援移動票ヘッダRow.月).ThenBy(a => a.過去応援移動票ヘッダRow.日).ThenBy(a => a.社員番号);

                if (s.Count() == 0)
                {
                    MessageBox.Show("対象年月のデータがありません", "過去応援移動票データCSV出力");
                    return;
                }

                StringBuilder sb = new StringBuilder();
                int iX = 0;

                foreach (var t in s)
                {
                    // ヘッダファイル出力
                    if (pblFirstGyouFlg)
                    {
                        sb.Clear();
                        sb.Append(H1).Append(",");
                        sb.Append(H2).Append(",");
                        sb.Append(H3).Append(",");
                        sb.Append("応援先部署,");
                        sb.Append("応援時間,");
                        sb.Append(H10).Append(",");
                        sb.Append(H11).Append(",");
                        sb.Append(H12).Append(",");
                        sb.Append(H13).Append(",");
                        sb.Append(H17).Append(",");
                        sb.Append(H18).Append(",");
                        sb.Append(H19).Append(",");
                        sb.Append(H20);

                        // 配列にデータを出力
                        Array.Resize(ref csvArray, iX + 1);
                        csvArray[iX] = sb.ToString();

                        iX++;
                        pblFirstGyouFlg = false;
                    }

                    sb.Clear();
                    sb.Append(t.過去応援移動票ヘッダRow.年.ToString() + "/" + t.過去応援移動票ヘッダRow.月.ToString().PadLeft(2, '0') + "/" + t.過去応援移動票ヘッダRow.日.ToString().PadLeft(2, '0')).Append(",");
                    sb.Append(t.社員番号).Append(",");
                    sb.Append(t.社員名).Append(",");
                    sb.Append(t.過去応援移動票ヘッダRow.部署コード).Append(",");
                    sb.Append(t.応援時.PadLeft(1, '0') + "." + t.応援分.PadLeft(1, '0')).Append(",");
                    sb.Append(t.残業理由1).Append(",");
                    sb.Append(t.残業時1.PadLeft(1, '0') + "." + t.残業分1.PadLeft(1, '0')).Append(",");
                    sb.Append(t.残業理由2).Append(",");
                    sb.Append(t.残業時2.PadLeft(1, '0') + "." + t.残業分2.PadLeft(1, '0')).Append(",");

                    if (!t.IsラインNull())
                    {
                        sb.Append(t.ライン).Append(",");
                    }
                    else
                    {
                        sb.Append(",");
                    }

                    if (!t.Is部門Null())
                    {
                        sb.Append(t.部門).Append(",");
                    }
                    else
                    {
                        sb.Append(",");
                    }

                    if (!t.Is製品群Null())
                    {
                        sb.Append(t.製品群).Append(",");
                    }
                    else
                    {
                        sb.Append(",");
                    }

                    sb.Append(t.更新年月日);

                    Array.Resize(ref csvArray, iX + 1);
                    csvArray[iX] = sb.ToString();

                    iX++;
                }

                DialogResult ret;

                //ダイアログボックスの初期設定
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Title = "過去応援移動票データ出力";
                saveFileDialog1.OverwritePrompt = true;
                saveFileDialog1.RestoreDirectory = true;
                saveFileDialog1.FileName = "過去応援移動票データ";
                saveFileDialog1.Filter = "CSVファイル(*.CSV)|*.CSV|全てのファイル(*.*)|*.*";

                //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
                string fileName;
                ret = saveFileDialog1.ShowDialog();

                if (ret == System.Windows.Forms.DialogResult.OK)
                {
                    fileName = saveFileDialog1.FileName;

                    // CSVファイル出力
                    csvFileWrite(fileName, csvArray);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (sdCon.Cn.State == ConnectionState.Open)
                {
                    sdCon.Close();
                }

                Cursor = Cursors.Default;
            }
        }

        ///----------------------------------------------------------------------------
        /// <summary>
        ///     CSVファイルを出力する</summary>
        /// <param name="outFilePath">
        ///     出力するフォルダ</param>
        /// <param name="arrayData">
        ///     書き込む配列データ</param>
        ///----------------------------------------------------------------------------
        private void csvFileWrite(string sPath, string[] arrayData)
        {
            // 付加文字列（タイムスタンプ）
            string newFileName = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') +
                                    DateTime.Now.Day.ToString().PadLeft(2, '0') + DateTime.Now.Hour.ToString().PadLeft(2, '0') +
                                    DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0');

            // ファイル名
            string outFileName = sPath;

            // csvファイル出力
            System.IO.File.WriteAllLines(outFileName, arrayData, System.Text.Encoding.GetEncoding(932));
        }


        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     対象シフトコードの開始時刻と終了時刻を取得する </summary>
        /// <param name="_dbName">
        ///     データベース名</param>
        /// <param name="sftCode">
        ///     シフトコード </param>
        /// <param name="sTime">
        ///     開始時刻</param>
        /// <param name="eTime">
        ///     終了時刻</param>
        ///----------------------------------------------------------------------------------
        private void GetSftTime(string sftCode, out string sTime, out string eTime, sqlControl.DataControl sdCon)
        {
            // 対象のシフトコード取得する
            DateTime sDt = DateTime.Now;
            DateTime eDt = DateTime.Now;

            // 勤務体系（シフト）コード情報取得
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("SELECT tbLaborSystem.LaborSystemID,tbLaborSystem.LaborSystemCode,");
            sb.Append("tbLaborSystem.LaborSystemName,tbLaborTimeSpanRule.StartTime,");
            sb.Append("tbLaborTimeSpanRule.EndTime ");
            sb.Append("FROM tbLaborSystem inner join tbLaborTimeSpanRule ");
            sb.Append("on tbLaborSystem.LaborSystemID = tbLaborTimeSpanRule.LaborSystemID ");
            sb.Append("where tbLaborTimeSpanRule.LaborTimeSpanRuleType = 1 ");
            sb.Append("and tbLaborSystem.LaborSystemCode = '").Append(sftCode).Append("'");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                sDt = DateTime.Parse(dR["StartTime"].ToString());
                eDt = DateTime.Parse(dR["EndTime"].ToString());
                break;
            }

            dR.Close();

            // 開始時刻
            sTime = sDt.Hour.ToString() + ":" + sDt.Minute.ToString().PadLeft(2, '0');

            // 終了時刻
            eTime = string.Empty;
            //if (sDt.Day < eDt.Day)
            //{
            //    // 翌日のとき
            //    eTime = "翌日";
            //}

            eTime += eDt.Hour.ToString() + ":" + eDt.Minute.ToString().PadLeft(2, '0');
        }
    }        
}
