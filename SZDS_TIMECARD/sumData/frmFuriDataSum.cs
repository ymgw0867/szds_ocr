using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SZDS_TIMECARD.Common;

namespace SZDS_TIMECARD.sumData
{
    public partial class frmFuriDataSum : Form
    {
        public frmFuriDataSum()
        {
            InitializeComponent();
        }

        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.応援集計TableAdapter Adp = new DataSet1TableAdapters.応援集計TableAdapter();

        // データヘッダ項目
        const string H1 = "年";
        const string H2 = "月";
        const string H3 = "社員番号";
        const string H4 = "振替先部門コード";
        const string H5 = "時間";
        const string H6 = "分";
        const string H7 = "合計分";

        private void frmFuriDataSum_Load(object sender, EventArgs e)
        {
            //ウィンドウズ最大サイズ
            Utility.WindowsMaxSize(this, this.Size.Width, this.Size.Height);

            //ウィンドウズ最小サイズ
            Utility.WindowsMinSize(this, this.Size.Width, this.Size.Height);

            // 年月初期値 : 2017/09/19
            txtYear.Text = DateTime.Today.AddMonths(-1).Year.ToString();
            txtMonth.Text = DateTime.Today.AddMonths(-1).Month.ToString();            
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
            if (txtYear.Text == string.Empty)
            {
                MessageBox.Show("対象年を入力してください", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtYear.Focus();
                return;
            }

            if (txtMonth.Text == string.Empty)
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

            if (MessageBox.Show("社員別給与振替データを出力します。よろしいですか。","確認",MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            // メイン処理
            main(Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text)); 
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     メイン処理 </summary>
        /// <param name="yy">
        ///     対象年</param>
        /// <param name="mm">
        ///     対象月</param>
        ///---------------------------------------------------------------
        private void main(int yy, int mm)
        {
            // 応援振替データ作成
            if (putOuenFurikae(yy, mm, 1))
            {
                // 残業振替データ作成
                putOuenFurikae(yy, mm, 2);

                // 閉じる
                this.Close();
            }

            txtMonth.Focus();
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     社員別部門別給与振替データ作成 : 2017/09/19</summary>
        /// <param name="yy">
        ///     対象年</param>
        /// <param name="mm">
        ///     対象月</param>
        /// <param name="dKbn">
        ///     データ区分　１：日中応援、２：応援残業</param>
        /// <returns>
        ///     正常終了：true, 対象データなし：false</returns>
        ///------------------------------------------------------------------
        private bool putOuenFurikae(int yy, int mm, int dKbn)
        {
            bool rtn = true;
            bool pblFirstGyouFlg = true;
            string[] csvArray = null;
            int iX = 0;

            Cursor = Cursors.WaitCursor;

            try
            {
                Adp.FillByYYMM(dts.応援集計, yy, mm, yy, mm, yy, mm);

                if (dts.応援集計.Count == 0)
                {
                    MessageBox.Show("対象年月のデータがありません", "社員別振替データ作成");
                    rtn = false;
                }

                // 社員別応援先別の応援時間を集計 : 部門別 2017/09/19
                var s = dts.応援集計.Where(a => a.データ区分 == dKbn)
                                    .OrderBy(a => a.社員番号)
                                    .GroupBy(a => a.社員番号)
                                    .Select(g => new
                                    {
                                        shainNum = g.Key,
                                        bg = g.GroupBy(b => b.部門)
                                        .Select(a => new
                                        {
                                            busho = a.Key,
                                            ouenMinute = a.Sum(b => Utility.StrtoInt(Utility.NulltoStr(b.時)) * 60 + (Utility.StrtoInt(Utility.NulltoStr(b.分)) * 60 / 10))
                                        })
                                    });
                
                foreach (var t in s)
                {
                    // ヘッダファイル出力
                    if (pblFirstGyouFlg)
                    {
                        string strHd = H1 + "," + H2 + "," + H3 + "," + H4 + "," + H5 + "," + H6 + "," + H7;

                        // 配列にデータを出力
                        Array.Resize(ref csvArray, iX + 1);
                        csvArray[iX] = strHd;

                        iX++;
                        pblFirstGyouFlg = false;
                    }

                    string sNum = t.shainNum;

                    foreach (var j in t.bg)
                    {
                        string str = txtYear.Text + "," + txtMonth.Text + "," + sNum + "," + j.busho + "," + (int)(j.ouenMinute / 60) + "," + (j.ouenMinute % 60) + "," + j.ouenMinute;

                        Array.Resize(ref csvArray, iX + 1);
                        csvArray[iX] = str;
                        iX++;
                    }
                }

                if (csvArray != null)
                {
                    string fTittle = "";
                    if (dKbn == 1)
                    {
                        fTittle = txtYear.Text + "年" + txtMonth.Text + "月 部門別応援振替 ";
                    }
                    else
                    {
                        fTittle = txtYear.Text + "年" + txtMonth.Text + "月 部門別残業振替 ";
                    }

                    // 付加文字列（タイムスタンプ）
                    string newFileName = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') +
                                            DateTime.Now.Day.ToString().PadLeft(2, '0') + DateTime.Now.Hour.ToString().PadLeft(2, '0') +
                                            DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0');

                    DialogResult ret;

                    //ダイアログボックスの初期設定
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.Title = fTittle + "データ出力";
                    saveFileDialog1.OverwritePrompt = true;
                    saveFileDialog1.RestoreDirectory = true;
                    saveFileDialog1.FileName = fTittle + newFileName;
                    saveFileDialog1.Filter = "CSVファイル(*.CSV)|*.CSV|全てのファイル(*.*)|*.*";

                    //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
                    string fileName;
                    ret = saveFileDialog1.ShowDialog();

                    if (ret == System.Windows.Forms.DialogResult.OK)
                    {
                        fileName = saveFileDialog1.FileName;

                        // CSVファイル出力
                        txtFileWrite(fileName, csvArray, fTittle);
                    }

                    MessageBox.Show(iX + "件の" + fTittle + "データが出力されました", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("出力データがありませんでした", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }

            return rtn;
        }

        ///----------------------------------------------------------------------------
        /// <summary>
        ///     テキストファイルを出力する</summary>
        /// <param name="outFilePath">
        ///     出力するフォルダ</param>
        /// <param name="arrayData">
        ///     書き込む配列データ</param>
        ///----------------------------------------------------------------------------
        private void txtFileWrite(string sPath, string[] arrayData, string fTittle)
        {
            //// 付加文字列（タイムスタンプ）
            //string newFileName = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') +
            //                        DateTime.Now.Day.ToString().PadLeft(2, '0') + DateTime.Now.Hour.ToString().PadLeft(2, '0') +
            //                        DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0');

            // ファイル名
            //string outFileName = sPath + fTittle + newFileName + ".csv";
            string outFileName = sPath;

            // テキストファイル出力
            System.IO.File.WriteAllLines(outFileName, arrayData, System.Text.Encoding.GetEncoding(932));
        }
    }
}
