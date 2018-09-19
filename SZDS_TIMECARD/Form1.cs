using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;    // 2018/03/22
using SZDS_TIMECARD.OCR;
using SZDS_TIMECARD.Common;

namespace SZDS_TIMECARD
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.勤務票ヘッダTableAdapter adp = new DataSet1TableAdapters.勤務票ヘッダTableAdapter();
        DataSet1TableAdapters.応援移動票ヘッダTableAdapter eAdp = new DataSet1TableAdapters.応援移動票ヘッダTableAdapter();

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            frmComSelect frm = new frmComSelect();
            frm.ShowDialog();

            if (frm._pblDbName != string.Empty)
            {
                // 選択領域のデータベース名を取得します
                string _ComName = frm._pblComName;
                string _ComDBName = frm._pblDbName;
                frm.Dispose();

                // 勤怠データＩ／Ｐ票
                frmCorrect frmC = new frmCorrect(_ComDBName, _ComName, string.Empty, true);
                frmC.ShowDialog();
            }
            else frm.Dispose();

            this.Show();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            frmComSelect frm = new frmComSelect();
            frm.ShowDialog();

            if (frm._pblDbName != string.Empty)
            {
                // 選択領域のデータベース名を取得します
                string _ComName = frm._pblComName;
                string _ComDBName = frm._pblDbName;
                frm.Dispose();

                // 応援移動票
                frmOuenCorrect frmC = new frmOuenCorrect(_ComDBName, _ComName, string.Empty, true);
                frmC.ShowDialog();
            }
            else frm.Dispose();

            this.Show();
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            frmComSelect frm = new frmComSelect();
            frm.ShowDialog();

            if (frm._pblDbName != string.Empty)
            {
                // 選択領域のデータベース名を取得します
                string _ComName = frm._pblComName;
                string _ComDBName = frm._pblDbName;
                frm.Dispose();

                // 社員名なしの過去勤務票明細、過去応援移動票明細に社員名をセットする：2018/03/22
                Utility.getNoNameRecovery(_ComDBName);

                // 過去勤怠データＩ／Ｐ票データビューワー
                frmUnSubmit frmC = new frmUnSubmit(_ComDBName, _ComName);
                frmC.ShowDialog();
            }
            else frm.Dispose();

            this.Show();
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            frmComSelect frm = new frmComSelect();
            frm.ShowDialog();

            if (frm._pblDbName != string.Empty)
            {
                // 選択領域のデータベース名を取得します
                string _ComName = frm._pblComName;
                string _ComDBName = frm._pblDbName;
                frm.Dispose();

                // 過去データ社員番号変更
                frmNumUpdate frmC = new frmNumUpdate(_ComDBName);
                frmC.ShowDialog();
            }
            else frm.Dispose();

            this.Show();
        }

        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("部署別残業計画など各種設定用エクセルシートの準備は出来ていますか", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }
            
            this.Hide();
            frmComSelect frm = new frmComSelect();
            frm.ShowDialog();

            if (frm._pblDbName != string.Empty)
            {
                // 選択領域のデータベース名を取得します
                string _ComName = frm._pblComName;
                string _ComDBName = frm._pblDbName;
                frm.Dispose();

                // 残業集計表
                //SZDS_TIMECARD.sumData.frmSumZanList frmZ = new sumData.frmSumZanList(_ComDBName);
                SZDS_TIMECARD.sumData.frmSumZanList_New frmZ = new sumData.frmSumZanList_New(_ComDBName);
                frmZ.ShowDialog();
            }
            else frm.Dispose();

            this.Show();
        }

        private void linkLabel6_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("部署別残業計画など各種設定用エクセルシートの準備は出来ていますか", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }
            
            this.Hide();
            frmComSelect frm = new frmComSelect();
            frm.ShowDialog();
            
            if (frm._pblDbName != string.Empty)
            {
                // 選択領域のデータベース名を取得します
                string _ComName = frm._pblComName;
                string _ComDBName = frm._pblDbName;
                frm.Dispose();

                // 残業推移グラフ
                //SZDS_TIMECARD.sumData.frmZanChartXls frmZ = new sumData.frmZanChartXls(_ComDBName);
                //SZDS_TIMECARD.sumData.frmZanChartXls_New frmZ = new sumData.frmZanChartXls_New(_ComDBName);
                SZDS_TIMECARD.sumData.frmZanChartXls_New201804 frmZ = new sumData.frmZanChartXls_New201804(_ComDBName);  // 2018/04/09
                frmZ.ShowDialog();
            }
            else frm.Dispose();

            this.Show();
        }

        private void linkLabel7_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();

            // 休日設定
            SZDS_TIMECARD.config.frmCalendar frm = new SZDS_TIMECARD.config.frmCalendar();
            frm.ShowDialog();

            this.Show();
        }

        private void linkLabel8_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            frmComSelect frm = new frmComSelect();
            frm.ShowDialog();

            if (frm._pblDbName != string.Empty)
            {
                // 選択領域のデータベース名を取得します
                string _ComName = frm._pblComName;
                string _ComDBName = frm._pblDbName;
                frm.Dispose();

                // 勤怠データＩ／Ｐ票・応援移動票発行
                prePrint.prePrint frmP = new prePrint.prePrint(_ComDBName, _ComName);
                frmP.ShowDialog();
            }
            else frm.Dispose();

            this.Show();
        }

        private void linkLabel9_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // OCR認証
            Hide();
            frmOCR frm = new frmOCR();
            frm.ShowDialog();
            Show();
        }

        private void linkLabel10_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // フォームを閉じる
            this.Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
        }

        private void linkLabel11_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            frmComSelect frm = new frmComSelect();
            frm.ShowDialog();

            if (frm._pblDbName != string.Empty)
            {
                // 選択領域のデータベース名を取得します
                string _ComName = frm._pblComName;
                string _ComDBName = frm._pblDbName;
                frm.Dispose();

                // 帰宅後勤務登録
                config.frmKitakugoWork frmK = new config.frmKitakugoWork(_ComDBName);
                frmK.ShowDialog();
            }
            else frm.Dispose();

            this.Show();
        }

        private void linkLabel12_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            frmComSelect frm = new frmComSelect();
            frm.ShowDialog();

            if (frm._pblDbName != string.Empty)
            {
                // 選択領域のデータベース名を取得します
                string _ComName = frm._pblComName;
                string _ComDBName = frm._pblDbName;
                frm.Dispose();

                // 勤怠表
                sumData.frmKintaiRepNew frmR = new sumData.frmKintaiRepNew(_ComDBName);
                frmR.ShowDialog();
            }
            else frm.Dispose();

            this.Show();
        }

        private void linkLabel13_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            sumData.frmFuriDataSum frm = new sumData.frmFuriDataSum();
            frm.ShowDialog();
            Show();
        }

        private void linkLabel14_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            frmComSelect frm = new frmComSelect();
            frm.ShowDialog();

            if (frm._pblDbName != string.Empty)
            {
                // 選択領域のデータベース名を取得します
                string _ComName = frm._pblComName;
                string _ComDBName = frm._pblDbName;
                frm.Dispose();

                // 社員名なしの過去勤務票明細、過去応援移動票明細に社員名をセットする：2018/03/22
                Utility.getNoNameRecovery(_ComDBName);

                // 勤怠表
                sumData.frmPastDatatoCsv frmR = new sumData.frmPastDatatoCsv(_ComDBName);
                frmR.ShowDialog();
            }
            else frm.Dispose();

            this.Show();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // フォームの最大・最小サイズ
            Utility.WindowsMaxSize(this, this.Width, this.Height);
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // バージョン情報表示
            Text = Text + "  ver " + Application.ProductVersion;

            // 帰宅後勤務、過去帰宅後勤務テーブルに出勤日、退勤日フィールドを追加 2018/03/08
            mdbAlter();
        }

        ///---------------------------------------------------------------------------------------
        /// <summary>
        ///     帰宅後勤務テーブルに「出勤日」「退勤日」フィールドを追加 : 2018/03/08 </summary>
        ///---------------------------------------------------------------------------------------
        private void mdbAlter()
        {
            // データベース接続文字列
            StringBuilder sb = new StringBuilder();
            OleDbCommand cm = new OleDbCommand();
            OleDbConnection Cn = new OleDbConnection();

            try
            {
                Cn.ConnectionString = Properties.Settings.Default.SZDSConnectionString;
                Cn.Open();

                cm.Connection = Cn;

                cm.CommandText = "ALTER TABLE 帰宅後勤務 ADD COLUMN 出勤日 TEXT(2) default ''";
                cm.ExecuteNonQuery();
                cm.CommandText = "ALTER TABLE 帰宅後勤務 ADD COLUMN 退勤日 TEXT(2) default ''";
                cm.ExecuteNonQuery();
                cm.CommandText = "ALTER TABLE 過去帰宅後勤務 ADD COLUMN 出勤日 TEXT(2) default ''";
                cm.ExecuteNonQuery();
                cm.CommandText = "ALTER TABLE 過去帰宅後勤務 ADD COLUMN 退勤日 TEXT(2) default ''";
                cm.ExecuteNonQuery();
            }
            catch (Exception)
            {

            }
            finally
            {
                if (Cn.State == ConnectionState.Open)
                {
                    Cn.Close();
                }
            }
        }

        ///--------------------------------------------------------------------------------------
        /// <summary>
        ///     社員名なしの過去勤務票明細、過去応援移動票明細に社員名をセットする：2018/03/22</summary>
        /// <param name="dbName">
        ///     会社領域データベース名</param>
        ///--------------------------------------------------------------------------------------
        //private void getNoNameRecovery(string dbName)
        //{
        //    // 接続文字列取得 2018/03/22
        //    string sc = sqlControl.obcConnectSting.get(dbName);
        //    sqlControl.DataControl sdCon = new Common.sqlControl.DataControl(sc);

        //    DataSet1TableAdapters.過去勤務票明細TableAdapter kAdp = new DataSet1TableAdapters.過去勤務票明細TableAdapter();
        //    DataSet1TableAdapters.過去応援移動票明細TableAdapter uAdp = new DataSet1TableAdapters.過去応援移動票明細TableAdapter();
        //    SqlDataReader dR = null;

        //    try
        //    {
        //        // 社員名なしの過去勤務票明細データに社員名をセットする 2018/03/22
        //        kAdp.FillByNoName(dts.過去勤務票明細);

        //        foreach (var nn in dts.過去勤務票明細)
        //        {
        //            string bCode = nn.社員番号.PadLeft(10, '0');
        //            dR = sdCon.free_dsReader(Utility.getEmployee(bCode));

        //            while (dR.Read())
        //            {
        //                // 社員名セット 2018/03/22
        //                nn.社員名 = dR["Name"].ToString().Trim();
        //            }

        //            kAdp.Update(dts.過去勤務票明細);
        //            dR.Close();
        //        }

        //        // 社員名なしの過去応援移動票明細に社員名をセットする 2018/03/22
        //        uAdp.FillByNoName(dts.過去応援移動票明細);

        //        foreach (var uu in dts.過去応援移動票明細)
        //        {
        //            string bCode = uu.社員番号.PadLeft(10, '0');
        //            dR = sdCon.free_dsReader(Utility.getEmployee(bCode));

        //            while (dR.Read())
        //            {
        //                // 社員名セット 2018/03/22
        //                uu.社員名 = dR["Name"].ToString().Trim();
        //            }

        //            uAdp.Update(dts.過去応援移動票明細);
        //            dR.Close();
        //        }

        //    }
        //    catch (Exception ex)
        //    {

        //    }
        //    finally
        //    {
        //        if (dR != null && !dR.IsClosed)
        //        {
        //            dR.Close();
        //        }

        //        if (sdCon.Cn.State == ConnectionState.Open)
        //        {
        //            sdCon.Close();
        //        }
        //    }
        //}
    }
}
