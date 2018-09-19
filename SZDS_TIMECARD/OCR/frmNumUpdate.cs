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

namespace SZDS_TIMECARD.OCR
{
    public partial class frmNumUpdate : Form
    {
        public frmNumUpdate(string dbName)
        {
            InitializeComponent();

            // テーブルアダプターマネージャーに割り付け
            mAdp.過去勤務票明細TableAdapter = adp;
            mAdp.帰宅後勤務TableAdapter = kAdp;
            mAdp.過去応援移動票明細TableAdapter = eAdp;

            mAdp.過去勤務票明細TableAdapter.Fill(dts.過去勤務票明細);
            mAdp.帰宅後勤務TableAdapter.Fill(dts.帰宅後勤務);
            mAdp.過去応援移動票明細TableAdapter.Fill(dts.過去応援移動票明細);

            _dbName = dbName;
        }

        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.TableAdapterManager mAdp = new DataSet1TableAdapters.TableAdapterManager();
        DataSet1TableAdapters.過去勤務票明細TableAdapter adp = new DataSet1TableAdapters.過去勤務票明細TableAdapter();
        DataSet1TableAdapters.帰宅後勤務TableAdapter kAdp = new DataSet1TableAdapters.帰宅後勤務TableAdapter();
        DataSet1TableAdapters.過去応援移動票明細TableAdapter eAdp = new DataSet1TableAdapters.過去応援移動票明細TableAdapter();

        string _dbName = string.Empty;

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void frmNumUpdate_Load(object sender, EventArgs e)
        {
            // フォーム最大値
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小値
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // 画面初期化
            txtNumOld.Text = string.Empty;
            txtNumNew.Text = string.Empty;
            lblNameOld.Text = string.Empty;
            lblNameNew.Text = string.Empty;
            linkLabel4.Enabled = false;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void txtNumNew_TextChanged(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            lblNameNew.Text = string.Empty;

            if (txtNumNew.Text == txtNumOld.Text)
            {
                MessageBox.Show("同じ社員番号は指定できません","確認",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                return;
            }

            string sName = string.Empty;
            if (getBugyoName(txtNumNew.Text, out sName))
            {
                lblNameNew.Text = sName;
            }
            else
            {
                MessageBox.Show("該当する社員が就業奉行に登録されていません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtNumNew.Focus();
            }
        }

        private bool getBugyoName(string sNum, out string sName)
        {
            bool rtn = false;
            sName = string.Empty;

            // 接続文字列取得
            string sc = sqlControl.obcConnectSting.get(_dbName);
            sqlControl.DataControl sdCon = new Common.sqlControl.DataControl(sc);

            string bCode = sNum.PadLeft(10, '0');
            SqlDataReader dR = sdCon.free_dsReader(Utility.getEmployee(bCode));

            while (dR.Read())
            {
                // 社員名表示
                sName = Utility.NulltoStr(dR["Name"]).Trim();
                rtn = true;
            }

            dR.Close();
            sdCon.Close();

            return rtn;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string sName = string.Empty;
            lblNameOld.Text = string.Empty;

            if (!getName(txtNumOld.Text, out sName))
            {
                MessageBox.Show("該当する社員が過去データに登録されていません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtNumOld.Focus();
            }
            else
            {
                lblNameOld.Text = sName;
            }
        }

        private bool getName(string sNum, out string sName)
        {
            bool rtn = false;
            sName = string.Empty;

            foreach (var t in dts.過去勤務票明細.Where(a => a.社員番号 == sNum || a.社員番号 == sNum.PadLeft(6, '0')))
            {
                if (t.Is社員名Null())
                {
                    sName = "氏名記載なし";
                }
                else
                {
                    sName = Utility.NulltoStr(t.社員名);
                }

                rtn = true;
                break;
            }

            return rtn;
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (!chkData())
            {
                return;
            }

            if (MessageBox.Show(lblNameNew.Text + "さんの過去データの社員番号を全て" + txtNumNew.Text + "に変更します。よろしいですか？", "更新確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            // 社員番号更新
            shainNumUpdate(txtNumOld.Text, txtNumNew.Text);

            // 閉じる
            this.Close();
        }

        private bool chkData()
        {
            bool rtn = false; ;

            if (lblNameOld.Text != lblNameNew.Text)
            {
                if (MessageBox.Show("氏名が一致していませんがよろしいですか？", "更新確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                {
                    rtn = false;
                }
                else
                {
                    rtn = true;
                }
            }
            else
            {
                rtn = true;
            }

            return rtn;
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Close();
        }

        private void frmNumUpdate_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void lblNameOld_TextChanged(object sender, EventArgs e)
        {
            if (lblNameOld.Text == string.Empty || lblNameNew.Text == string.Empty)
            {
                linkLabel4.Enabled = false;
            }
            else
            {
                linkLabel4.Enabled = true;
            }
        }

        private void lblNameNew_TextChanged(object sender, EventArgs e)
        {
            if (lblNameOld.Text == string.Empty || lblNameNew.Text == string.Empty)
            {
                linkLabel4.Enabled = false;
            }
            else
            {
                linkLabel4.Enabled = true;
            }
        }

        private void shainNumUpdate(string oldNum, string newNum)
        {
            int nCnt = 0;

            // 過去出勤簿明細
            foreach (var t in dts.過去勤務票明細.Where(a => a.社員番号 == oldNum || a.社員番号 == oldNum.PadLeft(6, '0')))
            {
                t.社員番号 = newNum.PadLeft(6, '0');
                nCnt++;
            }

            // 過去応援移動票明細
            foreach (var t in dts.過去応援移動票明細.Where(a => a.社員番号 == oldNum || a.社員番号 == oldNum.PadLeft(6, '0')))
            {
                t.社員番号 = newNum.PadLeft(6, '0');
                nCnt++;
            }

            // 帰宅後勤務
            foreach (var t in dts.帰宅後勤務.Where(a => a.社員番号 == oldNum || a.社員番号 == oldNum.PadLeft(6, '0')))
            {
                t.社員番号 = newNum.PadLeft(6, '0');
                nCnt++;
            }

            // データベース更新
            mAdp.UpdateAll(dts);

            MessageBox.Show(nCnt.ToString() + "件のデータの社員番号を更新しました", "処理完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

    }
}
