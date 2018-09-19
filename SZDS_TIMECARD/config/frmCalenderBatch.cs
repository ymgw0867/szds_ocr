using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SZDS_TIMECARD.config
{
    public partial class frmCalenderBatch : Form
    {
        public frmCalenderBatch()
        {
            InitializeComponent();

            adp.Fill(dts.休日);
        }


        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.休日TableAdapter adp = new DataSet1TableAdapters.休日TableAdapter();

        private void button2_Click(object sender, EventArgs e)
        {
        }

        private void frmCalenderBatch_Load(object sender, EventArgs e)
        {

        }

        private void frmCalenderBatch_FormClosing(object sender, FormClosingEventArgs e)
        {
            Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
        }

        ///-----------------------------------------------------------------
        /// <summary>
        ///     任意の期間内の土・日曜日を全て休日登録する </summary>
        /// <returns>
        ///     登録件数</returns>
        ///-----------------------------------------------------------------
        private int batchUpdate(DateTime sDt, DateTime eDt)
        {
            DateTime nDt =  DateTime.Parse(sDt.ToShortDateString());

            int iX = 0;
            int rC = 0;

            while(true)
            {
                nDt = sDt.AddDays(iX);

                if (nDt.CompareTo(eDt) == 1)
                {
                    break;
                }

                if (nDt.DayOfWeek == DayOfWeek.Saturday)
                {
                    if (recUpdate(nDt, "土曜日"))
                    {
                        rC++;
                    }
                }
                else if (nDt.DayOfWeek == DayOfWeek.Sunday)
                {
                    if (recUpdate(nDt, "日曜日"))
                    {
                        rC++;
                    }
                }

                iX++;
            }

            return rC;
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     任意の日付を休日登録します </summary>
        /// <param name="dt">
        ///     日付</param>
        /// <param name="week">
        ///     土または日</param>
        /// <returns>
        ///     true:登録、false:未登録</returns>
        ///---------------------------------------------------------------
        private bool recUpdate(DateTime dt, string week)
        {
            // 既に同日が登録済みのときは何もしない
            if (!dataSearch(dt))
            {
                return false;
            }

            // 休日登録する
            DataSet1.休日Row r = dts.休日.New休日Row();
            r.年月日 = DateTime.Parse(dt.ToShortDateString());
            r.名称 = week;
            r.備考 = string.Empty;
            r.更新年月日 = DateTime.Now;
            dts.休日.Add休日Row(r);
            adp.Update(dts.休日);

            return true;
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

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker2.MinDate = dateTimePicker1.Value.AddDays(1);
        }

        private void lnkLblUpdate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            DateTime sDt = DateTime.Parse(dateTimePicker1.Value.ToShortDateString());
            DateTime eDt = DateTime.Parse(dateTimePicker2.Value.ToShortDateString());

            if (sDt.CompareTo(eDt) == 1)
            {
                MessageBox.Show("日付期間が正しくありません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            string msg = string.Empty;

            msg = sDt.ToShortDateString() + "～" + eDt.ToShortDateString() + "の土・日曜日を一括して休日登録します。よろしいですか";

            if (MessageBox.Show(msg, "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            int r = batchUpdate(sDt, eDt);

            MessageBox.Show(r.ToString() + "件を一括登録しました");
            Close();
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Close();
        }
    }
}
