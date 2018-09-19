using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SZDS_TIMECARD.Common;
using GrapeCity.Win.MultiRow;

namespace SZDS_TIMECARD.OCR
{
    partial class frmPastOuenCorrect
    {
        #region 単位時間フィールド
        /// <summary> 
        ///     ３０分単位 </summary>
        private int tanMin30 = 30;

        /// <summary> 
        ///     １５分単位 </summary> 
        private int tanMin15 = 15;

        /// <summary> 
        ///     １０分単位 </summary> 
        private int tanMin10 = 10;

        /// <summary> 
        ///     １分単位 </summary>
        private int tanMin1 = 1;
        #endregion

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     勤務票ヘッダと勤務票明細のデータセットにデータを読み込む </summary>
        ///------------------------------------------------------------------------------------
        private void getDataSet()
        {
            adpMn.過去応援移動票ヘッダTableAdapter.Fill(dts.過去応援移動票ヘッダ);
            adpMn.過去応援移動票明細TableAdapter.Fill(dts.過去応援移動票明細);
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     データを画面に表示します </summary>
        /// <param name="iX">
        ///     ヘッダデータインデックス</param>
        ///------------------------------------------------------------------------------------
        private void showOcrData(string dID)
        {
            // 過去応援移動票ヘッダテーブル行を取得
            DataSet1.過去応援移動票ヘッダRow r = dts.過去応援移動票ヘッダ.Single(a => a.ID == dID);

            // フォーム初期化
            formInitialize();

            // ヘッダ情報表示
            gcMultiRow1[0, "txtYear"].Value = r.年.ToString(); ;
            gcMultiRow1[0, "txtMonth"].Value = Utility.EmptytoZero(r.月.ToString());
            gcMultiRow1[0, "txtDay"].Value = Utility.EmptytoZero(r.日.ToString());
            gcMultiRow1[0, "txtBushoCode"].Value = r.部署コード.ToString();
            gcMultiRow1.CurrentCell = null;

            //global.ChangeValueStatus = false;   // チェンジバリューステータス
            //global.ChangeValueStatus = true;    // チェンジバリューステータス

            showItem(r.ID, gcMultiRow2, 1);     // 日中応援勤怠表示
            showItemZan(r.ID, gcMultiRow3, 2);  // 残業応援勤怠表示

            // エラー情報表示初期化
            lblErrMsg.Visible = false;
            lblErrMsg.Text = string.Empty;

            // 画像表示 ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓　2015/09/25
            ShowImage(Properties.Settings.Default.tifPath  + r.画像名.ToString());
            _img = Properties.Settings.Default.tifPath + r.画像名.ToString();

            // ログ書き込み状態とする
            //editLogStatus = true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     勤怠明細表示 </summary>
        /// <param name="hID">
        ///     ヘッダID</param>
        /// <param name="sYY">
        ///     年</param>
        /// <param name="sMM">
        ///     月</param>
        ///------------------------------------------------------------------------------------
        private void showItem(string hID, GcMultiRow mr, int dKbn)
        {
            // 社員別勤務実績表示
            int mC = dts.過去応援移動票明細.Count(a => a.ヘッダID == hID && a.データ区分 == 1);

            // 行数を設定して表示色を初期化
            mr.Rows.Clear();
            mr.RowCount = 5;

            for (int i = 0; i < mC; i++)
            {
                mr.Rows[i].DefaultCellStyle.BackColor = Color.FromName("Control");
                mr.Rows[i].ReadOnly = true;    // 初期設定は編集不可とする
            }

            // 行インデックス初期化
            int mRow = 0;

            // 日中応援
            foreach (var t in dts.過去応援移動票明細.Where(a => a.ヘッダID == hID && a.データ区分 == dKbn).OrderBy(a => a.ID))
            {
                mr[mRow, "txtShainNum"].Value = t.社員番号.ToString();

                gl.ChangeValueStatus = false;           // これ以下ChangeValueイベントを発生させない

                mr[mRow, "txtLineNum"].Value = t.ライン.ToString();
                mr[mRow, "txtBmn"].Value = t.部門.ToString();
                mr[mRow, "txtHin"].Value = t.製品群.ToString();
                mr[mRow, "txtOh"].Value = t.応援時;
                mr[mRow, "txtOm"].Value = t.応援分;

                gl.ChangeValueStatus = true;

                // 取消チェック
                if (t.取消 == global.FLGON)
                {
                    mr[mRow, "chkTorikeshi"].Value = true;
                }
                else
                {
                    mr[mRow, "chkTorikeshi"].Value = false;
                }

                gl.ChangeValueStatus = false;

                mr[mRow, "txtID"].Value = t.ID.ToString();     // 明細ＩＤ

                gl.ChangeValueStatus = true;            // ChangeValueStatusをtrueに戻す

                // 行インデックス加算
                mRow++;
            }

            //カレントセル選択状態としない
            mr.CurrentCell = null;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     勤怠明細表示 </summary>
        /// <param name="hID">
        ///     ヘッダID</param>
        /// <param name="sYY">
        ///     年</param>
        /// <param name="sMM">
        ///     月</param>
        ///------------------------------------------------------------------------------------
        private void showItemZan(string hID, GcMultiRow mr, int dKbn)
        {
            // 社員別勤務実績表示
            int mC = dts.過去応援移動票明細.Count(a => a.ヘッダID == hID && a.データ区分 == 2);

            // 行数を設定して表示色を初期化
            mr.Rows.Clear();
            mr.RowCount = 5;

            for (int i = 0; i < mC; i++)
            {
                mr.Rows[i].DefaultCellStyle.BackColor = Color.FromName("Control");
                mr.Rows[i].ReadOnly = true;    // 初期設定は編集不可とする
            }

            // 行インデックス初期化
            int mRow = 0;

            // 残業応援
            foreach (var t in dts.過去応援移動票明細.Where(a => a.ヘッダID == hID && a.データ区分 == dKbn).OrderBy(a => a.ID))
            {
                mr[mRow, "txtShainNum"].Value = t.社員番号.ToString();

                gl.ChangeValueStatus = false;           // これ以下ChangeValueイベントを発生させない

                mr[mRow, "txtLineNum"].Value = t.ライン.ToString();
                mr[mRow, "txtBmn"].Value = t.部門.ToString();
                mr[mRow, "txtHin"].Value = t.製品群.ToString();
                mr[mRow, "txtZanRe1"].Value = t.残業理由1;
                mr[mRow, "txtZanH1"].Value = t.残業時1;
                mr[mRow, "txtZanM1"].Value = t.残業分1;
                mr[mRow, "txtZanRe2"].Value = t.残業理由2;
                mr[mRow, "txtZanH2"].Value = t.残業時2;
                mr[mRow, "txtZanM2"].Value = t.残業分2;

                gl.ChangeValueStatus = true;

                // 取消チェック
                if (t.取消 == global.FLGON)
                {
                    mr[mRow, "chkTorikeshi"].Value = true;
                }
                else
                {
                    mr[mRow, "chkTorikeshi"].Value = false;
                }

                gl.ChangeValueStatus = false;

                mr[mRow, "txtID"].Value = t.ID.ToString();     // 明細ＩＤ

                gl.ChangeValueStatus = true;            // ChangeValueStatusをtrueに戻す

                // 行インデックス加算
                mRow++;
            }

            //カレントセル選択状態としない
            mr.CurrentCell = null;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     フォーム表示初期化 </summary>
        /// <param name="sID">
        ///     過去データ表示時のヘッダID</param>
        /// <param name="cIx">
        ///     応援移動票ヘッダカレントレコードインデックス</param>
        ///------------------------------------------------------------------------------------
        private void formInitialize()
        {
            // 表示色設定
            gcMultiRow1[0, "txtYear"].Style.BackColor = SystemColors.Window;
            gcMultiRow1[0, "txtMonth"].Style.BackColor = SystemColors.Window;
            gcMultiRow1[0, "txtDay"].Style.BackColor = SystemColors.Window;
            gcMultiRow1[0, "lblWeek"].Style.BackColor = SystemColors.Window;
            gcMultiRow1[0, "txtBushoCode"].Style.BackColor = SystemColors.Window;

            gcMultiRow1.ReadOnly = true;
            gcMultiRow2.ReadOnly = true;
            gcMultiRow3.ReadOnly = true;

            //データ数表示
            gcMultiRow1[0, "lblPage"].Value = "";

            // 勤怠データＩ／Ｐ票データ作成画面リンクボタン
            lnkIP.Visible = false;
        }
    }
}
