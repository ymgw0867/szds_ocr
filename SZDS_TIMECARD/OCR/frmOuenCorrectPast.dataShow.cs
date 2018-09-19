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
    partial class frmOuenCorrectPast
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
        ///     過去勤務票ヘッダと過去勤務票明細のデータセットにデータを読み込む </summary>
        ///------------------------------------------------------------------------------------
        private void getDataSet()
        {
            iphAdp.Fill(dts.過去勤務票ヘッダ);
            ipmAdp.Fill(dts.過去勤務票明細);
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     応援移動票ヘッダと応援移動票明細のデータセットにデータを読み込む </summary>
        ///------------------------------------------------------------------------------------
        private void getOuenDataSet()
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
        private void showOcrData(string iX)
        {
            // 過去応援移動票ヘッダテーブル行を取得
            DataSet1.過去応援移動票ヘッダRow r = dts.過去応援移動票ヘッダ.Single(a => a.ID == iX);

            // フォーム初期化
            formInitialize(dID);

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
            ShowImage(Properties.Settings.Default.tifPath + r.画像名.ToString());
            _img = Properties.Settings.Default.tifPath + r.画像名.ToString(); // プリントイメージ

            // 確認チェック
            if (r.確認 == global.flgOff)
            {
                checkBox1.Checked = false;
            }
            else
            {
                checkBox1.Checked = true;
            }

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
            mr.RowCount = mC;

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
                // 表示色を初期化
                mr.Rows[mRow].DefaultCellStyle.BackColor = Color.Empty;

                // 編集を可能とする
                mr.Rows[mRow].ReadOnly = false;
                
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
            mr.RowCount = mC;

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
                // 表示色を初期化
                mr.Rows[mRow].DefaultCellStyle.BackColor = Color.Empty;

                // 編集を可能とする
                mr.Rows[mRow].ReadOnly = false;

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
        ///     勤務票ヘッダカレントレコードインデックス</param>
        ///------------------------------------------------------------------------------------
        private void formInitialize(string sID)
        {
            // 表示色設定
            gcMultiRow1[0, "txtYear"].Style.BackColor = SystemColors.Window;
            gcMultiRow1[0, "txtMonth"].Style.BackColor = SystemColors.Window;
            gcMultiRow1[0, "txtDay"].Style.BackColor = SystemColors.Window;
            gcMultiRow1[0, "lblWeek"].Style.BackColor = SystemColors.Window;
            gcMultiRow1[0, "txtBushoCode"].Style.BackColor = SystemColors.Window;

            lblNoImage.Visible = false;

            gcMultiRow1.ReadOnly = true;
            gcMultiRow2.ReadOnly = false;
            gcMultiRow3.ReadOnly = false;
            
            // その他のボタンを無効とする
            lnkErrCheck.Visible = true;

            //データ数表示
            gcMultiRow1[0, "lblPage"].Value = "";

            // 確認チェック欄
            checkBox1.BackColor = SystemColors.Control;
            checkBox1.Checked = false;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     エラー表示 </summary>
        /// <param name="ocr">
        ///     OCRDATAクラス</param>
        ///------------------------------------------------------------------------------------
        private void ErrShow(OCRData ocr)
        {
            if (ocr._errNumber != ocr.eNothing)
            {
                // グリッドビューCellEnterイベント処理は実行しない
                gridViewCellEnterStatus = false;

                lblErrMsg.Visible = true;
                lblErrMsg.Text = ocr._errMsg;

                // 対象年月
                if (ocr._errNumber == ocr.eDataCheck)
                {
                    checkBox1.BackColor = Color.Yellow;
                    checkBox1.Focus();
                }

                // 対象年月
                if (ocr._errNumber == ocr.eYearMonth)
                {
                    gcMultiRow1[0, "txtYear"].Style.BackColor = Color.Yellow;
                    gcMultiRow1[0, "txtMonth"].Style.BackColor = Color.Yellow;
                    gcMultiRow1[0, "txtDay"].Style.BackColor = Color.Yellow;
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtYear"];
                    gcMultiRow1.BeginEdit(true);
                }

                // 部署コード
                if (ocr._errNumber == ocr.eBushoCode)
                {
                    gcMultiRow1[0, "txtBushoCode"].Style.BackColor = Color.Yellow;
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[0, "txtBushoCode"];
                    gcMultiRow1.BeginEdit(true);
                }
                
                // 社員番号
                if (ocr._errNumber == ocr.eShainNo)
                {
                    gcMultiRow2[ocr._errRow, "txtShainNum"].Style.BackColor = Color.Yellow;
                    gcMultiRow2.Focus();
                    gcMultiRow2.CurrentCell = gcMultiRow2[ocr._errRow, "txtShainNum"];
                    gcMultiRow2.BeginEdit(true);
                }

                if (ocr._errNumber == ocr.eShainNo2)
                {
                    gcMultiRow3[ocr._errRow, "txtShainNum"].Style.BackColor = Color.Yellow;
                    gcMultiRow3.Focus();
                    gcMultiRow3.CurrentCell = gcMultiRow3[ocr._errRow, "txtShainNum"];
                    gcMultiRow3.BeginEdit(true);
                }

                // ライン
                if (ocr._errNumber == ocr.eLine)
                {
                    gcMultiRow2[ocr._errRow, "txtLineNum"].Style.BackColor = Color.Yellow;
                    gcMultiRow2.Focus();
                    gcMultiRow2.CurrentCell = gcMultiRow2[ocr._errRow, "txtLineNum"];
                    gcMultiRow2.BeginEdit(true);
                }

                if (ocr._errNumber == ocr.eLine2)
                {
                    gcMultiRow3[ocr._errRow, "txtLineNum"].Style.BackColor = Color.Yellow;
                    gcMultiRow3.Focus();
                    gcMultiRow3.CurrentCell = gcMultiRow3[ocr._errRow, "txtLineNum"];
                    gcMultiRow3.BeginEdit(true);
                }

                // 部門
                if (ocr._errNumber == ocr.eBmn)
                {
                    gcMultiRow2[ocr._errRow, "txtBmn"].Style.BackColor = Color.Yellow;
                    gcMultiRow2.Focus();
                    gcMultiRow2.CurrentCell = gcMultiRow2[ocr._errRow, "txtBmn"];
                    gcMultiRow2.BeginEdit(true);
                }

                if (ocr._errNumber == ocr.eBmn2)
                {
                    gcMultiRow3[ocr._errRow, "txtBmn"].Style.BackColor = Color.Yellow;
                    gcMultiRow3.Focus();
                    gcMultiRow3.CurrentCell = gcMultiRow3[ocr._errRow, "txtBmn"];
                    gcMultiRow3.BeginEdit(true);
                }

                // 製品群
                if (ocr._errNumber == ocr.eHin)
                {
                    gcMultiRow2[ocr._errRow, "txtHin"].Style.BackColor = Color.Yellow;
                    gcMultiRow2.Focus();
                    gcMultiRow2.CurrentCell = gcMultiRow2[ocr._errRow, "txtHin"];
                    gcMultiRow2.BeginEdit(true);
                }

                if (ocr._errNumber == ocr.eHin2)
                {
                    gcMultiRow3[ocr._errRow, "txtHin"].Style.BackColor = Color.Yellow;
                    gcMultiRow3.Focus();
                    gcMultiRow3.CurrentCell = gcMultiRow3[ocr._errRow, "txtHin"];
                    gcMultiRow3.BeginEdit(true);
                }

                // 応援分
                if (ocr._errNumber == ocr.eOuenM)
                {
                    gcMultiRow2[ocr._errRow, "txtOm"].Style.BackColor = Color.Yellow;
                    gcMultiRow2.Focus();
                    gcMultiRow2.CurrentCell = gcMultiRow2[ocr._errRow, "txtOm"];
                    gcMultiRow2.BeginEdit(true);
                }

                // 残業理由１
                if (ocr._errNumber == ocr.eZanRe1)
                {
                    gcMultiRow3[ocr._errRow, "txtZanRe1"].Style.BackColor = Color.Yellow;
                    gcMultiRow3.Focus();
                    gcMultiRow3.CurrentCell = gcMultiRow3[ocr._errRow, "txtZanRe1"];
                    gcMultiRow3.BeginEdit(true);
                }

                // 残業時１
                if (ocr._errNumber == ocr.eZanH1)
                {
                    gcMultiRow3[ocr._errRow, "txtZanH1"].Style.BackColor = Color.Yellow;
                    gcMultiRow3.Focus();
                    gcMultiRow3.CurrentCell = gcMultiRow3[ocr._errRow, "txtZanH1"];
                    gcMultiRow3.BeginEdit(true);
                }

                // 残業分１
                if (ocr._errNumber == ocr.eZanM1)
                {
                    gcMultiRow3[ocr._errRow, "txtZanM1"].Style.BackColor = Color.Yellow;
                    gcMultiRow3.Focus();
                    gcMultiRow3.CurrentCell = gcMultiRow3[ocr._errRow, "txtZanM1"];
                    gcMultiRow3.BeginEdit(true);
                }

                // 残業理由２
                if (ocr._errNumber == ocr.eZanRe2)
                {
                    gcMultiRow3[ocr._errRow, "txtZanRe2"].Style.BackColor = Color.Yellow;
                    gcMultiRow3.Focus();
                    gcMultiRow3.CurrentCell = gcMultiRow3[ocr._errRow, "txtZanRe2"];
                    gcMultiRow3.BeginEdit(true);
                }

                // 残業時２
                if (ocr._errNumber == ocr.eZanH2)
                {
                    gcMultiRow3[ocr._errRow, "txtZanH2"].Style.BackColor = Color.Yellow;
                    gcMultiRow3.Focus();
                    gcMultiRow3.CurrentCell = gcMultiRow3[ocr._errRow, "txtZanH2"];
                    gcMultiRow3.BeginEdit(true);
                }

                // 残業分２
                if (ocr._errNumber == ocr.eZanM2)
                {
                    gcMultiRow3[ocr._errRow, "txtZanM2"].Style.BackColor = Color.Yellow;
                    gcMultiRow3.Focus();
                    gcMultiRow3.CurrentCell = gcMultiRow3[ocr._errRow, "txtZanM2"];
                    gcMultiRow3.BeginEdit(true);
                }

                // 勤怠データＩ／Ｐ票データとのチェック
                if (ocr._errNumber == ocr.eOuenIP)
                {
                    gcMultiRow2[ocr._errRow, "txtShainNum"].Style.BackColor = Color.Yellow;
                    gcMultiRow2.Focus();
                    gcMultiRow2.CurrentCell = gcMultiRow2[ocr._errRow, "txtShainNum"];
                    gcMultiRow2.BeginEdit(true);
                    //lnkIP.Visible = true;
                }

                if (ocr._errNumber == ocr.eOuenIP2)
                {
                    gcMultiRow3[ocr._errRow, "txtShainNum"].Style.BackColor = Color.Yellow;
                    gcMultiRow3.Focus();
                    gcMultiRow3.CurrentCell = gcMultiRow3[ocr._errRow, "txtShainNum"];
                    gcMultiRow3.BeginEdit(true);
                    //lnkIP.Visible = true;
                }

                // グリッドビューCellEnterイベントステータスを戻す
                gridViewCellEnterStatus = true;
            }
        }
    }
}
