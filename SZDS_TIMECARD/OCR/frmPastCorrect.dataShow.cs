using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Data.OleDb;
using SZDS_TIMECARD.Common;
using GrapeCity.Win.MultiRow;

namespace SZDS_TIMECARD.OCR
{
    partial class frmPastCorrect
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
            adpMn.過去勤務票ヘッダTableAdapter.Fill(dts.過去勤務票ヘッダ);
            adpMn.過去勤務票明細TableAdapter.Fill(dts.過去勤務票明細);
        }
        
        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     データを画面に表示します </summary>
        /// <param name="iX">
        ///     ヘッダデータインデックス</param>
        ///------------------------------------------------------------------------------------
        private void showOcrData(string iX)
        {
            // 非ログ書き込み状態とする
            editLogStatus = false;

            // 過去勤務票ヘッダテーブル行を取得
            DataSet1.過去勤務票ヘッダRow r = dts.過去勤務票ヘッダ.Single(a => a.ID == iX);

            // フォーム初期化
            formInitialize(dID);

            // ヘッダ情報表示
            gcMultiRow2[0, "txtYear"].Value = r.年.ToString(); ;
            gcMultiRow2[0, "txtMonth"].Value = Utility.EmptytoZero(r.月.ToString());
            gcMultiRow2[0, "txtDay"].Value = Utility.EmptytoZero(r.日.ToString());
            gcMultiRow2[0, "txtBushoCode"].Value = r.部署コード.ToString();
            gcMultiRow2[0, "txtSftCode"].Value = r.シフトコード.ToString();
            gcMultiRow2.CurrentCell = null;

            // 社員別勤怠表示
            showItem(r.ID, gcMultiRow1, r.年.ToString(), r.月.ToString());
     
            // エラー情報表示初期化
            lblErrMsg.Visible = false;
            lblErrMsg.Text = string.Empty;

            // 画像表示 ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓　2015/09/25
            ShowImage(Properties.Settings.Default.tifPath + r.画像名.ToString());
            _img = Properties.Settings.Default.tifPath + r.画像名.ToString(); // プリントイメージ

            // ログ書き込み状態とする
            editLogStatus = true;
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
        private void showItem(string hID,  GcMultiRow mr, string sYY, string sMM)
        {
            // 社員別勤務実績表示
            int mC = dts.過去勤務票明細.Count(a => a.ヘッダID == hID);
                
            // 行数を設定して表示色を初期化
            mr.Rows.Clear();
            mr.RowCount = 9;
            //mr.RowCount = mC;

            for (int i = 0; i < mC; i++)
            {
                mr.Rows[i].DefaultCellStyle.BackColor = Color.FromName("Control");
                mr.Rows[i].ReadOnly = true;    // 初期設定は編集不可とする
            }
                        
            // 行インデックス初期化
            int mRow = 0;

            foreach (var t in dts.過去勤務票明細.Where(a => a.ヘッダID == hID).OrderBy(a => a.ID))
            {
                // 表示色を初期化
                //mr.Rows[mRow].DefaultCellStyle.BackColor = Color.Empty;

                // 編集を可能とする
                //mr.Rows[mRow].ReadOnly = false;

                // 応援チェック
                if (t.応援 == global.FLGON)
                {
                    mr[mRow, "chkOuen"].Value = true;
                }
                else
                {
                    mr[mRow, "chkOuen"].Value = false;
                }

                // シフト通りチェック
                if (t.シフト通り == global.FLGON)
                {
                    mr[mRow, "chkSft"].Value = true;
                }
                else
                {
                    mr[mRow, "chkSft"].Value = false;
                }

                mr[mRow, "txtShainNum"].Value = t.社員番号.ToString();
                mr[mRow, "txtSftCode"].Value = t.シフトコード;

                gl.ChangeValueStatus = false;           // これ以下ChangeValueイベントを発生させない

                mr[mRow, "txtJiyu1"].Value = t.事由1.ToString();
                mr[mRow, "txtJiyu2"].Value = t.事由2.ToString();
                mr[mRow, "txtJiyu3"].Value = t.事由3.ToString();

                gl.ChangeValueStatus = true;            // 出勤時間はchangeValueイベントをtrueに戻す

                mr[mRow, "txtSh"].Value = t.出勤時;
                mr[mRow, "txtSm"].Value = t.出勤分;

                gl.ChangeValueStatus = false;           // これ以下ChangeValueイベントを発生させない

                mr[mRow, "txtEh"].Value = t.退勤時;
                mr[mRow, "txtEm"].Value = t.退勤分;
                mr[mRow, "txtZanRe1"].Value = t.残業理由1;
                mr[mRow, "txtZanH1"].Value = t.残業時1;
                mr[mRow, "txtZanM1"].Value = t.残業分1;
                mr[mRow, "txtZanRe2"].Value = t.残業理由2;
                mr[mRow, "txtZanH2"].Value = t.残業時2;
                mr[mRow, "txtZanM2"].Value = t.残業分2;

                gl.ChangeValueStatus = true;            // 出勤時間はchangeValueイベントをtrueに戻す

                // 取消チェック
                if (t.取消 == global.FLGON)
                {
                    mr[mRow, "chkTorikeshi"].Value = true;
                }
                else
                {
                    mr[mRow, "chkTorikeshi"].Value = false;
                }
                
                mr[mRow, "txtID"].Value = t.ID.ToString();     // 明細ＩＤ

                //// 帰宅後勤務データ登録済みか？
                //if (t.Get帰宅後勤務Rows().Count() > 0)
                //{
                //    mr[mRow, "btnCell"].Value = "○";
                //}
                //else
                //{
                //    mr[mRow, "btnCell"].Value = "";
                //}

                gl.ChangeValueStatus = true;            // ChangeValueStatusをtrueに戻す

                // 行インデックス加算
                mRow++;
            }

            //カレントセル選択状態としない
            mr.CurrentCell = null;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     画像を表示する </summary>
        /// <param name="pic">
        ///     pictureBoxオブジェクト</param>
        /// <param name="imgName">
        ///     イメージファイルパス</param>
        /// <param name="fX">
        ///     X方向のスケールファクター</param>
        /// <param name="fY">
        ///     Y方向のスケールファクター</param>
        ///------------------------------------------------------------------------------------
        private void ImageGraphicsPaint(PictureBox pic, string imgName, float fX, float fY, int RectDest, int RectSrc)
        {
            Image _img = Image.FromFile(imgName);
            Graphics g = Graphics.FromImage(pic.Image);

            // 各変換設定値のリセット
            g.ResetTransform();

            // X軸とY軸の拡大率の設定
            g.ScaleTransform(fX, fY);

            // 画像を表示する
            g.DrawImage(_img, RectDest, RectSrc);

            // 現在の倍率,座標を保持する
            gl.ZOOM_NOW = fX;
            gl.RECTD_NOW = RectDest;
            gl.RECTS_NOW = RectSrc;
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
            gcMultiRow2[0, "txtYear"].Style.BackColor = SystemColors.Window;
            gcMultiRow2[0, "txtMonth"].Style.BackColor = SystemColors.Window;
            gcMultiRow2[0, "txtDay"].Style.BackColor = SystemColors.Window;
            gcMultiRow2[0, "lblWeek"].Style.BackColor = SystemColors.Window;
            gcMultiRow2[0, "txtBushoCode"].Style.BackColor = SystemColors.Window;
            gcMultiRow2[0, "txtSftCode"].Style.BackColor = SystemColors.Window;

            //gcMultiRow2[0, "txtYear"].ReadOnly = true;
            //gcMultiRow2[0, "txtMonth"].ReadOnly = true;
            //gcMultiRow2[0, "txtDay"].ReadOnly = true;
            //gcMultiRow2[0, "txtBushoCode"].ReadOnly = true;
            //gcMultiRow2[0, "txtSftCode"].ReadOnly = true;

            //gcMultiRow1[0, "chkSft"].ReadOnly = true;
            //gcMultiRow1[0, "chkOuen"].ReadOnly = true;
            //gcMultiRow1[0, "txtSftCode"].ReadOnly = true;
            //gcMultiRow1[0, "txtShainNum"].ReadOnly = true;
            //gcMultiRow1[0, "txtJiyu1"].ReadOnly = true;
            //gcMultiRow1[0, "txtJiyu2"].ReadOnly = true;
            //gcMultiRow1[0, "txtJiyu3"].ReadOnly = true;
            //gcMultiRow1[0, "txtSh"].ReadOnly = true;
            //gcMultiRow1[0, "txtSm"].ReadOnly = true;
            //gcMultiRow1[0, "txtEh"].ReadOnly = true;
            //gcMultiRow1[0, "txtZanRe1"].ReadOnly = true;
            //gcMultiRow1[0, "txtZanH1"].ReadOnly = true;
            //gcMultiRow1[0, "txtZanM1"].ReadOnly = true;
            //gcMultiRow1[0, "txtZanRe2"].ReadOnly = true;
            //gcMultiRow1[0, "txtZanH2"].ReadOnly = true;
            //gcMultiRow1[0, "txtZanM2"].ReadOnly = true;
            //gcMultiRow1[0, "chkTorikeshi"].ReadOnly = true;
            
            lblNoImage.Visible = false;

            // ヘッダ情報
            //txtYear.ReadOnly = true;
            //txtMonth.ReadOnly = true;
            //txtSftCode.ReadOnly = true;

            gcMultiRow1.ReadOnly = true;
            gcMultiRow2.ReadOnly = true;

            //データ数表示
            gcMultiRow2[0, "lblPage"].Value = "";

            // 応援移動票データ作成画面リンク
            lnkOuen.Visible = false;
        }
    }
}
