using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using SZDS_TIMECARD.Common;

namespace SZDS_TIMECARD.OCR
{
    partial class frmPastData
    {
        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     データを画面に表示します </summary>
        /// <param name="iX">
        ///     ヘッダデータインデックス</param>
        ///------------------------------------------------------------------------------------
        private void showOcrData()
        {
            // 過去勤務票ヘッダテーブル行を取得
            DataSet1.過去勤務票ヘッダRow r = (DataSet1.過去勤務票ヘッダRow)dts.過去勤務票ヘッダ.FindByID(dID);

            // フォーム初期化
            formInitialize(dID);

            // ヘッダ情報表示
            lblTime.Text = "OCR：" + r.ID.Substring(0, 4) + "/" + r.ID.Substring(4, 2) + "/" + r.ID.Substring(6, 2) + " " +
                           r.ID.Substring(8, 2) + ":" + r.ID.Substring(10, 2) + ":" + r.ID.Substring(12, 2);
 
            txtYear.Text = (r.年 - Properties.Settings.Default.rekiHosei).ToString();
            txtMonth.Text = Utility.EmptytoZero(r.月.ToString());
            txtDay.Text = Utility.EmptytoZero(r.日.ToString());
            
            txtTaikeiCode.Text = r.シフトコード.ToString();

            //global.ChangeValueStatus = false;   // チェンジバリューステータス
            //global.ChangeValueStatus = true;    // チェンジバリューステータス

            // 社員別勤怠表示
            showItem(r.ID, dGV);
     
            // エラー情報表示初期化
            lblErrMsg.Visible = false;
            lblErrMsg.Text = string.Empty;

            // 画像表示
            ShowImage(Properties.Settings.Default.tifPath + r.画像名.ToString());
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     勤怠明細表示 </summary>
        /// <param name="hID">
        ///     ヘッダID</param>
        /// <param name="dGV">
        ///     データグリッドビューオブジェクト</param>
        ///------------------------------------------------------------------------------------
        private void showItem(string hID, DataGridView dGV)
        {
            // 社員別勤務実績表示
            var h = dts.過去勤務票明細.Where(a => a.ヘッダID == hID).OrderBy(a => a.ID);
                
            // 行数を設定して表示色を初期化
            dGV.Rows.Clear();
            dGV.RowCount = h.Count();
            for (int i = 0; i < h.Count(); i++)
            {
                dGV.Rows[i].DefaultCellStyle.BackColor = Color.FromName("Control");
                dGV.Rows[i].ReadOnly = true;    // 初期設定は編集不可とする
            }

            //// タテヨコ方向に合わせて高さを調整する
            //if (h.Count() < 17)
            //{
            //    //dGV.Height = 322;
            //    dGV.Height = 382;
            //}
            //else
            //{
            //    //dGV.Height = 362;
            //    dGV.Height = 430;
            //}

            // 行インデックス初期化
            int mRow = 0;
            foreach (var t in h)
            {
                // 表示色を初期化
                dGV.Rows[mRow].DefaultCellStyle.BackColor = Color.Empty;

                // 編集を可能とする
                dGV.Rows[mRow].ReadOnly = false;

                // 取消チェック 2015/03/10
                if (t.取消 == global.FLGON)
                {
                    dGV[cCheck, mRow].Value = true;
                }
                else
                {
                    dGV[cCheck, mRow].Value = false;
                }
                
                dGV[cShainNum, mRow].Value = t.社員番号;

                gl.ChangeValueStatus = false;           // これ以下ChangeValueイベントを発生させない

                dGV[cKinmu, mRow].Value = t.事由1;

                if (Utility.NulltoStr(t.残業時1) != string.Empty && Utility.NulltoStr(t.残業時1) != string.Empty)
                {
                    dGV[cZH, mRow].Value = Utility.NulltoStr(t.残業時1).PadLeft(1, '0') + "." + Utility.NulltoStr(t.残業分1).PadLeft(1, '0') + "h";
                }
                else
                {
                    dGV[cZH, mRow].Value = string.Empty;
                }

                if (Utility.NulltoStr(t.残業時2) != string.Empty && Utility.NulltoStr(t.残業時2) != string.Empty)
                {
                    dGV[cSIH, mRow].Value = Utility.NulltoStr(t.残業時2).PadLeft(1, '0') + "." + Utility.NulltoStr(t.残業分2).PadLeft(1, '0') + "h";
                }
                else
                {
                    dGV[cSIH, mRow].Value = string.Empty;
                }

                //dGV[cZM, mRow].Value = Utility.NulltoStr(t.残業分1);
                //dGV[cSIH, mRow].Value = Utility.NulltoStr(t.残業時2) + "." + Utility.NulltoStr(t.残業分2) + "h";
                //dGV[cSIM, mRow].Value = Utility.NulltoStr(t.残業分2);
                dGV[cSH, mRow].Value = Utility.NulltoStr(t.出勤時);
                dGV[cSM, mRow].Value = Utility.NulltoStr(t.出勤分);
                dGV[cEH, mRow].Value = Utility.NulltoStr(t.退勤時);
                dGV[cEM, mRow].Value = Utility.NulltoStr(t.退勤分);
                dGV[cID, mRow].Value = Utility.NulltoStr(t.ID);     // 明細ＩＤ

                gl.ChangeValueStatus = true;            // ChangeValueStatusをtrueに戻す
                                       
                // 行インデックス加算
                mRow++;
            }

            //カレントセル選択状態としない
            dGV.CurrentCell = null;
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
        ///------------------------------------------------------------------------------------
        private void formInitialize(string sID)
        {
            // テキストボックス表示色設定
            txtYear.BackColor = Color.White;
            txtMonth.BackColor = Color.White;
            txtDay.BackColor = Color.White;
            txtWeekDay.BackColor = Color.White;
            txtTaikeiCode.BackColor = Color.White;

            txtYear.ForeColor = Color.Navy;
            txtMonth.ForeColor = Color.Navy;
            txtDay.ForeColor = Color.Navy;
            txtWeekDay.ForeColor = Color.Navy;
            txtTaikeiCode.ForeColor = Color.Navy;

            // ヘッダ情報表示欄
            txtYear.Text = string.Empty;
            txtMonth.Text = string.Empty;
            txtDay.Text = string.Empty;
            txtWeekDay.Text = string.Empty;
            txtTaikeiCode.Text = string.Empty;
            lblNoImage.Visible = false;

            // ヘッダ情報
            txtYear.ReadOnly = true;
            txtMonth.ReadOnly = true;
            txtDay.ReadOnly = true;
            txtTaikeiCode.ReadOnly = true;

            //データ数表示
            lblPage.Text = string.Empty;
        }
    }
}
