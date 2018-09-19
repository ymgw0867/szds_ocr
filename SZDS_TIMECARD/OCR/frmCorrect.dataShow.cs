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
    partial class frmCorrect
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
            adpMn.勤務票ヘッダTableAdapter.Fill(dts.勤務票ヘッダ);
            adpMn.勤務票明細TableAdapter.Fill(dts.勤務票明細);
            pAdpMn.過去勤務票ヘッダTableAdapter.Fill(dts.過去勤務票ヘッダ);
            pAdpMn.過去勤務票明細TableAdapter.Fill(dts.過去勤務票明細);
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     応援移動票ヘッダと応援移動票明細のデータセットにデータを読み込む </summary>
        ///------------------------------------------------------------------------------------
        private void getOuenDataSet()
        {
            ouenAdp.応援移動票ヘッダTableAdapter.Fill(dts.応援移動票ヘッダ);
            ouenAdp.応援移動票明細TableAdapter.Fill(dts.応援移動票明細);
            ouenAdp.過去応援移動票ヘッダTableAdapter.Fill(dts.過去応援移動票ヘッダ);
            ouenAdp.過去応援移動票明細TableAdapter.Fill(dts.過去応援移動票明細);
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     データを画面に表示します </summary>
        /// <param name="iX">
        ///     ヘッダデータインデックス</param>
        ///------------------------------------------------------------------------------------
        private void showOcrData(int iX)
        {
            // 非ログ書き込み状態とする
            editLogStatus = false;

            // 勤務票ヘッダテーブル行を取得
            DataSet1.勤務票ヘッダRow r = dts.勤務票ヘッダ.Single(a => a.ID == cID[iX]);

            // フォーム初期化
            formInitialize(dID, iX);

            // ヘッダ情報表示
            gcMultiRow2[0, "txtYear"].Value = r.年.ToString(); ;
            gcMultiRow2[0, "txtMonth"].Value = Utility.EmptytoZero(r.月.ToString());
            gcMultiRow2[0, "txtDay"].Value = Utility.EmptytoZero(r.日.ToString());
            gcMultiRow2[0, "txtBushoCode"].Value = r.部署コード.ToString();
            gcMultiRow2[0, "txtSftCode"].Value = r.シフトコード.ToString();
            gcMultiRow2.CurrentCell = null;

            //global.ChangeValueStatus = false;   // チェンジバリューステータス
            //global.ChangeValueStatus = true;    // チェンジバリューステータス

            // 社員別勤怠表示
            showItem(r.ID, gcMultiRow1, r.年.ToString(), r.月.ToString());
     
            // エラー情報表示初期化
            lblErrMsg.Visible = false;
            lblErrMsg.Text = string.Empty;

            // 画像表示 ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓　2015/09/25
            ShowImage(Properties.Settings.Default.dataPathIP + r.画像名.ToString());

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
            int mC = dts.勤務票明細.Count(a => a.ヘッダID == hID);
                
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

            foreach (var t in dts.勤務票明細.Where(a => a.ヘッダID == hID).OrderBy(a => a.ID))
            {
                // 表示色を初期化
                //mr.Rows[mRow].DefaultCellStyle.BackColor = Color.Empty;

                // 編集を可能とする
                mr.Rows[mRow].ReadOnly = false;

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

        /// --------------------------------------------------------------------------------
        /// <summary>
        ///     時間外記入チェック </summary>
        /// <param name="wkSpan">
        ///     所定労働時間 </param>
        /// <param name="wkSpanName">
        ///     勤務体系名称 </param>
        /// <param name="mRow">
        ///     グリッド行インデックス </param>
        /// <param name="TaikeiCode">
        ///     勤務体系コード </param>
        /// --------------------------------------------------------------------------------
        private void zanCheckShow(long wkSpan, string wkSpanName, int mRow, int TaikeiCode)
        {
            //Int64 s10 = 0;  // 深夜勤務時間中の10分または15分休憩時間

            //// 所定勤務時間が取得されていないとき戻る
            //if (wkSpan == 0)
            //{
            //    return;
            //}
            
            //// 所定勤務時間が取得されているとき残業時間計算チェックを行う
            //Int64 restTm = 0;

            //// 所定時間ごとの休憩時間
            ////if (wkSpanName == WKSPAN0750)
            ////{
            ////    restTm = RESTTIME0750;
            ////}
            ////else if (wkSpanName == WKSPAN0755)
            ////{
            ////    restTm = RESTTIME0755;
            ////}
            ////else if (wkSpanName == WKSPAN0800)
            ////{
            ////    restTm = RESTTIME0800;
            ////}
                
            //// 時間外勤務時間取得 2015/09/30
            //Int64 zan = getZangyoTime(mRow, (Int64)tanMin30, wkSpan, restTm, out s10, TaikeiCode);

            //// 時間外記入時間チェック 2015/09/30
            //errCheckZanTm(mRow, zan);

            //OCRData ocr = new OCRData(_dbName, bs);

            //string sh = Utility.NulltoStr(dGV[cSH, mRow].Value.ToString());
            //string sm = Utility.NulltoStr(dGV[cSM, mRow].Value.ToString());
            //string eh = Utility.NulltoStr(dGV[cEH, mRow].Value.ToString());
            //string em = Utility.NulltoStr(dGV[cEM, mRow].Value.ToString());

            //// 深夜勤務時間を取得
            //double shinyaTm = ocr.getShinyaWorkTime(sh, sm, eh, em, tanMin10, s10);

            //// 深夜勤務時間チェック
            //errCheckShinyaTm(mRow, (Int64)shinyaTm);
        }

        /// -----------------------------------------------------------------------------------
        /// <summary>
        ///     時間外勤務時間取得 </summary>
        /// <param name="m">
        ///     グリッド行インデックス</param>
        /// <param name="Tani">
        ///     丸め単位</param>
        /// <param name="ws">
        ///     所定労働時間</param>
        /// <param name="restTime">
        ///     勤務体系別の所定労働時間内の休憩時間</param>
        /// <param name="s10Rest">
        ///     勤務体系別の所定労働時間以降の休憩時間単位</param>
        /// <param name="taikeiCode">
        ///     勤務体系コード</param>
        /// <returns>
        ///     時間外勤務時間</returns>
        /// -----------------------------------------------------------------------------------
        private Int64 getZangyoTime(int m, Int64 Tani, Int64 ws, Int64 restTime, out Int64 s10Rest, int taikeiCode)
        {
            Int64 zan = 0;  // 計算後時間外勤務時間
            s10Rest = 0;    // 深夜勤務時間帯の10分休憩時間

            //DateTime cTm;
            //DateTime sTm;
            //DateTime eTm;
            //DateTime zsTm;
            //DateTime pTm;

            //if (dGV[cSH, m].Value != null && dGV[cSM, m].Value != null && dGV[cEH, m].Value != null && dGV[cEM, m].Value != null)
            //{
            //    int ss = Utility.StrtoInt(dGV[cSH, m].Value.ToString()) * 100 + Utility.StrtoInt(dGV[cSM, m].Value.ToString());
            //    int ee = Utility.StrtoInt(dGV[cEH, m].Value.ToString()) * 100 + Utility.StrtoInt(dGV[cEM, m].Value.ToString());
            //    DateTime dt = DateTime.Today;
            //    string sToday = dt.Year.ToString() + "/" + dt.Month.ToString() + "/" + dt.Day.ToString();

            //    // 始業時刻
            //    if (DateTime.TryParse(sToday + " " + dGV[cSH, m].Value.ToString() + ":" + dGV[cSM, m].Value.ToString(), out cTm))
            //    {
            //        sTm = cTm;
            //    }
            //    else return 0;

            //    // 終業時刻
            //    if (ss > ee)
            //    {
            //        // 翌日
            //        dt = DateTime.Today.AddDays(1);
            //        sToday = dt.Year.ToString() + "/" + dt.Month.ToString() + "/" + dt.Day.ToString();
            //        if (DateTime.TryParse(sToday + " " + dGV[cEH, m].Value.ToString() + ":" + dGV[cEM, m].Value.ToString(), out cTm))
            //        {
            //            eTm = cTm;
            //        }
            //        else return 0;
            //    }
            //    else
            //    {
            //        // 同日
            //        if (DateTime.TryParse(sToday + " " + dGV[cEH, m].Value.ToString() + ":" + dGV[cEM, m].Value.ToString(), out cTm))
            //        {
            //            eTm = cTm;
            //        }
            //        else return 0;
            //    }

            //    // 作業日報に記入されている始業から就業までの就業時間取得
            //    double w = Utility.GetTimeSpan(sTm, eTm).TotalMinutes - restTime;

            //    // 所定労働時間内なら時間外なし
            //    if (w <= ws)
            //    {
            //        return 0;
            //    }

            //    // 所定労働時間＋休憩時間＋10分または15分経過後の時刻を取得（時間外開始時刻）
            //    zsTm = sTm.AddMinutes(ws);          // 所定労働時間
            //    zsTm = zsTm.AddMinutes(restTime);   // 休憩時間
            //    int zSpan = 0;

            //    if (taikeiCode == 100)
            //    {
            //        zsTm = zsTm.AddMinutes(10);         // 体系コード：100 所定労働時間後の10分休憩
            //        zSpan = 130;
            //    }
            //    else if (taikeiCode == 200 || taikeiCode == 300)
            //    {
            //        zsTm = zsTm.AddMinutes(15);         // 体系コード：200,300 所定労働時間後の15分休憩
            //        zSpan = 135;
            //    }

            //    pTm = zsTm;                         // 時間外開始時刻

            //    // 該当時刻から終業時刻まで130分または135分以上あればループさせる
            //    while (Utility.GetTimeSpan(pTm, eTm).TotalMinutes > zSpan)
            //    {
            //        // 終業時刻まで2時間につき10分休憩として時間外を算出
            //        // 時間外として2時間加算
            //        zan += 120;

            //        // 130分、または135分後の時刻を取得（2時間＋10分、または15分）
            //        pTm = pTm.AddMinutes(zSpan);

            //        // 深夜勤務時間中の10分または15分休憩時間を取得する
            //        s10Rest += getShinya10Rest(pTm, eTm, zSpan - 120);
            //    }

            //    // 130分（135分）以下の時間外を加算
            //    zan += (Int64)Utility.GetTimeSpan(pTm, eTm).TotalMinutes;

            //    // 単位で丸める
            //    zan -= (zan % Tani);
            //}

            return zan;
        }


        /// --------------------------------------------------------------------
        /// <summary>
        ///     深夜勤務時間中の10分または15分休憩時間を取得する </summary>
        /// <param name="pTm">
        ///     時刻</param>
        /// <param name="eTm">
        ///     終業時刻</param>
        /// <param name="taikeiRest">
        ///     勤務体系別の休憩時間(10分または15分）</param>
        /// <returns>
        ///     休憩時間</returns>
        /// --------------------------------------------------------------------
        private int getShinya10Rest(DateTime pTm, DateTime eTm, int taikeiRest)
        {
            int restTime = 0;

            // 130(135)分後の時刻が終業時刻以内か
            TimeSpan ts = eTm.TimeOfDay;

            if (pTm <= eTm)
            {
                // 時刻が深夜時間帯か？
                if (pTm.Hour >= 22 || pTm.Hour <= 5)
                {
                    if (pTm.Hour == 22)
                    {
                        // 22時帯は22時以降の経過分を対象とします。
                        // 例）21:57～22:07のとき22時台の7分が休憩時間
                        if (pTm.Minute >= taikeiRest)
                        {
                            restTime = taikeiRest;
                        }
                        else
                        {
                            restTime = pTm.Minute;
                        }
                    }
                    else if (pTm.Hour == 5)
                    {
                        // 4時帯の経過分を対象とするので5時帯は減算します。
                        // 例）4:57～5:07のとき5時台の7分は差し引いて3分が休憩時間
                        if (pTm.Minute < taikeiRest)
                        {
                            restTime = (taikeiRest - pTm.Minute);
                        }
                    }
                    else
                    {
                        restTime = taikeiRest;
                    }
                }
            }

            return restTime;
        }


        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     時間外記入チェック </summary>
        /// <param name="m">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <param name="zan">
        ///     算出残業時間</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private void errCheckZanTm(int m, Int64 zan)
        {
            Int64 mZan = 0;

            mZan = (Utility.StrtoInt(gcMultiRow1[m, "txtZanH1"].Value.ToString()) * 60) + (Utility.StrtoInt(gcMultiRow1[m, "txtZanM1"].Value.ToString()) * 60 / 10);

            // 記入時間と計算された残業時間が不一致のとき
            if (zan != mZan)
            {
                gcMultiRow1[m, "txtZanH1"].Style.BackColor = Color.LightPink;
                gcMultiRow1[m, "txtZanH1"].Style.BackColor = Color.LightPink;
            }
            else
            {
                gcMultiRow1[m, "txtZanM1"].Style.BackColor = Color.White;
                gcMultiRow1[m, "txtZanM1"].Style.BackColor = Color.White;
            }
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
        private void formInitialize(string sID, int cIx)
        {
            // 表示色設定
            gcMultiRow2[0, "txtYear"].Style.BackColor = SystemColors.Window;
            gcMultiRow2[0, "txtMonth"].Style.BackColor = SystemColors.Window;
            gcMultiRow2[0, "txtDay"].Style.BackColor = SystemColors.Window;
            gcMultiRow2[0, "lblWeek"].Style.BackColor = SystemColors.Window;
            gcMultiRow2[0, "txtBushoCode"].Style.BackColor = SystemColors.Window;
            gcMultiRow2[0, "txtSftCode"].Style.BackColor = SystemColors.Window;
            
            lblNoImage.Visible = false;

            // 編集可否
            gcMultiRow2.ReadOnly = false;
            gcMultiRow1.ReadOnly = false;
                
            // スクロールバー設定
            hScrollBar1.Enabled = true;
            hScrollBar1.Minimum = 0;
            hScrollBar1.Maximum =  dts.勤務票ヘッダ.Count - 1;
            hScrollBar1.Value = cIx;
            hScrollBar1.LargeChange = 1;
            hScrollBar1.SmallChange = 1;

            //移動ボタン制御
            btnFirst.Enabled = true;
            btnNext.Enabled = true;
            btnBefore.Enabled = true;
            btnEnd.Enabled = true;

            //最初のレコード
            if (cIx == 0)
            {
                btnBefore.Enabled = false;
                btnFirst.Enabled = false;
            }

            //最終レコード
            if ((cIx + 1) == dts.勤務票ヘッダ.Count)
            {
                btnNext.Enabled = false;
                btnEnd.Enabled = false;
            }

            if (_eMode)
            {
                // その他のボタンを有効とする
                lnkErrCheck.Visible = true;
                lnkDataMake.Visible = true;
                lnkDel.Visible = true;
                lnkOuen.Visible = true;     // 応援移動票データ作成画面リンク
            }
            else
            {
                // 応援移動票画面から遷移のときその他のボタンを無効とする
                lnkErrCheck.Visible = false;
                lnkDataMake.Visible = false;
                lnkDel.Visible = false;
                lnkOuen.Visible = false;    // 応援移動票データ作成画面リンク
            }

            //データ数表示
            gcMultiRow2[0, "lblPage"].Value = " (" + (cI + 1).ToString() + "/" + dts.勤務票ヘッダ.Rows.Count.ToString() + ")";
            
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
                    gcMultiRow2[0, "txtYear"].Style.BackColor = Color.Yellow;
                    gcMultiRow2[0, "txtMonth"].Style.BackColor = Color.Yellow;
                    gcMultiRow2[0, "txtDay"].Style.BackColor = Color.Yellow;
                    gcMultiRow2.Focus();
                    gcMultiRow2.CurrentCell = gcMultiRow2[0, "txtYear"];
                    gcMultiRow2.BeginEdit(true);
                }

                // 部署コード
                if (ocr._errNumber == ocr.eBushoCode)
                {
                    gcMultiRow2[0, "txtBushoCode"].Style.BackColor = Color.Yellow;
                    gcMultiRow2.Focus();
                    gcMultiRow2.CurrentCell = gcMultiRow2[0, "txtBushoCode"];
                    gcMultiRow2.BeginEdit(true);
                }

                // 勤務体系（シフト）コード
                if (ocr._errNumber == ocr.eKinmuTaikeiCode)
                {
                    gcMultiRow2[0, "txtSftCode"].Style.BackColor = Color.Yellow;
                    gcMultiRow2.Focus();
                    gcMultiRow2.CurrentCell = gcMultiRow2[0, "txtSftCode"];
                    gcMultiRow2.BeginEdit(true);
                }

                // 応援チェック
                if (ocr._errNumber == ocr.eChkOuen)
                {
                    gcMultiRow1[ocr._errRow, "chkOuen"].Style.BackColor = Color.Yellow;
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "chkOuen"];
                    gcMultiRow1.BeginEdit(true);

                    lnkOuen.Visible = true;
                }

                // シフト通りチェック
                if (ocr._errNumber == ocr.eChksft)
                {
                    gcMultiRow1[ocr._errRow, "chkSft"].Style.BackColor = Color.Yellow;
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "chkSft"];
                    gcMultiRow1.BeginEdit(true);
                }

                // 社員番号
                if (ocr._errNumber == ocr.eShainNo)
                {
                    gcMultiRow1[ocr._errRow, "txtShainNum"].Style.BackColor = Color.Yellow;
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "txtShainNum"];
                    gcMultiRow1.BeginEdit(true);
                }

                // 事由
                if (ocr._errNumber == ocr.eJiyu1)
                {
                    gcMultiRow1[ocr._errRow, "txtJiyu1"].Style.BackColor = Color.Yellow;
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "txtJiyu1"];
                    gcMultiRow1.BeginEdit(true);
                }

                if (ocr._errNumber == ocr.eJiyu2)
                {
                    gcMultiRow1[ocr._errRow, "txtJiyu2"].Style.BackColor = Color.Yellow;
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "txtJiyu2"];
                    gcMultiRow1.BeginEdit(true);
                }

                if (ocr._errNumber == ocr.eJiyu3)
                {
                    gcMultiRow1[ocr._errRow, "txtJiyu3"].Style.BackColor = Color.Yellow;
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "txtJiyu3"];
                    gcMultiRow1.BeginEdit(true);
                }

                // 変更シフトコード
                if (ocr._errNumber == ocr.eSftCode)
                {
                    gcMultiRow1[ocr._errRow, "txtSftCode"].Style.BackColor = Color.Yellow;
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "txtSftCode"];
                    gcMultiRow1.BeginEdit(true);
                }
                
                // 開始時
                if (ocr._errNumber == ocr.eSH)
                {
                    gcMultiRow1[ocr._errRow, "txtSh"].Style.BackColor = Color.Yellow;
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "txtSh"];
                    gcMultiRow1.BeginEdit(true);
                }

                // 開始分
                if (ocr._errNumber == ocr.eSM)
                {
                    gcMultiRow1[ocr._errRow, "txtSm"].Style.BackColor = Color.Yellow;
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "txtSm"];
                    gcMultiRow1.BeginEdit(true);
                }

                // 終了時
                if (ocr._errNumber == ocr.eEH)
                {
                    gcMultiRow1[ocr._errRow, "txtEh"].Style.BackColor = Color.Yellow;
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "txtEh"];
                    gcMultiRow1.BeginEdit(true);
                }

                // 終了分
                if (ocr._errNumber == ocr.eEM)
                {
                    gcMultiRow1[ocr._errRow, "txtEm"].Style.BackColor = Color.Yellow;
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "txtEm"];
                    gcMultiRow1.BeginEdit(true);
                }

                // 残業理由１
                if (ocr._errNumber == ocr.eZanRe1)
                {
                    gcMultiRow1[ocr._errRow, "txtZanRe1"].Style.BackColor = Color.Yellow;
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "txtZanRe1"];
                    gcMultiRow1.BeginEdit(true);
                }

                // 残業時１
                if (ocr._errNumber == ocr.eZanH1)
                {
                    gcMultiRow1[ocr._errRow, "txtZanH1"].Style.BackColor = Color.Yellow;
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "txtZanH1"];
                    gcMultiRow1.BeginEdit(true);
                }

                // 残業分１
                if (ocr._errNumber == ocr.eZanM1)
                {
                    gcMultiRow1[ocr._errRow, "txtZanM1"].Style.BackColor = Color.Yellow;
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "txtZanM1"];
                    gcMultiRow1.BeginEdit(true);
                }

                // 残業理由２
                if (ocr._errNumber == ocr.eZanRe2)
                {
                    gcMultiRow1[ocr._errRow, "txtZanRe2"].Style.BackColor = Color.Yellow;
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "txtZanRe2"];
                    gcMultiRow1.BeginEdit(true);
                }

                // 残業時２
                if (ocr._errNumber == ocr.eZanH2)
                {
                    gcMultiRow1[ocr._errRow, "txtZanH2"].Style.BackColor = Color.Yellow;
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "txtZanH2"];
                    gcMultiRow1.BeginEdit(true);
                }

                // 残業分２
                if (ocr._errNumber == ocr.eZanM2)
                {
                    gcMultiRow1[ocr._errRow, "txtZanM2"].Style.BackColor = Color.Yellow;
                    gcMultiRow1.Focus();
                    gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "txtZanM2"];
                    gcMultiRow1.BeginEdit(true);
                }

                // 応援移動票に対応する勤怠データＩ／Ｐ票データが存在するか
                if (ocr._errNumber == ocr.eIpOuen)
                {
                    //gcMultiRow1[ocr._errRow, "chkOuen"].Style.BackColor = Color.Yellow;
                    //gcMultiRow1.Focus();
                    //gcMultiRow1.CurrentCell = gcMultiRow1[ocr._errRow, "chkOuen"];
                    //gcMultiRow1.BeginEdit(true);

                    lnkOuen.Visible = true;
                }

                // グリッドビューCellEnterイベントステータスを戻す
                gridViewCellEnterStatus = true;
            }
        }
    }
}
