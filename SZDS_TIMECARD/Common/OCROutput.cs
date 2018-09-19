using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace SZDS_TIMECARD.Common
{
    ///------------------------------------------------------------------
    /// <summary>
    ///     給与計算受け渡しデータ作成クラス </summary>     
    ///------------------------------------------------------------------
    class OCROutput
    {
        // 親フォーム
        Form _preForm;

        #region データテーブルインスタンス
        DataSet1.勤務票ヘッダDataTable _hTbl;
        DataSet1.勤務票明細DataTable _mTbl;
        DataSet1.帰宅後勤務DataTable _kTbl;
        #endregion

        private const string TXTFILENAME = "就業データ";

        DataSet1 _dts = new DataSet1();

        // 就業奉行汎用データヘッダ項目
        const string H1 = @"""EBAS001""";   // 社員番号
        const string H2 = @"""LTLT001""";   // 日付
        const string H3 = @"""LTLT002""";   // 勤務回
        const string H4 = @"""LTLT003""";   // 勤務体系コード
        const string H5 = @"""LTLT004""";   // 事由コード１
        const string H6 = @"""LTLT005""";   // 事由コード２
        const string H7 = @"""LTLT006""";   // 事由コード３
        const string H8 = @"""LTDT001""";   // 出勤時刻
        const string H9 = @"""LTDT002""";   // 退出時刻

        ///--------------------------------------------------------------------------
        /// <summary>
        ///     給与計算用計算用受入データ作成クラスコンストラクタ</summary>
        /// <param name="preFrm">
        ///     親フォーム</param>
        /// <param name="hTbl">
        ///     勤務票ヘッダDataTable</param>
        /// <param name="mTbl">
        ///     勤務票明細DataTable</param>
        ///--------------------------------------------------------------------------
        public OCROutput(Form preFrm, DataSet1 dts)
        {
            _preForm = preFrm;
            _dts = dts;
            _hTbl = dts.勤務票ヘッダ;
            _mTbl = dts.勤務票明細;
            _kTbl = dts.帰宅後勤務;
        }

        ///--------------------------------------------------------------------------------------
        /// <summary>
        ///     就業奉行用受入データ作成</summary>
        ///--------------------------------------------------------------------------------------     
        public void SaveData(string dbName)
        {
            #region 出力配列
            string[] arrayCsv = null;     // 出力配列
            #endregion

            #region 出力件数変数
            int sCnt = 0;   // 社員出力件数
            #endregion

            StringBuilder sb = new StringBuilder();
            Boolean pblFirstGyouFlg = true;
            string wID = string.Empty;
            string hKinmutaikei = string.Empty;

            global gl = new global();

            // 奉行SQLServer接続文字列取得
            string sc = sqlControl.obcConnectSting.get(dbName);
            sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            // 出力先フォルダがあるか？なければ作成する
            string cPath = Properties.Settings.Default.okPath;
            if (!System.IO.Directory.Exists(cPath)) System.IO.Directory.CreateDirectory(cPath);

            try
            {
                //オーナーフォームを無効にする
                _preForm.Enabled = false;

                //プログレスバーを表示する
                frmPrg frmP = new frmPrg();
                frmP.Owner = _preForm;
                frmP.Show();

                int rCnt = 1;

                // 伝票最初行フラグ
                pblFirstGyouFlg = true;

                // 勤務票データ取得
                // 取消行は対象外とする
                var s = _mTbl.Where(a => a.社員番号 != string.Empty && a.取消 == global.FLGOFF).OrderBy(a => a.ID);

                foreach (var r in s)
                {
                    // プログレスバー表示
                    frmP.Text = "就業奉行用受入データ作成中です・・・" + rCnt.ToString() + "/" + s.Count().ToString();
                    frmP.progressValue = rCnt * 100 / s.Count();
                    frmP.ProgressStep();

                    // ヘッダファイル出力
                    if (pblFirstGyouFlg == true)
                    {
                        sb.Clear();
                        sb.Append(H1).Append(",");      // 社員番号
                        sb.Append(H2).Append(",");      // 日付
                        sb.Append(H3).Append(",");      // 勤務回
                        sb.Append(H4).Append(",");      // 勤務体系コード
                        sb.Append(H5).Append(",");      // 事由１
                        sb.Append(H6).Append(",");      // 事由２
                        sb.Append(H7).Append(",");      // 事由３
                        sb.Append(H8).Append(",");      // 出勤時刻
                        sb.Append(H9);                  // 退勤時刻

                        // 配列にデータを出力
                        sCnt++;
                        Array.Resize(ref arrayCsv, sCnt);
                        arrayCsv[sCnt - 1] = sb.ToString();
                    }

                    // 勤務票明細から受入れデータを作成する
                    string uCSV = getUkeireCsv(r, out hKinmutaikei, sdCon);

                    // 配列にデータを格納します
                    sCnt++;
                    Array.Resize(ref arrayCsv, sCnt);
                    arrayCsv[sCnt - 1] = uCSV;

                    // データ件数加算
                    rCnt++;

                    // 同社員の同日の帰宅後勤務データから受入れデータを作成する
                    string kCSV = getKitakugoCsv(r);
                    if (kCSV != string.Empty)
                    {
                        // 配列にデータを格納します
                        sCnt++;
                        Array.Resize(ref arrayCsv, sCnt);
                        arrayCsv[sCnt - 1] = kCSV;
                    }

                    pblFirstGyouFlg = false;
                }

                // 勤怠CSVファイル出力
                if (arrayCsv != null) txtFileWrite(cPath, arrayCsv);

                // いったんオーナーをアクティブにする
                _preForm.Activate();

                // 進行状況ダイアログを閉じる
                frmP.Close();

                // オーナーのフォームを有効に戻す
                _preForm.Enabled = true;
            }
            catch (Exception e)
            {
                MessageBox.Show("就業奉行受入データ作成中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
            }
            finally
            {
                //if (OutData.sCom.Connection.State == ConnectionState.Open) OutData.sCom.Connection.Close();

                if (sdCon.Cn.State == System.Data.ConnectionState.Open)
                {
                    sdCon.Close();
                }
            }
        }

        ///-----------------------------------------------------------
        /// <summary>
        ///     勤務票明細から受入れデータ文字列を作成する </summary>
        /// <param name="r">
        ///     DataSet1.勤務票明細Row </param>
        /// <param name="dbName">
        ///     データベース名</param>
        /// <param name="hKinmutaikei">
        ///     勤務体系コード </param>
        /// <returns>
        ///     受入れデータ文字列</returns>
        ///-----------------------------------------------------------
        private string getUkeireCsv(DataSet1.勤務票明細Row r, out string hKinmutaikei, sqlControl.DataControl sdCon)
        {
            string hDate = string.Empty;
            hKinmutaikei = string.Empty;
            StringBuilder sb = new StringBuilder();

            sb.Clear();

            // 社員番号
            sb.Append(r.社員番号).Append(",");

            // ヘッダテーブルから「日付」「勤務体系コード」を取得
            foreach (var t in _hTbl.Where(a => a.ID == r.ヘッダID))
            {
                // 日付
                hDate = t.年.ToString() + "/" + t.月.ToString().PadLeft(2, '0') + "/" + t.日.ToString().PadLeft(2, '0');

                // 勤務体系（シフト）コード
                hKinmutaikei = t.シフトコード.ToString();
            }

            // 変更シフトコードがあるとき
            if (r.シフトコード != string.Empty)
            {
                hKinmutaikei = r.シフトコード;
            }

            // 日付
            sb.Append(hDate).Append(",");

            // 勤務回
            sb.Append("1").Append(",");

            // 勤務体系コード
            sb.Append(hKinmutaikei).Append(",");

            // 事由
            sb.Append(r.事由1).Append(",");
            sb.Append(r.事由2).Append(",");
            sb.Append(r.事由3).Append(",");




            // 休暇事由が1日単位か調べる
            string[] mJiyu = { r.事由1, r.事由2, r.事由3 };
            OCR.clsJiyuDiv ss = new OCR.clsJiyuDiv(mJiyu);
            double jiyuDay = ss.jiyuDivCount(sdCon);

            // 休暇事由のときは出勤退勤時刻は不要
            //if (Utility.StrtoInt(r.事由1) == 1 || Utility.StrtoInt(r.事由2) == 1 || Utility.StrtoInt(r.事由3) == 1 ||
            //    Utility.StrtoInt(r.事由1) == 4 || Utility.StrtoInt(r.事由2) == 4 || Utility.StrtoInt(r.事由3) == 4 ||
            //    Utility.StrtoInt(r.事由1) == 7 || Utility.StrtoInt(r.事由2) == 7 || Utility.StrtoInt(r.事由3) == 7 ||
            //    Utility.StrtoInt(r.事由1) == 10 || Utility.StrtoInt(r.事由2) == 10 || Utility.StrtoInt(r.事由3) == 10 ||
            //    Utility.StrtoInt(r.事由1) == 12 || Utility.StrtoInt(r.事由2) == 12 || Utility.StrtoInt(r.事由3) == 12 ||
            //    Utility.StrtoInt(r.事由1) == 13 || Utility.StrtoInt(r.事由2) == 13 || Utility.StrtoInt(r.事由3) == 13 ||
            //    Utility.StrtoInt(r.事由1) == 40 || Utility.StrtoInt(r.事由2) == 40 || Utility.StrtoInt(r.事由3) == 40)

            // 事由が終日（半日＋半日含む）のときは出勤退勤時刻は不要
            if (jiyuDay == 1)
            {
                // 出勤時刻
                sb.Append("").Append(",");
                sb.Append("");
            }
            else
            {
                // 勤務体系（シフト）コードの開始終了時刻を取得する 2018/02/06
                string sftSt = string.Empty;
                string sftEt = string.Empty;
                DateTime cdt = DateTime.Today;  // 2018/02/06

                GetSftTime(hKinmutaikei.PadLeft(4, '0'), out sftSt, out sftEt, out cdt, sdCon);

                //出勤退出時刻
                if (r.出勤時 == string.Empty && r.出勤分 == string.Empty &&
                    r.退勤時 == string.Empty && r.退勤分 == string.Empty)
                {
                    // 勤務体系（シフト）コードの開始終了時刻をセットする
                    sb.Append(sftSt).Append(",").Append(sftEt);
                }
                else
                {
                    // 出退勤時刻が記入されているとき

                    // 出勤時刻
                    int sTime = int.Parse(r.出勤時) * 100 + int.Parse(r.出勤分);
                    int cTime = 0;

                    // 以下、休出以外 2018/02/06
                    if (Utility.StrtoInt(hKinmutaikei) != global.SFT_KYUSHUTSU &&
                        Utility.StrtoInt(hKinmutaikei) != global.SFT_KYUKEI_KYUSHUTSU)
                    {
                        // 日替わり時刻取得 // 2018/02/06
                        cTime = cdt.Hour * 100 + cdt.Minute;

                        // 日替わり時刻以前のときは翌日と判断する // 2018/02/06
                        if (sTime < cTime)
                        {
                            sb.Append("翌日");
                        }
                    }

                    sb.Append(r.出勤時 + ":" + r.出勤分.PadLeft(2, '0')).Append(",");

                    //退出時刻
                    int eTime = int.Parse(r.退勤時) * 100 + int.Parse(r.退勤分);

                    if (sTime >= eTime)
                    {
                        // 日付を跨いでいるとき
                        sb.Append("翌日");
                    }
                    else
                    {
                        // 以下、休出以外 2018/02/06
                        if (Utility.StrtoInt(hKinmutaikei) != global.SFT_KYUSHUTSU &&
                            Utility.StrtoInt(hKinmutaikei) != global.SFT_KYUKEI_KYUSHUTSU)
                        {
                            if (eTime < cTime)
                            {
                                // 終了時刻が日替わり時刻以前（開始時刻、終了時刻ともに翌日のとき）
                                sb.Append("翌日");
                            }
                        }
                    }

                    sb.Append(r.退勤時 + ":" + r.退勤分.PadLeft(2, '0'));
                }
            }

            return sb.ToString();
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     帰宅後勤務受入れデータ作成 </summary>
        /// <param name="r">
        ///     DataSet1.勤務票明細Row </param>
        /// <returns>
        ///     受入れデータ文字列</returns>
        ///----------------------------------------------------------------------
        private string getKitakugoCsv(DataSet1.勤務票明細Row r)
        {
            string c = string.Empty;
            StringBuilder sb = new StringBuilder();

            var k = _dts.帰宅後勤務.Where(a => a.年 == r.勤務票ヘッダRow.年 && a.月 == r.勤務票ヘッダRow.月 && a.日 == r.勤務票ヘッダRow.日 &&
                                           a.社員番号.PadLeft(6, '0') == r.社員番号.PadLeft(6, '0'));
            foreach (var s in k)
            {
                sb.Append(s.社員番号).Append(",");
                sb.Append(s.年.ToString() + "/" + s.月.ToString() + "/" + s.日.ToString()).Append(",");
                sb.Append("2").Append(",");
                sb.Append(s.シフトコード).Append(",");
                sb.Append(s.事由1).Append(",");
                sb.Append(s.事由2).Append(",");
                sb.Append(s.事由3).Append(",");

                string sMark = string.Empty;    // 2018/03/08

                // 出勤日 2018/03/08
                if (s.Is出勤日Null())
                {
                    sMark = string.Empty;
                }
                else
                {
                    sMark = s.出勤日;
                }

                // 出勤時刻 : 2018/03/08
                //sb.Append(s.出勤時 + ":" + s.出勤分.PadLeft(2, '0')).Append(",");
                sb.Append(sMark + s.出勤時 + ":" + s.出勤分.PadLeft(2, '0')).Append(",");

                //退出時刻
                // 2018/03/08 コメント化
                //int sTime = int.Parse(s.出勤時) * 100 + int.Parse(s.出勤分);
                //int eTime = int.Parse(s.退勤時) * 100 + int.Parse(s.退勤分);

                //if (sTime >= eTime)
                //{
                //    sb.Append("翌日");
                //}

                // 退勤日 2018/03/08
                if (s.Is退勤日Null())
                {
                    sMark = string.Empty;
                }
                else
                {
                    sMark = s.退勤日;
                }

                // 退勤時刻 : 2018/03/08
                //sb.Append(s.退勤時 + ":" + s.退勤分.PadLeft(2, '0'));
                sb.Append(sMark + s.退勤時 + ":" + s.退勤分.PadLeft(2, '0'));

                c = sb.ToString();

                break;
            }

            return c;
        }

        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     対象シフトコードの開始時刻と終了時刻を取得する 
        ///     : 日替わり時刻を取得 2018/02/06 </summary>
        /// <param name="_dbName">
        ///     データベース名</param>
        /// <param name="sftCode">
        ///     シフトコード </param>
        /// <param name="sTime">
        ///     開始時刻</param>
        /// <param name="eTime">
        ///     終了時刻</param>
        /// <param name="changeTime">
        ///     日替わり時刻</param>
        ///----------------------------------------------------------------------------------
        private void GetSftTime(string sftCode, out string sTime, out string eTime, out DateTime changeTime, sqlControl.DataControl sdCon)
        {
            // 対象のシフトコード取得する
            DateTime sDt = DateTime.Now;
            DateTime eDt = DateTime.Now;
            DateTime cDt = DateTime.Now;    // 2018/02/06

            // 勤務体系（シフト）コード情報取得
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("SELECT tbLaborSystem.LaborSystemID,tbLaborSystem.LaborSystemCode,");
            sb.Append("tbLaborSystem.LaborSystemName,tbLaborTimeSpanRule.StartTime,");
            sb.Append("tbLaborTimeSpanRule.EndTime,tbLaborSystem.DayChangeTime ");
            sb.Append("FROM tbLaborSystem inner join tbLaborTimeSpanRule ");
            sb.Append("on tbLaborSystem.LaborSystemID = tbLaborTimeSpanRule.LaborSystemID ");
            sb.Append("where tbLaborTimeSpanRule.LaborTimeSpanRuleType = 1 ");
            sb.Append("and tbLaborSystem.LaborSystemCode = '").Append(sftCode).Append("'");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                sDt = DateTime.Parse(dR["StartTime"].ToString());
                eDt = DateTime.Parse(dR["EndTime"].ToString());
                cDt = DateTime.Parse(dR["DayChangeTime"].ToString());   // 2018/02/06
                break;
            }

            dR.Close();

            // 開始時刻
            sTime = sDt.Hour.ToString() + ":" + sDt.Minute.ToString().PadLeft(2, '0');

            // 終了時刻
            eTime = string.Empty;
            if (sDt.Day < eDt.Day)
            {
                // 翌日のとき
                eTime = "翌日";
            }

            eTime += eDt.Hour.ToString() + ":" + eDt.Minute.ToString().PadLeft(2, '0');

            // 2018/02/06
            changeTime = cDt;
        }
        
        ///----------------------------------------------------------------------------
        /// <summary>
        ///     配列にテキストデータをセットする </summary>
        /// <param name="array">
        ///     社員、パート、出向社員の各配列</param>
        /// <param name="cnt">
        ///     拡張する配列サイズ</param>
        /// <param name="txtData">
        ///     セットする文字列</param>
        ///----------------------------------------------------------------------------
        private void txtArraySet(string [] array, int cnt, string txtData)
        {
            Array.Resize(ref array, cnt);   // 配列のサイズ拡張
            array[cnt - 1] = txtData;       // 文字列のセット
        }
        
        ///----------------------------------------------------------------------------
        /// <summary>
        ///     テキストファイルを出力する</summary>
        /// <param name="outFilePath">
        ///     出力するフォルダ</param>
        /// <param name="arrayData">
        ///     書き込む配列データ</param>
        ///----------------------------------------------------------------------------
        private void txtFileWrite(string sPath, string [] arrayData)
        {
            // 付加文字列（タイムスタンプ）
            string newFileName = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') +
                                    DateTime.Now.Day.ToString().PadLeft(2, '0') + DateTime.Now.Hour.ToString().PadLeft(2, '0') +
                                    DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0'); 

            // ファイル名
            string outFileName = sPath + TXTFILENAME + newFileName + ".csv";
            
            // テキストファイル出力
            File.WriteAllLines(outFileName, arrayData, System.Text.Encoding.GetEncoding(932));
        }
    }
}
