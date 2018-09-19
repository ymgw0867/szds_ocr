using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;

namespace SZDS_TIMECARD.Common
{
    class OCROuen
    {
        public OCROuen(string dbName, xlsData exlms)
        {
            _dbName = dbName;
            bs = exlms;
        }

        // 奉行シリーズデータ領域データベース名
        string _dbName = string.Empty;

        Common.xlsData bs;

        #region エラー項目番号プロパティ
        //---------------------------------------------------
        //          エラー情報
        //---------------------------------------------------
        /// <summary>
        ///     エラーヘッダ行RowIndex</summary>
        public int _errHeaderIndex { get; set; }

        /// <summary>
        ///     エラー項目番号</summary>
        public int _errNumber { get; set; }

        /// <summary>
        ///     エラー明細行RowIndex </summary>
        public int _errRow { get; set; }

        /// <summary> 
        ///     エラーメッセージ </summary>
        public string _errMsg { get; set; }

        /// <summary> 
        ///     エラーなし </summary>
        public int eNothing = 0;

        /// <summary> 
        ///     エラー項目 = 対象年月日 </summary>
        public int eYearMonth = 1;

        /// <summary> 
        ///     エラー項目 = 対象月 </summary>
        public int eMonth = 2;

        /// <summary> 
        ///     エラー項目 = 日 </summary>
        public int eDay = 3;

        /// <summary> 
        ///     エラー項目 = 勤務体系コード </summary>
        public int eKinmuTaikeiCode = 4;

        /// <summary> 
        ///     エラー項目 = 個人番号 </summary>
        public int eShainNo = 5;

        /// <summary> 
        ///     エラー項目 = 勤務記号 </summary>
        public int eKintaiKigou = 6;

        /// <summary> 
        ///     エラー項目 = 部署コード </summary>
        public int eBushoCode = 7;

        /// <summary> 
        ///     エラー項目 = 応援チェック </summary>
        public int eChkOuen = 7;

        /// <summary> 
        ///     エラー項目 = シフト変更 </summary>
        public int eChksft = 8;

        /// <summary> 
        ///     エラー項目 = 事由 </summary>
        public int eJiyu1 = 9;
        public int eJiyu2 = 10;
        public int eJiyu3 = 11;

        /// <summary> 
        ///     エラー項目 = シフトコード </summary>
        public int eSftCode = 12;
        
        /// <summary> 
        ///     エラー項目 = 開始時 </summary>
        public int eSH = 13;

        /// <summary> 
        ///     エラー項目 = 開始分 </summary>
        public int eSM = 14;

        /// <summary> 
        ///     エラー項目 = 終了時 </summary>
        public int eEH = 15;

        /// <summary> 
        ///     エラー項目 = 終了分 </summary>
        public int eEM = 16;

        /// <summary> 
        ///     エラー項目 = 残業理由1 </summary>
        public int eZanRe1 = 17;

        /// <summary> 
        ///     エラー項目 = 残業時1 </summary>
        public int eZanH1 = 18;

        /// <summary> 
        ///     エラー項目 = 残業分1 </summary>
        public int eZanM1 = 19;

        /// <summary> 
        ///     エラー項目 = 残業理由2 </summary>
        public int eZanRe2 = 20;

        /// <summary> 
        ///     エラー項目 = 残業時2 </summary>
        public int eZanH2 = 21;

        /// <summary> 
        ///     エラー項目 = 残業分2 </summary>
        public int eZanM2 = 22;
        #endregion
        
        #region 警告項目
        ///     <!--警告項目配列 -->
        public int[] warArray = new int[6];

        /// <summary>
        ///     警告項目番号</summary>
        public int _warNumber { get; set; }

        /// <summary>
        ///     警告明細行RowIndex </summary>
        public int _warRow { get; set; }

        /// <summary> 
        ///     警告項目 = 勤怠記号1&2 </summary>
        public int wKintaiKigou = 0;

        /// <summary> 
        ///     警告項目 = 開始終了時分 </summary>
        public int wSEHM = 1;

        /// <summary> 
        ///     警告項目 = 時間外時分 </summary>
        public int wZHM = 2;

        /// <summary> 
        ///     警告項目 = 深夜勤務時分 </summary>
        public int wSIHM = 3;

        /// <summary> 
        ///     警告項目 = 休日出勤時分 </summary>
        public int wKSHM = 4;

        /// <summary> 
        ///     警告項目 = 出勤形態 </summary>
        public int wShukeitai = 5;

        #endregion

        #region フィールド定義
        /// <summary> 
        ///     警告項目 = 時間外1.25時 </summary>
        public int [] wZ125HM = new int[global.MAX_GYO];

        /// <summary> 
        ///     実働時間 </summary>
        public double _workTime;

        /// <summary> 
        ///     深夜稼働時間 </summary>
        public double _workShinyaTime;
        #endregion

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

        /// <summary> 
        ///     残業分単位０ </summary>
        private string zanMinTANI0 = "0";

        /// <summary> 
        ///     残業分単位５ </summary>
        private string zanMinTANI5 = "5";

        #endregion

        #region 時間チェック記号定数
        private const string cHOUR = "H";           // 時間をチェック
        private const string cMINUTE = "M";         // 分をチェック
        private const string cTIME = "HM";          // 時間・分をチェック
        #endregion

        private const string WKSPAN0750 = "7時間50分";
        private const string WKSPAN0755 = "7時間55分";
        private const string WKSPAN0800 = "8時間";
        private const string WKSPAN_KYUJITSU = "休日出勤";

        // 休憩時間
        private const Int64 RESTTIME0750 = 60;      // 7時間50分
        private const Int64 RESTTIME0755 = 65;      // 7時間55分
        private const Int64 RESTTIME0800 = 60;      // 8時間

        // テーブルアダプターマネージャーインスタンス
        DataSet1TableAdapters.TableAdapterManager adpMn = new DataSet1TableAdapters.TableAdapterManager();
        DataSet1TableAdapters.休日TableAdapter kAdp = new DataSet1TableAdapters.休日TableAdapter();
        
        ///-----------------------------------------------------------------------
        /// <summary>
        ///     CSVデータをMDBに登録する：DataSet Version </summary>
        /// <param name="_InPath">
        ///     CSVデータパス</param>
        /// <param name="frmP">
        ///     プログレスバーフォームオブジェクト</param>
        /// <param name="dts">
        ///     データセット</param>
        /// <param name="dbName">
        ///     データ領域データベース名</param>
        ///-----------------------------------------------------------------------
        public void CsvToMdb(string _inPath, frmPrg frmP, string dbName)
        {
            string headerKey = string.Empty;    // ヘッダキー

            // テーブルセットオブジェクト
            DataSet1 dts = new SZDS_TIMECARD.DataSet1();

            try
            {
                // 勤務表ヘッダデータセット読み込み
                DataSet1TableAdapters.勤務票ヘッダTableAdapter hAdp = new DataSet1TableAdapters.勤務票ヘッダTableAdapter();
                adpMn.勤務票ヘッダTableAdapter = hAdp;
                adpMn.勤務票ヘッダTableAdapter.Fill(dts.勤務票ヘッダ);

                // 勤務表明細データセット読み込み
                DataSet1TableAdapters.勤務票明細TableAdapter iAdp = new DataSet1TableAdapters.勤務票明細TableAdapter();
                adpMn.勤務票明細TableAdapter = iAdp;
                adpMn.勤務票明細TableAdapter.Fill(dts.勤務票明細);

                // 対象CSVファイル数を取得
                string [] t = System.IO.Directory.GetFiles(_inPath, "*.csv");
                int cLen = t.Length;

                //CSVデータをMDBへ取込
                int cCnt = 0;
                foreach (string files in System.IO.Directory.GetFiles(_inPath, "*.csv"))
                {
                    //件数カウント
                    cCnt++;

                    //プログレスバー表示
                    frmP.Text = "OCR変換CSVデータロード中　" + cCnt.ToString() + "/" + cLen.ToString();
                    frmP.progressValue = cCnt * 100 / cLen;
                    frmP.ProgressStep();

                    ////////OCR処理対象のCSVファイルかファイル名の文字数を検証する
                    //////string fn = Path.GetFileName(files);

                    // CSVファイルインポート
                    var s = System.IO.File.ReadAllLines(files, Encoding.Default);
                    foreach (var stBuffer in s)
                    {
                        // カンマ区切りで分割して配列に格納する
                        string[] stCSV = stBuffer.Split(',');

                        // ヘッダ行
                        if (stCSV[0] == "*")
                        {
                            // ヘッダーキー取得
                            headerKey = Utility.GetStringSubMax(stCSV[1].Trim(), 17);

                            // データセットに勤務票ヘッダデータを追加する
                            dts.勤務票ヘッダ.Add勤務票ヘッダRow(setNewHeadRecRow(dts, stCSV, dbName));
                        }
                        else　// 明細行
                        {
                            // データセットに勤務表明細データを追加する
                            dts.勤務票明細.Add勤務票明細Row(setNewItemRecRow(dts, headerKey, stCSV));
                        }
                    }
                }

                // ローカルのデータベースを更新
                adpMn.UpdateAll(dts);

                //CSVファイルを削除する
                foreach (string files in System.IO.Directory.GetFiles(_inPath, "*.csv"))
                {
                    System.IO.File.Delete(files);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "勤務票CSVインポート処理", MessageBoxButtons.OK);
            }
            finally
            {
            }
        }

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     応援移動票CSVデータをMDBに登録する：DataSet Version </summary>
        /// <param name="_InPath">
        ///     CSVデータパス</param>
        /// <param name="frmP">
        ///     プログレスバーフォームオブジェクト</param>
        /// <param name="dts">
        ///     データセット</param>
        /// <param name="dbName">
        ///     データ領域データベース名</param>
        ///-----------------------------------------------------------------------
        public void csvToMdbOuen(string _inPath, frmPrg frmP, string dbName)
        {
            string headerKey = string.Empty;    // ヘッダキー

            // テーブルセットオブジェクト
            DataSet1 dts = new SZDS_TIMECARD.DataSet1();

            try
            {
                // 応援移動票ヘッダデータセット読み込み
                DataSet1TableAdapters.応援移動票ヘッダTableAdapter hAdp = new DataSet1TableAdapters.応援移動票ヘッダTableAdapter();
                adpMn.応援移動票ヘッダTableAdapter = hAdp;
                adpMn.応援移動票ヘッダTableAdapter.Fill(dts.応援移動票ヘッダ);

                // 勤務表明細データセット読み込み
                DataSet1TableAdapters.応援移動票明細TableAdapter iAdp = new DataSet1TableAdapters.応援移動票明細TableAdapter();
                adpMn.応援移動票明細TableAdapter = iAdp;
                adpMn.応援移動票明細TableAdapter.Fill(dts.応援移動票明細);

                // 対象CSVファイル数を取得
                string[] t = System.IO.Directory.GetFiles(_inPath, "*.csv");
                int cLen = t.Length;

                //CSVデータをMDBへ取込
                int cCnt = 0;
                foreach (string files in System.IO.Directory.GetFiles(_inPath, "*.csv"))
                {
                    //件数カウント
                    cCnt++;

                    //プログレスバー表示
                    frmP.Text = "OCR変換CSVデータロード中　" + cCnt.ToString() + "/" + cLen.ToString();
                    frmP.progressValue = cCnt * 100 / cLen;
                    frmP.ProgressStep();

                    ////////OCR処理対象のCSVファイルかファイル名の文字数を検証する
                    //////string fn = Path.GetFileName(files);

                    // CSVファイルインポート
                    var s = System.IO.File.ReadAllLines(files, Encoding.Default);
                    foreach (var stBuffer in s)
                    {
                        // カンマ区切りで分割して配列に格納する
                        string[] stCSV = stBuffer.Split(',');

                        // ヘッダ行
                        if (stCSV[0] == "*")
                        {
                            // ヘッダーキー取得
                            headerKey = Utility.GetStringSubMax(stCSV[1].Trim(), 17);

                            // データセットに勤務票ヘッダデータを追加する
                            dts.応援移動票ヘッダ.Add応援移動票ヘッダRow(setNewHeadOuen(dts, stCSV, dbName));
                        }
                        else　// 明細行
                        {
                            // データセットに勤務表明細データを追加する
                            dts.応援移動票明細.Add応援移動票明細Row(setNewItemOuen(dts, headerKey, stCSV));
                        }
                    }
                }

                // ローカルのデータベースを更新
                adpMn.UpdateAll(dts);

                //CSVファイルを削除する
                foreach (string files in System.IO.Directory.GetFiles(_inPath, "*.csv"))
                {
                    System.IO.File.Delete(files);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "応援移動票CSVインポート処理", MessageBoxButtons.OK);
            }
            finally
            {
            }
        }

        ///---------------------------------------------------------------------------------
        /// <summary>
        ///     追加用勤務票ヘッダRowオブジェクトを作成する </summary>
        /// <param name="tblSt">
        ///     テーブルセット</param>
        /// <param name="stCSV">
        ///     CSV配列</param>
        /// <param name="dbName">
        ///     データ領域データベース名</param>
        /// <returns>
        ///     追加する勤務票ヘッダRowオブジェクト</returns>
        ///---------------------------------------------------------------------------------
        private DataSet1.勤務票ヘッダRow setNewHeadRecRow(DataSet1 tblSt, string[] stCSV, string dbName)
        {
            DataSet1.勤務票ヘッダRow r = tblSt.勤務票ヘッダ.New勤務票ヘッダRow();
            r.ID = Utility.GetStringSubMax(stCSV[1].Trim(), 17);
            r.年 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[3].Trim().Replace("-", ""), 2));
            r.月 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[4].Trim().Replace("-", ""), 2));
            r.日 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[5].Trim().Replace("-", ""), 2));
            r.部署コード = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[6].Trim().Replace("-", ""), 5));
            r.シフトコード = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[7].Trim().Replace("-", ""), 3));
            r.画像名 = Utility.GetStringSubMax(stCSV[1].Trim(), 17) + ".tif";
            r.データ領域名 = dbName;
            r.確認 = global.flgOff;
            r.更新年月日 = DateTime.Now;

            return r;
        }

        ///---------------------------------------------------------------------------------
        /// <summary>
        ///     追加用応援移動票ヘッダRowオブジェクトを作成する </summary>
        /// <param name="tblSt">
        ///     テーブルセット</param>
        /// <param name="stCSV">
        ///     CSV配列</param>
        /// <param name="dbName">
        ///     データ領域データベース名</param>
        /// <returns>
        ///     追加する応援移動票ヘッダRowオブジェクト</returns>
        ///---------------------------------------------------------------------------------
        private DataSet1.応援移動票ヘッダRow setNewHeadOuen(DataSet1 tblSt, string[] stCSV, string dbName)
        {
            DataSet1.応援移動票ヘッダRow r = tblSt.応援移動票ヘッダ.New応援移動票ヘッダRow();
            r.ID = Utility.GetStringSubMax(stCSV[1].Trim(), 17);
            r.年 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[3].Trim().Replace("-", ""), 4));
            r.月 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[4].Trim().Replace("-", ""), 2));
            r.日 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[5].Trim().Replace("-", ""), 2));
            r.部署コード = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[6].Trim().Replace("-", ""), 5));
            r.画像名 = Utility.GetStringSubMax(stCSV[1].Trim(), 17) + ".tif";
            r.データ領域名 = dbName;
            r.確認 = global.flgOff;
            r.更新年月日 = DateTime.Now;
            return r;
        }

        ///---------------------------------------------------------------------------------
        /// <summary>
        ///     追加用勤務票明細Rowオブジェクトを作成する </summary>
        /// <param name="tblSt">
        ///     テーブルセットオブジェクト</param>
        /// <param name="headerKey">
        ///     ヘッダキー</param>
        /// <param name="stCSV">
        ///     CSV配列</param>
        /// <param name="sftCode">
        ///     ヘッダシフトコード</param>
        /// <returns>
        ///     追加する勤務票明細Rowオブジェクト</returns>
        ///---------------------------------------------------------------------------------
        private DataSet1.勤務票明細Row setNewItemRecRow(DataSet1 tblSt, string headerKey, string[] stCSV)
        {
            DataSet1.勤務票明細Row r = tblSt.勤務票明細.New勤務票明細Row();

            r.ヘッダID = headerKey;
            r.応援 = stCSV[0].Trim();
            r.シフト通り = stCSV[1].Trim();
            r.社員番号 = Utility.GetStringSubMax(stCSV[2].Trim().Replace("-", ""), 6);
            r.事由1 = Utility.GetStringSubMax(stCSV[3].Trim().Replace("-", ""), 2);
            r.事由2 = Utility.GetStringSubMax(stCSV[4].Trim().Replace("-", ""), 2);
            r.事由3 = Utility.GetStringSubMax(stCSV[5].Trim().Replace("-", ""), 2);

            if (r.シフト通り == global.FLGON)
            {
                // シフト通りのとき
                r.シフトコード = string.Empty;
            }
            else
            {
                // シフト通りでないとき記入されたシフトコードを適用
                r.シフトコード = Utility.GetStringSubMax(stCSV[6].Trim().Replace("-", ""), 3);
            }

            r.出勤時 = Utility.GetStringSubMax(stCSV[7].Trim().Replace("-", ""), 2);
            r.出勤分 = Utility.GetStringSubMax(stCSV[8].Trim().Replace("-", ""), 2);
            r.退勤時= Utility.GetStringSubMax(stCSV[9].Trim().Replace("-", ""), 2);
            r.退勤分 = Utility.GetStringSubMax(stCSV[10].Trim().Replace("-", ""), 2);
            r.残業理由1 = Utility.GetStringSubMax(stCSV[11].Trim().Replace("-", ""), 2);
            r.残業時1 = Utility.GetStringSubMax(stCSV[12].Trim().Replace("-", ""), 2);
            r.残業分1 = Utility.GetStringSubMax(stCSV[13].Trim().Replace("-", ""), 1);
            r.残業理由2 = Utility.GetStringSubMax(stCSV[14].Trim().Replace("-", ""), 2);
            r.残業時2 = Utility.GetStringSubMax(stCSV[15].Trim().Replace("-", ""), 2);
            r.残業分2 = Utility.GetStringSubMax(stCSV[16].Trim().Replace("-", ""), 1);
            r.取消 = global.FLGOFF;
            r.データ領域名 = string.Empty;
            r.更新年月日 = DateTime.Now;
            
            return r;
        }

        ///---------------------------------------------------------------------------------
        /// <summary>
        ///     追加用勤務票明細Rowオブジェクトを作成する </summary>
        /// <param name="tblSt">
        ///     テーブルセットオブジェクト</param>
        /// <param name="headerKey">
        ///     ヘッダキー</param>
        /// <param name="stCSV">
        ///     CSV配列</param>
        /// <param name="sftCode">
        ///     ヘッダシフトコード</param>
        /// <returns>
        ///     追加する勤務票明細Rowオブジェクト</returns>
        ///---------------------------------------------------------------------------------
        private DataSet1.応援移動票明細Row setNewItemOuen(DataSet1 tblSt, string headerKey, string[] stCSV)
        {
            DataSet1.応援移動票明細Row r = tblSt.応援移動票明細.New応援移動票明細Row();

            r.ヘッダID = headerKey;
            r.データ区分 = Utility.StrtoInt(stCSV[1]);
            r.社員番号 = Utility.GetStringSubMax(stCSV[2].Trim().Replace("-", ""), 6);
            r.ライン = Utility.GetStringSubMax(stCSV[3].Trim().Replace("-", ""), 3);
            r.部門 = Utility.GetStringSubMax(stCSV[4].Trim().Replace("-", ""), 3);
            r.製品群 = Utility.GetStringSubMax(stCSV[5].Trim().Replace("-", ""), 3);

            if (Utility.StrtoInt(stCSV[1]) == 1)
            {
                // 日中応援
                r.応援時 = Utility.GetStringSubMax(stCSV[6].Trim().Replace("-", ""), 2);
                r.応援分 = Utility.GetStringSubMax(stCSV[7].Trim().Replace("-", ""), 1);
                r.残業理由1 = string.Empty;
                r.残業時1 = string.Empty;
                r.残業分1 = string.Empty; 
                r.残業理由2 = string.Empty;
                r.残業時2 = string.Empty;
                r.残業分2 = string.Empty;
            }
            else if (Utility.StrtoInt(stCSV[1]) == 2)
            {
                // 残業応援
                r.応援時 = string.Empty;
                r.応援分 = string.Empty;
                r.残業理由1 = Utility.GetStringSubMax(stCSV[6].Trim().Replace("-", ""), 2);
                r.残業時1 = Utility.GetStringSubMax(stCSV[7].Trim().Replace("-", ""), 2);
                r.残業分1 = Utility.GetStringSubMax(stCSV[8].Trim().Replace("-", ""), 1);
                r.残業理由2 = Utility.GetStringSubMax(stCSV[9].Trim().Replace("-", ""), 2);
                r.残業時2 = Utility.GetStringSubMax(stCSV[10].Trim().Replace("-", ""), 2);
                r.残業分2 = Utility.GetStringSubMax(stCSV[11].Trim().Replace("-", ""), 1);
            }
            
            r.取消 = string.Empty;
            r.データ領域名 = string.Empty;
            r.更新年月日 = DateTime.Now;

            return r;
        }

        ///----------------------------------------------------------------------------------------
        /// <summary>
        ///     値1がemptyで値2がNot string.Empty のとき "0"を返す。そうではないとき値1をそのまま返す</summary>
        /// <param name="str1">
        ///     値1：文字列</param>
        /// <param name="str2">
        ///     値2：文字列</param>
        /// <returns>
        ///     文字列</returns>
        ///----------------------------------------------------------------------------------------
        private string hmStrToZero(string str1, string str2)
        {
            string rVal = str1;
            if (str1 == string.Empty && str2 != string.Empty)
                rVal = "0";

            return rVal;
        }

        ///--------------------------------------------------------------------------------------------------
        /// <summary>
        ///     エラーチェックメイン処理。
        ///     エラーのときOCRDataクラスのヘッダ行インデックス、フィールド番号、明細行インデックス、
        ///     エラーメッセージが記録される </summary>
        /// <param name="sIx">
        ///     開始ヘッダ行インデックス</param>
        /// <param name="eIx">
        ///     終了ヘッダ行インデックス</param>
        /// <param name="frm">
        ///     親フォーム</param>
        /// <param name="dts">
        ///     データセット</param>
        /// <returns>
        ///     True:エラーなし、false:エラーあり</returns>
        ///-----------------------------------------------------------------------------------------------
        public Boolean errCheckMain(int sIx, int eIx, Form frm, DataSet1 dts)
        {
            int rCnt = 0;

            // オーナーフォームを無効にする
            frm.Enabled = false;

            // プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = frm;
            frmP.Show();

            // レコード件数取得
            int cTotal = dts.勤務票ヘッダ.Rows.Count;

            // 出勤簿データ読み出し
            Boolean eCheck = true;

            for (int i = 0; i < cTotal; i++)
            {
                //データ件数加算
                rCnt++;

                //プログレスバー表示
                frmP.Text = "エラーチェック実行中　" + rCnt.ToString() + "/" + cTotal.ToString();
                frmP.progressValue = rCnt * 100 / cTotal;
                frmP.ProgressStep();

                //指定範囲ならエラーチェックを実施する：（i:行index）
                if (i >= sIx && i <= eIx)
                {
                    // 勤務票ヘッダ行のコレクションを取得します
                    DataSet1.勤務票ヘッダRow r = (DataSet1.勤務票ヘッダRow)dts.勤務票ヘッダ.Rows[i];

                    // エラーチェック実施
                    eCheck = errCheckData(dts, r);

                    if (!eCheck)　//エラーがあったとき
                    {
                        _errHeaderIndex = i;     // エラーとなったヘッダRowIndex
                        break;
                    }
                }
            }

            // いったんオーナーをアクティブにする
            frm.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            frm.Enabled = true;

            return eCheck;
        }

        ///---------------------------------------------------------------------------------
        /// <summary>
        ///     エラー情報を取得します </summary>
        /// <param name="eID">
        ///     エラーデータのID</param>
        /// <param name="eNo">
        ///     エラー項目番号</param>
        /// <param name="eRow">
        ///     エラー明細行</param>
        /// <param name="eMsg">
        ///     表示メッセージ</param>
        ///---------------------------------------------------------------------------------
        private void setErrStatus(int eNo, int eRow, string eMsg)
        {
            //errHeaderIndex = eHRow;
            _errNumber = eNo;
            _errRow = eRow;
            _errMsg = eMsg;
        }

        ///-----------------------------------------------------------------------------------------------
        /// <summary>
        ///     項目別エラーチェック。
        ///     エラーのときヘッダ行インデックス、フィールド番号、明細行インデックス、エラーメッセージが記録される </summary>
        /// <param name="dts">
        ///     データセット</param>
        /// <param name="r">
        ///     勤務票ヘッダ行コレクション</param>
        /// <returns>
        ///     エラーなし：true, エラー有り：false</returns>
        ///-----------------------------------------------------------------------------------------------
        /// 
        public Boolean errCheckData(DataSet1 dts, DataSet1.勤務票ヘッダRow r)
        {
            string sDate;
            DateTime eDate;

            // 対象年月日
            sDate = r.年.ToString() + "/" + r.月.ToString() + "/" + r.日.ToString();
            if (!DateTime.TryParse(sDate, out eDate))
            {
                setErrStatus(eYearMonth, 0, "年月日が正しくありません");
                return false;
            }
            
            // 該当日が休日か調べる
            string sHol = global.FLGOFF;
            if (dts.休日.Any(a => a.年月日 == eDate))
            {
                sHol = global.FLGON;
            }

            // 勤務体系コード就業奉行登録チェック
            if (!chkSftCode(r.シフトコード.ToString()))
            {
                setErrStatus(eKinmuTaikeiCode, 0, "就業奉行に登録されていないシフトコードです");
                return false;
            }

            // 部署別勤務体系Excelシート登録チェック
            string sftName = string.Empty;
            if (!bs.getBushoSft(out sftName, r.部署コード.ToString(), r.シフトコード.ToString(), sHol))
            {
                setErrStatus(eKinmuTaikeiCode, 0, "該当部署に登録されていないか勤務日・休日に該当しないシフトコードです");
                return false;
            }

            //
            // 社員別勤怠記入欄データ
            //

            int iX = 0;
            string k = string.Empty;    // 特別休暇記号
            string yk = string.Empty;   // 有給記号
                        
            // 勤務票明細データ行を取得 2015/04/15
            List<DataSet1.勤務票明細Row> mList = dts.勤務票明細.Where(a => a.ヘッダID == r.ID).OrderBy(a => a.ID).ToList();

            foreach (var m in mList)
            {
                // 行数
                iX++;

                // 無記入の行はチェック対象外とする
                if (m.社員番号 == string.Empty && m.出勤時 == string.Empty &&
                    m.出勤分 == string.Empty && m.退勤時 == string.Empty &&
                    m.退勤分 == string.Empty && m.残業理由1 == string.Empty &&
                    m.残業時1 == string.Empty && m.残業分1 == string.Empty &&
                    m.残業理由2 == string.Empty && m.残業時2 == string.Empty &&
                    m.残業分2 == string.Empty && m.事由1 == string.Empty &&
                    m.事由2 == string.Empty && m.事由3 == string.Empty &&
                    m.応援 == global.FLGOFF && m.シフト通り == global.FLGOFF &&
                    m.シフトコード == string.Empty)
                {
                    continue;
                }

                // 取消行はチェック対象外とする   // 2015/03/10
                if (m.取消 == global.FLGON)
                {
                    continue;
                }

                // 社員番号：数字以外のとき
                if (!Utility.NumericCheck(Utility.NulltoStr(m.社員番号)))
                {
                    setErrStatus(eShainNo, iX - 1, "社員番号が入力されていません");
                    return false;
                }
                
                // 登録済み社員番号マスター検証
                if (!chkShainCode(m.社員番号))
                {
                    setErrStatus(eShainNo, iX - 1, "マスター未登録の社員番号です");
                    return false;
                }

                // 同一作業日報内で同じ社員番号が複数記入されているとエラー 2015/04/15
                if (!getSameNumber(mList, m.社員番号))
                {
                    setErrStatus(eShainNo, iX - 1, "同じ社員番号のデータが複数あります");
                    return false;
                }

                // 明細記入チェック
                if (!errCheckRow(m, "勤怠データＩ／Ｐ票内容", iX)) return false;

                // 応援チェック
                // 応援チェックありで応援依頼票がないとき

                // 応援チェックなしで応援依頼票があるとき


                // シフト通りチェック
                if (m.シフト通り == global.FLGON && m.シフトコード != string.Empty)
                {
                    setErrStatus(eChksft, iX - 1, "変更シフトコードが記入されています");
                    return false;
                }

                // 事由コード
                if (m.事由1 != string.Empty)
                {
                    if (!chkJiyu(m.事由1))
                    {
                        setErrStatus(eJiyu1, iX - 1, "マスター未登録の事由です");
                        return false;
                    }
                }

                if (m.事由2 != string.Empty)
                {
                    if (!chkJiyu(m.事由2))
                    {
                        setErrStatus(eJiyu2, iX - 1, "マスター未登録の事由です");
                        return false;
                    }
                }

                if (m.事由3 != string.Empty)
                {
                    if (!chkJiyu(m.事由3))
                    {
                        setErrStatus(eJiyu3, iX - 1, "マスター未登録の事由です");
                        return false;
                    }
                }

                // 変更シフトコード
                if (m.シフト通り == string.Empty && m.シフトコード == string.Empty)
                {
                    setErrStatus(eSftCode, iX - 1, "変更シフトコードが未記入です");
                    return false;
                }

                // 勤務体系（シフト）コード奉行登録チェック
                if (m.シフトコード != string.Empty)
                {
                    if (!chkSftCode(m.シフトコード))
                    {
                        setErrStatus(eSftCode, iX - 1, "就業奉行に登録されていないシフトコードです");
                        return false;
                    }

                    //// 部署別勤務体系Excelシート登録チェック
                    //if (!bs.getBushoSft(out sftName, r.部署コード.ToString(), m.シフトコード.ToString(), sHol))
                    //{
                    //    setErrStatus(eSftCode, iX - 1, "該当部署に登録されていないか勤務日・休日に該当しないシフトコードです");
                    //    return false;
                    //}
                }

                // 始業時刻・終業時刻チェック
                if (!errCheckTime(m, "出退時間", tanMin1, iX)) return false;

                // 残業理由
                if (!chkZangyoRe(m.残業理由1, m.残業時1, m.残業分1))
                {
                    setErrStatus(eZanRe1, iX - 1, "残業理由が未記入です");
                    return false;
                }

                if (!chkZangyoRe2(m.残業理由1, m.残業時1, m.残業分1))
                {
                    setErrStatus(eZanH1, iX - 1, "残業時間が未記入です");
                    return false;
                }

                // 部署別残業理由Excelシート登録チェック
                string reName = string.Empty;
                if (m.残業理由1 != string.Empty)
                {
                    if (!bs.getBushoZanRe(out reName, r.部署コード.ToString(), m.残業理由1.ToString()))
                    {
                        setErrStatus(eZanRe1, iX - 1, "該当部署に登録されていない残業理由です");
                        return false;
                    }
                }

                if (!chkZangyoRe(m.残業理由2, m.残業時2, m.残業分2))
                {
                    setErrStatus(eZanRe2, iX - 1, "残業理由が未記入です");
                    return false;
                }

                if (!chkZangyoRe2(m.残業理由2, m.残業時2, m.残業分2))
                {
                    setErrStatus(eZanH2, iX - 1, "残業時間が未記入です");
                    return false;
                }

                // 部署別残業理由Excelシート登録チェック
                if (m.残業理由2 != string.Empty)
                {
                    if (!bs.getBushoZanRe(out reName, r.部署コード.ToString(), m.残業理由2.ToString()))
                    {
                        setErrStatus(eZanRe2, iX - 1, "該当部署に登録されていない残業理由です");
                        return false;
                    }
                }

                // 残業分単位
                if (!chkZangyoMin(m.残業分1))
                {
                    setErrStatus(eZanM1, iX - 1, "残業分単位は０または５です");
                    return false;
                }

                if (!chkZangyoMin(m.残業分2))
                {
                    setErrStatus(eZanM2, iX - 1, "残業分単位は０または５です");
                    return false;
                }


                // 時間外チェック
                if (!errCheckZan(m, "時間外", tanMin30, iX)) return false;

                // エラーチェックから時間外、深夜時間計算チェックを撤廃 2015/10/01

                //// 所定勤務時間が取得されているとき残業時間計算チェックを行う
                //Int64 s10 = 0;  // 深夜勤務時間中の10分休憩時間

                //if (wkSpan != 0)
                //{
                //    Int64 restTm = 0;

                //    // 所定時間ごとの休憩時間
                //    if (wkSpanName == WKSPAN0750)
                //    {
                //        restTm = RESTTIME0750;
                //    }
                //    else if (wkSpanName == WKSPAN0755)
                //    {
                //        restTm = RESTTIME0755;
                //    }
                //    else if (wkSpanName == WKSPAN0800)
                //    {
                //        restTm = RESTTIME0800;
                //    }

                //    // 時間外取得 2015/09/15
                //    Int64 zan = getZangyoTime(m, (Int64)tanMin30, wkSpan, restTm, out s10, r.勤務体系コード);

                //    // 時間外記入時間チェック 2015/09/15
                //    if (!errCheckZanTm(m, "時間外", iX, zan)) return false;
                //}

                // 深夜勤務チェック
                if (!errCheckShinya(m, "深夜残業", tanMin10, iX)) return false;

                // 深夜チェックしない
                //// 所定労働時間が取得された勤務体系のとき深夜勤務時間をチェックする：休日出勤、契約社員は対象外
                //if (wkSpan != 0)
                //{
                //    // 深夜勤務時間を取得
                //    double shinyaTm = getShinyaWorkTime(m.開始時, m.開始分, m.終了時, m.終了分, tanMin10, s10);

                //    // 深夜勤務時間チェック
                //    if (!errCheckShinyaTm(m, "深夜残業", iX, (Int64)shinyaTm)) return false;
                //}
            }

            return true;
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     残業理由 </summary>
        /// <param name="zanRe">
        ///     残業理由</param>
        /// <param name="zH">
        ///     残業時</param>
        /// <param name="zM">
        ///     残業分</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------------
        private bool chkZangyoRe(string zanRe, string zH, string zM)
        {
            bool rtn = true;

            int z = Utility.StrtoInt(zH) + Utility.StrtoInt(zM);

            // 残業時間に有効数値が記入されているとき
            if (z > 0)
            {
                // 残業理由が無記入のとき
                if ((zH != string.Empty || zM != string.Empty) && zanRe == string.Empty)
                {
                    rtn = false;
                }
            }

            return rtn;
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     残業理由 </summary>
        /// <param name="zanRe">
        ///     残業理由</param>
        /// <param name="zH">
        ///     残業時</param>
        /// <param name="zM">
        ///     残業分</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------------
        private bool chkZangyoRe2(string zanRe, string zH, string zM)
        {
            bool rtn = true;

            int z = Utility.StrtoInt(zH) + Utility.StrtoInt(zM);

            // 残業理由の記入があって残業が無記入のとき
            if (zanRe != string.Empty)
            {
                // 残業時間に有効数値が未記入のとき
                if (z == 0)
                {
                    rtn = false;
                }
            }

            return rtn;
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     残業分単位 </summary>
        /// <param name="zM">
        ///     残業分</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------------
        private bool chkZangyoMin(string zM)
        {
            bool rtn = true;

            if (zM != string.Empty)
            {
                if (zM != zanMinTANI0 && zM != zanMinTANI5)
                {
                    rtn = false;
                }
            }

            return rtn;
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     社員コードチェック </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl オブジェクト </param>
        /// <param name="j">
        ///     社員コード</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        private bool chkShainCode(string s)
        {
            bool dm = false;

            // 奉行SQLServer接続文字列取得
            string sc = sqlControl.obcConnectSting.get(_dbName);
            sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            // 登録済み事由コード検証
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select EmployeeNo,RetireCorpDate from tbEmployeeBase ");
            sb.Append("where EmployeeNo = '" + s.PadLeft(10, '0') + "'");
            sb.Append(" and BeOnTheRegisterDivisionID != 9");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                dm = true;
                break;
            }

            dR.Close();

            return dm;
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     シフトコードチェック </summary>
        /// <param name="sdCon">
        ///     qlControl.DataControl オブジェクト </param>
        /// <param name="j">
        ///     シフトコード</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        private bool chkSftCode(string s)
        {
            bool dm = false;

            // 奉行SQLServer接続文字列取得
            string sc = sqlControl.obcConnectSting.get(_dbName);
            sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            // 登録済み勤務体系（シフト）コード検証
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select LaborSystemCode, LaborSystemName from tbLaborSystem ");
            sb.Append("where LaborSystemCode = '" + s.ToString().PadLeft(4, '0') + "'");
            
            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                dm = true;
                break;
            }

            dR.Close();

            return dm;
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     事由コードチェック </summary>
        /// <param name="sdCon">
        ///     qlControl.DataControl オブジェクト </param>
        /// <param name="j">
        ///     事由コード</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        private bool chkJiyu(string s)
        {
            bool dm = false;

            // 奉行SQLServer接続文字列取得
            string sc = sqlControl.obcConnectSting.get(_dbName);
            sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            // 登録済み事由コード検証
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select LaborReasonCode from tbLaborReason ");
            sb.Append("where LaborReasonCode = '" + s.PadLeft(2, '0') + "'");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                dm = true;
                break;
            }

            dR.Close();

            return dm;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     指定した社員番号の件数を調べる </summary>
        /// <param name="mList">
        ///     List(DataSet1.勤務票明細Row)</param>
        /// <param name="sNum">
        ///     社員番号</param>
        /// <returns>
        ///     件数</returns>
        ///------------------------------------------------------------------------------------
        private bool getSameNumber(List<DataSet1.勤務票明細Row> mList, string sNum)
        {
            bool rtn = true;

            if (sNum == string.Empty) return rtn;

            // 指定した社員番号の件数を調べる
            if (mList.Count(a => Utility.StrtoInt(a.社員番号) == Utility.StrtoInt(sNum) && a.取消 == global.FLGOFF) > 1)
            {
                rtn = false;
            }

            return rtn;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     明細記入チェック </summary>
        /// <param name="obj">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="iX">
        ///     行を表すインデックス</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckRow(DataSet1.勤務票明細Row m, string tittle, int iX)
        {
            // 社員番号以外に記入項目なしのときエラーとする
            if (m.社員番号 != string.Empty && m.出勤時 == string.Empty &&
                m.出勤分 == string.Empty && m.退勤時 == string.Empty &&
                m.退勤分 == string.Empty && m.残業理由1 == string.Empty &&
                m.残業時1 == string.Empty && m.残業分1 == string.Empty &&
                m.残業理由2 == string.Empty && m.残業時2 == string.Empty && 
                m.残業分2 == string.Empty && m.事由1 == string.Empty && 
                m.事由2 == string.Empty && m.事由3 == string.Empty && 
                m.応援 == global.FLGOFF && m.シフト通り == global.FLGOFF && 
                m.シフトコード == string.Empty)
            {
                setErrStatus(eSH, iX - 1, tittle + "が未入力です");
                return false;
            }

            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     時間記入チェック </summary>
        /// <param name="obj">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="Tani">
        ///     分記入単位</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <param name="stKbn">
        ///     勤怠記号の出勤怠区分</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckTime(DataSet1.勤務票明細Row m, string tittle, int Tani, int iX)
        {
            ///* 勤怠記号が「2:休日勤務」「3:早出」「4:残業」「0:半休」「7:遅刻」「8:早退」「B:宿日直」「C:保安」で
            //   時刻が無記入のときNGとする
            //   2014/10/10 条件より「B:宿日直」「C:保安」を撤廃（始業、終業時刻の記入は必要なし）*/

            //string kigou = m.勤怠記号1.Trim() + m.勤怠記号2.Trim();

            //if (kigou.Contains(KINTAIKIGOU_0) || kigou.Contains(KINTAIKIGOU_2) || kigou.Contains(KINTAIKIGOU_3) ||
            //    kigou.Contains(KINTAIKIGOU_4) || kigou.Contains(KINTAIKIGOU_7) || kigou.Contains(KINTAIKIGOU_8)
            //    //kigou.Contains(KINTAIKIGOU_B) || kigou.Contains(KINTAIKIGOU_C))
            //    )
            //{
            //    if (m.開始時 == string.Empty)
            //    {
            //        setErrStatus(eSH, iX - 1, tittle + "が未入力です");
            //        return false;
            //    }

            //    if (m.開始分 == string.Empty)
            //    {
            //        setErrStatus(eSM, iX - 1, tittle + "が未入力です");
            //        return false;
            //    }

            //    if (m.終了時 == string.Empty)
            //    {
            //        setErrStatus(eEH, iX - 1, tittle + "が未入力です");
            //        return false;
            //    }

            //    if (m.終了分 == string.Empty)
            //    {
            //        setErrStatus(eEM, iX - 1, tittle + "が未入力です");
            //        return false;
            //    }
            //}

            //// 勤怠記号が「5:欠勤」「9:有休」「A:特別休暇」で時刻が記入されているときNGとする
            //if (kigou.Contains(KINTAIKIGOU_5) || kigou.Contains(KINTAIKIGOU_9) || kigou.Contains(KINTAIKIGOU_A))
            //{
            //    string kigouMsg = string.Empty;

            //    if (kigou == KINTAIKIGOU_5) kigouMsg = "5.欠勤";
            //    else if (kigou == KINTAIKIGOU_9) kigouMsg = "9.有休";
            //    else if (kigou == KINTAIKIGOU_A) kigouMsg = "A.特別休暇";

            //    if (m.開始時 != string.Empty)
            //    {
            //        setErrStatus(eSH, iX - 1, "勤怠区分が「" + kigouMsg + "」で" + tittle + "が入力されています");
            //        return false;
            //    }

            //    if (m.開始分 != string.Empty)
            //    {
            //        setErrStatus(eSM, iX - 1, "勤怠区分が「" + kigouMsg + "」で" + tittle + "が入力されています");
            //        return false;
            //    }

            //    if (m.終了時 != string.Empty)
            //    {
            //        setErrStatus(eEH, iX - 1, "勤怠区分が「" + kigouMsg + "」で" + tittle + "が入力されています");
            //        return false;
            //    }

            //    if (m.終了分 != string.Empty)
            //    {
            //        setErrStatus(eEM, iX - 1, "勤怠区分が「" + kigouMsg + "」で" + tittle + "が入力されています");
            //        return false;
            //    }
            //}
            
            // 出勤時間と退勤時間
            string sTimeW = m.出勤時.Trim() + m.出勤分.Trim();
            string eTimeW = m.退勤時.Trim() + m.退勤分.Trim();

            if (sTimeW != string.Empty && eTimeW == string.Empty)
            {
                setErrStatus(eEH, iX - 1, tittle + "退勤時刻が未入力です");
                return false;
            }

            if (sTimeW == string.Empty && eTimeW != string.Empty)
            {
                setErrStatus(eSH, iX - 1, tittle + "出勤時刻が未入力です");
                return false;
            }

            // 記入のとき
            if (m.出勤時 != string.Empty || m.出勤分 != string.Empty ||
                m.退勤時 != string.Empty || m.退勤分 != string.Empty)
            {
                // 数字範囲、単位チェック
                if (!checkHourSpan(m.出勤時))
                {
                    setErrStatus(eSH, iX - 1, tittle + "が正しくありません");
                    return false;
                }

                if (!checkMinSpan(m.出勤分, Tani))
                {
                    setErrStatus(eSM, iX - 1, tittle + "が正しくありません");
                    return false;
                }

                if (!checkHourSpan(m.退勤時))
                {
                    setErrStatus(eEH, iX - 1, tittle + "が正しくありません");
                    return false;
                }

                if (!checkMinSpan(m.退勤分, Tani))
                {
                    setErrStatus(eEM, iX - 1, tittle + "が正しくありません");
                    return false;
                }

                //// 終了時刻範囲
                //if (Utility.StrtoInt(Utility.NulltoStr(m.終了時)) == 24 &&
                //    Utility.StrtoInt(Utility.NulltoStr(m.終了分)) > 0)
                //{
                //    setErrStatus(eEM, iX - 1, tittle + "終了時刻範囲を超えています（～２４：００）");
                //    return false;
                //}
            }

            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     時間記入範囲チェック 0～23の数値 </summary>
        /// <param name="h">
        ///     記入値</param>
        /// <returns>
        ///     正常:true, エラー:false</returns>
        ///------------------------------------------------------------------------------------
        private bool checkHourSpan(string h)
        {
            if (!Utility.NumericCheck(h)) return false;
            else if (int.Parse(h) < 0 || int.Parse(h) > 23) return false;
            else return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     分記入範囲チェック：0～59の数値及び記入単位 </summary>
        /// <param name="h">
        ///     記入値</param>
        /// <param name="tani">
        ///     記入単位分</param>
        /// <returns>
        ///     正常:true, エラー:false</returns>
        ///------------------------------------------------------------------------------------
        private bool checkMinSpan(string m, int tani)
        {
            if (!Utility.NumericCheck(m)) return false;
            else if (int.Parse(m) < 0 || int.Parse(m) > 59) return false;
            else if (int.Parse(m) % tani != 0) return false;
            else return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     時間外記入チェック </summary>
        /// <param name="obj">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="Tani">
        ///     分記入単位</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckZan(DataSet1.勤務票明細Row m, string tittle, int Tani, int iX)
        {
            //// 無記入なら終了
            //if (m.時間外時 == string.Empty && m.時間外分 == string.Empty) return true;
            
            ////  始業、終業時刻が無記入で普通残業が記入されているときエラー
            //if (m.開始時 == string.Empty && m.開始分 == string.Empty &&
            //     m.終了時 == string.Empty && m.終了分 == string.Empty)
            //{
            //    if (m.時間外時 != string.Empty)
            //    {
            //        setErrStatus(eZH, iX - 1, "始業、終業時刻が無記入で" + tittle + "が入力されています");
            //        return false;
            //    }

            //    if (m.時間外分 != string.Empty)
            //    {
            //        setErrStatus(eZM, iX - 1, "始業、終業時刻が無記入で" + tittle + "が入力されています");
            //        return false;
            //    }
            //}

            // 記入のとき
            //if (m.時間外時 != string.Empty || m.時間外分 != string.Empty)
            //{
            //    // 時間と分のチェック
            //    //if (!checkHourSpan(m.時間外時))
            //    //{
            //    //    setErrStatus(eZH, iX - 1, tittle + "が正しくありません");
            //    //    return false;
            //    //}

            //    if (!checkMinSpan(m.時間外分, Tani))
            //    {
            //        setErrStatus(eZM, iX - 1, tittle + "が正しくありません。（" + Tani.ToString() + "分単位）");
            //        return false;
            //    }
            //}

            return true;
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
        private bool errCheckZanTm(DataSet1.勤務票明細Row m, string tittle, int iX, Int64 zan)
        {
            //Int64 mZan = 0;

            //mZan = (Utility.StrtoInt(m.時間外時) * 60) + Utility.StrtoInt(m.時間外分);

            //// 記入時間と計算された残業時間が不一致のとき
            //if (zan != mZan)
            //{
            //    Int64 hh = zan / 60;
            //    Int64 mm = zan % 60;

            //    setErrStatus(eZH, iX - 1, tittle + "が正しくありません。（" + hh.ToString() + "時間" + mm.ToString() + "分）");
            //    return false;
            //}

            return true;
        }

        /// ----------------------------------------------------------------------------------
        /// <summary>
        ///     時間外算出 2015/09/16 </summary>
        /// <param name="m">
        ///     SCCSDataSet.勤務票明細Row </param>
        /// <param name="Tani">
        ///     丸め単位・分</param>
        /// <param name="ws">
        ///     1日の所定労働時間</param>
        /// <returns>
        ///     時間外・分</returns>
        /// ----------------------------------------------------------------------------------
        public Int64 getZangyoTime(DataSet1.勤務票明細Row m, Int64 Tani, Int64 ws, Int64 restTime, out Int64 s10Rest, int taikeiCode)
        {
            Int64 zan = 0;  // 計算後時間外勤務時間
            s10Rest = 0;    // 深夜勤務時間帯の10分休憩時間

            DateTime cTm;
            DateTime sTm;
            DateTime eTm;
            DateTime zsTm;
            DateTime pTm;

            if (!m.Is出勤時Null() && !m.Is出勤分Null() && !m.Is出勤時Null() && !m.Is出勤分Null())
            {
                int ss = Utility.StrtoInt(m.出勤時) * 100 + Utility.StrtoInt(m.出勤分);
                int ee = Utility.StrtoInt(m.退勤時) * 100 + Utility.StrtoInt(m.退勤分);
                DateTime dt = DateTime.Today;
                string sToday = dt.Year.ToString() + "/" + dt.Month.ToString() + "/" + dt.Day.ToString();

                // 始業時刻
                if (DateTime.TryParse(sToday + " " + m.出勤時 + ":" + m.出勤分, out cTm))
                {
                    sTm = cTm;
                }
                else return 0;

                // 終業時刻
                if (ss > ee)
                {
                    // 翌日
                    dt = DateTime.Today.AddDays(1);
                    sToday = dt.Year.ToString() + "/" + dt.Month.ToString() + "/" + dt.Day.ToString();
                    if (DateTime.TryParse(sToday + " " + m.退勤時 + ":" + m.退勤分, out cTm))
                    {
                        eTm = cTm;
                    }
                    else return 0;
                }
                else
                {
                    // 同日
                    if (DateTime.TryParse(sToday + " " + m.退勤時 + ":" + m.退勤分, out cTm))
                    {
                        eTm = cTm;
                    }
                    else return 0;
                }


                //MessageBox.Show(sTm.ToShortDateString() + " " + sTm.ToShortTimeString() + "    " + eTm.ToShortDateString() + " " + eTm.ToShortTimeString());


                // 作業日報に記入されている始業から就業までの就業時間取得
                double w = Utility.GetTimeSpan(sTm, eTm).TotalMinutes - restTime;

                // 所定労働時間内なら時間外なし
                if (w <= ws)
                {
                    return 0;
                }

                // 所定労働時間＋休憩時間＋10分または15分経過後の時刻を取得（時間外開始時刻）
                zsTm = sTm.AddMinutes(ws);          // 所定労働時間
                zsTm = zsTm.AddMinutes(restTime);   // 休憩時間
                int zSpan = 0;

                if (taikeiCode == 100)
                {
                    zsTm = zsTm.AddMinutes(10);         // 体系コード：100 所定労働時間後の10分休憩
                    zSpan = 130;
                }
                else if (taikeiCode == 200 || taikeiCode == 300)
                {
                    zsTm = zsTm.AddMinutes(15);         // 体系コード：200,300 所定労働時間後の15分休憩
                    zSpan = 135;
                }
                
                pTm = zsTm;                         // 時間外開始時刻

                // 該当時刻から終業時刻まで130分または135分以上あればループさせる
                while (Utility.GetTimeSpan(pTm, eTm).TotalMinutes > zSpan)
                {
                    // 終業時刻まで2時間につき10分休憩として時間外を算出
                    // 時間外として2時間加算
                    zan += 120;

                    // 130分、または135分後の時刻を取得（2時間＋10分、または15分）
                    pTm = pTm.AddMinutes(zSpan);

                    // 深夜勤務時間中の10分または15分休憩時間を取得する
                    s10Rest += getShinya10Rest(pTm, eTm, zSpan - 120);
                }

                // 130分（135分）以下の時間外を加算
                zan += (Int64)Utility.GetTimeSpan(pTm, eTm).TotalMinutes;

                // 単位で丸める
                zan -= (zan % Tani);

                //MessageBox.Show(pTm.ToShortDateString() + "    " + eTm.ToShortDateString());
            }
                        
            return zan;
        }

        /// --------------------------------------------------------------------
        /// <summary>
        ///     深夜勤務時間中の10分休憩時間を取得する </summary>
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
        ///     深夜勤務記入チェック </summary>
        /// <param name="m">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="Tani">
        ///     分記入単位</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckShinya(DataSet1.勤務票明細Row m, string tittle, int Tani, int iX)
        {
        //    // 無記入なら終了
        //    if (m.深夜時 == string.Empty && m.深夜分 == string.Empty) return true;

        //    //  始業、終業時刻が無記入で深夜が記入されているときエラー
        //    if (m.開始時 == string.Empty && m.開始分 == string.Empty &&
        //         m.終了時 == string.Empty && m.終了分 == string.Empty)
        //    {
        //        if (m.深夜時 != string.Empty)
        //        {
        //            setErrStatus(eSIH, iX - 1, "始業、終業時刻が無記入で" + tittle + "が入力されています");
        //            return false;
        //        }

        //        if (m.深夜分 != string.Empty)
        //        {
        //            setErrStatus(eSIM, iX - 1, "始業、終業時刻が無記入で" + tittle + "が入力されています");
        //            return false;
        //        }
        //    }

        //    // 記入のとき
        //    if (m.深夜時 != string.Empty || m.深夜分 != string.Empty)
        //    {
        //        // 時間と分のチェック
        //        //if (!checkHourSpan(m.時間外時))
        //        //{
        //        //    setErrStatus(eZH, iX - 1, tittle + "が正しくありません");
        //        //    return false;
        //        //}

        //        if (!checkMinSpan(m.深夜分, Tani))
        //        {
        //            setErrStatus(eSIM, iX - 1, tittle + "が正しくありません。（" + Tani.ToString() + "分単位）");
        //            return false;
        //        }
        //    }

            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     深夜勤務時間チェック </summary>
        /// <param name="m">
        ///     勤務票明細Rowコレクション</param>
        /// <param name="tittle">
        ///     チェック項目名称</param>
        /// <param name="iX">
        ///     日付を表すインデックス</param>
        /// <param name="shinya">
        ///     算出された深夜k勤務時間</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckShinyaTm(DataSet1.勤務票明細Row m, string tittle, int iX, Int64 shinya)
        {
            Int64 mShinya = 0;

            //mShinya = (Utility.StrtoInt(m.深夜時) * 60) + Utility.StrtoInt(m.深夜分);

            //// 記入時間と計算された深夜時間が不一致のとき
            //if (shinya != mShinya)
            //{
            //    Int64 hh = shinya / 60;
            //    Int64 mm = shinya % 60;

            //    setErrStatus(eSIH, iX - 1, tittle + "が正しくありません。（" + hh.ToString() + "時間" + mm.ToString() + "分）");
            //    return false;
            //}

            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     実働時間を取得する</summary>
        /// <param name="sH">
        ///     開始時</param>
        /// <param name="sM">
        ///     開始分</param>
        /// <param name="eH">
        ///     終了時</param>
        /// <param name="eM">
        ///     終了分</param>
        /// <param name="rH">
        ///     休憩時間・分</param>
        /// <returns>
        ///     実働時間</returns>
        ///------------------------------------------------------------------------------------
        public double getWorkTime(string sH, string sM, string eH, string eM, int rH)
        {
            DateTime sTm;
            DateTime eTm;
            DateTime cTm;
            double w = 0;   // 稼働時間

            // 時刻情報に不備がある場合は０を返す
            if (!Utility.NumericCheck(sH) || !Utility.NumericCheck(sM) || 
                !Utility.NumericCheck(eH) || !Utility.NumericCheck(eM))
                return 0;

            int ss = Utility.StrtoInt(sH) * 100 + Utility.StrtoInt(sM);
            int ee = Utility.StrtoInt(eH) * 100 + Utility.StrtoInt(eM);
            DateTime dt = DateTime.Today;
            string sToday = dt.Year.ToString() + "/" + dt.Month.ToString() + "/" + dt.Day.ToString();

            // 開始時刻取得
            if (Utility.StrtoInt(sH) == 24)
            {
                if (DateTime.TryParse(sToday + " 0:" + Utility.StrtoInt(sM).ToString(), out cTm))
                {
                    sTm = cTm;
                }
                else return 0;
            }
            else
            {
                if (DateTime.TryParse(sToday + " " + Utility.StrtoInt(sH).ToString() + ":" + Utility.StrtoInt(sM).ToString(), out cTm))
                {
                    sTm = cTm;
                }
                else return 0;
            }
            
            // 終業時刻
            if (ss > ee)
            {
                // 翌日
                dt = DateTime.Today.AddDays(1);
                sToday = dt.Year.ToString() + "/" + dt.Month.ToString() + "/" + dt.Day.ToString();
            }

            // 終了時刻取得
            if (Utility.StrtoInt(eH) == 24)
                eTm = DateTime.Parse(sToday + " 23:59");
            else
            {
                if (DateTime.TryParse(sToday + " " + Utility.StrtoInt(eH).ToString() + ":" + Utility.StrtoInt(eM).ToString(), out cTm))
                {
                    eTm = cTm;
                }
                else return 0;
            }

            // 終了時間が24:00記入のときは23:59までの計算なので稼働時間1分加算する
            if (Utility.StrtoInt(eH) == 24 && Utility.StrtoInt(eM) == 0)
            {
                w = Utility.GetTimeSpan(sTm, eTm).TotalMinutes + 1;
            }
            else if (sTm == eTm)    // 同時刻の場合は翌日の同時刻とみなす 2014/10/10
            {
                w = Utility.GetTimeSpan(sTm, eTm.AddDays(1)).TotalMinutes;  // 稼働時間
            }
            else
            {
                w = Utility.GetTimeSpan(sTm, eTm).TotalMinutes;  // 稼働時間
            }

            // 休憩時間を差し引く
            if (w >= rH) w = w - rH;
            else w = 0;

            // 値を返す
            return w;
        }

        ///--------------------------------------------------------------
        /// <summary>
        ///     深夜勤務時間を取得する</summary>
        /// <param name="sH">
        ///     開始時</param>
        /// <param name="sM">
        ///     開始分</param>
        /// <param name="eH">
        ///     終了時</param>
        /// <param name="eM">
        ///     終了分</param>
        /// <param name="tani">
        ///     丸め単位</param>
        /// <param name="s10">
        ///     深夜勤務時間中の10分休憩</param>
        /// <returns>
        ///     深夜勤務時間・分</returns>
        /// ------------------------------------------------------------
        public double getShinyaWorkTime(string sH, string sM, string eH, string eM, int tani, Int64 s10)
        {
            DateTime sTime;
            DateTime eTime;
            DateTime cTm;

            double wkShinya = 0;    // 深夜稼働時間

            // 時刻情報に不備がある場合は０を返す
            if (!Utility.NumericCheck(sH) || !Utility.NumericCheck(sM) ||
                !Utility.NumericCheck(eH) || !Utility.NumericCheck(eM))
                return 0;

            // 開始時間を取得
            if (DateTime.TryParse(Utility.StrtoInt(sH).ToString() + ":" + Utility.StrtoInt(sM).ToString(), out cTm))
            {
                sTime = cTm;
            }
            else return 0;

            // 終了時間を取得
            if (Utility.StrtoInt(eH) == 24 && Utility.StrtoInt(eM) == 0)
            {
                eTime = global.dt2359;
            }
            else if (DateTime.TryParse(Utility.StrtoInt(eH).ToString() + ":" + Utility.StrtoInt(eM).ToString(), out cTm))
            {
                eTime = cTm;
            }
            else return 0;


            // 当日内の勤務のとき
            if (sTime.TimeOfDay < eTime.TimeOfDay)
            {
                // 早出残業時間を求める
                if (sTime < global.dt0500)  // 開始時刻が午前5時前のとき
                {
                    // 早朝時間帯稼働時間
                    if (eTime >= global.dt0500)
                    {
                        wkShinya += Utility.GetTimeSpan(sTime, global.dt0500).TotalMinutes;
                    }
                    else
                    {
                        wkShinya += Utility.GetTimeSpan(sTime, eTime).TotalMinutes;
                    }
                }

                // 終了時刻が22:00以降のとき
                if (eTime >= global.dt2200)
                {
                    // 当日分の深夜帯稼働時間を求める
                    if (sTime <= global.dt2200)
                    {
                        // 出勤時刻が22:00以前のとき深夜開始時刻は22:00とする
                        wkShinya += Utility.GetTimeSpan(global.dt2200, eTime).TotalMinutes;
                    }
                    else
                    {
                        // 出勤時刻が22:00以降のとき深夜開始時刻は出勤時刻とする
                        wkShinya += Utility.GetTimeSpan(sTime, eTime).TotalMinutes;
                    }

                    // 終了時間が24:00記入のときは23:59までの計算なので稼働時間1分加算する
                    if (Utility.StrtoInt(eH) == 24 && Utility.StrtoInt(eM) == 0)
                        wkShinya += 1;
                }
            }
            else
            {
                // 日付を超えて終了したとき（開始時刻 >= 終了時刻）※2014/10/10 同時刻は翌日の同時刻とみなす

                // 早出残業時間を求める
                if (sTime < global.dt0500)  // 開始時刻が午前5時前のとき
                {
                    wkShinya += Utility.GetTimeSpan(sTime, global.dt0500).TotalMinutes;
                }

                // 当日分の深夜勤務時間（～０：００まで）
                if (sTime <= global.dt2200)
                {
                    // 出勤時刻が22:00以前のとき無条件に120分
                    wkShinya += global.TOUJITSU_SINYATIME;
                }
                else
                {
                    // 出勤時刻が22:00以降のとき出勤時刻から24:00までを求める
                    wkShinya += Utility.GetTimeSpan(sTime, global.dt2359).TotalMinutes + 1;
                }

                // 0:00以降の深夜勤務時間を加算（０：００～終了時刻）
                if (eTime.TimeOfDay > global.dt0500.TimeOfDay)
                {
                    wkShinya += Utility.GetTimeSpan(global.dt0000, global.dt0500).TotalMinutes;
                }
                else
                {
                    wkShinya += Utility.GetTimeSpan(global.dt0000, eTime).TotalMinutes;
                }
            }

            // 深夜勤務時間中の10分または15分休憩時間を差し引く
            wkShinya -= s10;

            // 単位分で丸め
            wkShinya -= (wkShinya % tani);

            return wkShinya;
        }
    }
}
