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
    class OCRData
    {
        public OCRData(string dbName, xlsData exlms)
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

        enum errCode
        {
            eNothing, eYearMonth, eMonth, eDay, eKinmuTaikeiCode
        }

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
        ///     エラー項目 = 確認チェック </summary>
        public int eDataCheck = 35;

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
        public int eShainNo2 = 27;

        /// <summary> 
        ///     エラー項目 = 勤務記号 </summary>
        public int eKintaiKigou = 6;

        /// <summary> 
        ///     エラー項目 = 部署コード </summary>
        public int eBushoCode = 31;

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

        /// <summary> 
        ///     エラー項目 = ライン </summary>
        public int eLine = 23;
        public int eLine2 = 28;

        /// <summary> 
        ///     エラー項目 = 部門 </summary>
        public int eBmn = 24;
        public int eBmn2 = 29;

        /// <summary> 
        ///     エラー項目 = 製品群 </summary>
        public int eHin = 25;
        public int eHin2 = 30;

        /// <summary> 
        ///     エラー項目 = 応援分 </summary>
        public int eOuenM = 26;

        /// <summary> 
        ///     エラー項目 = 応援分 </summary>
        public int eOuenIP = 32;
        public int eOuenIP2 = 33;

        /// <summary> 
        ///     エラー項目 = 応援移動票と勤怠データＩ／Ｐ票 </summary>
        public int eIpOuen = 34;

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
        
        // 雇用区分：派遣
        private const int KOYOU_HAKEN = 10;

        // 雇用区分：パート 2017/10/19
        private const int KOYOU_PART = 3;
        private const int KOYOU_PART8 = 8;
        private const int KOYOU_PART9= 9;

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
            r.年 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[3].Trim().Replace("-", ""), 4));
            r.月 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[4].Trim().Replace("-", ""), 2));
            r.日 = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[5].Trim().Replace("-", ""), 2));
            r.部署コード = Utility.GetStringSubMax(stCSV[6].Trim().Replace("-", ""), 5);
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
            r.部署コード = Utility.GetStringSubMax(stCSV[6].Trim().Replace("-", ""), 5);
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

            // 社員番号：先頭ゼロは除去
            string sN = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[2].Trim().Replace("-", ""), 6)).ToString();
            if (sN != global.FLGOFF)
            {
                r.社員番号 = sN;
            }
            else
            {
                r.社員番号 = string.Empty;
            }

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

            // 残業理由１：先頭ゼロは除去
            sN = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[11].Trim().Replace("-", ""), 2)).ToString();
            if (sN != global.FLGOFF)
            {
                r.残業理由1 = sN;
            }
            else
            {
                r.残業理由1 = string.Empty;
            }

            r.残業時1 = Utility.GetStringSubMax(stCSV[12].Trim().Replace("-", ""), 2);
            r.残業分1 = Utility.GetStringSubMax(stCSV[13].Trim().Replace("-", ""), 1);

            // 残業理由２：先頭ゼロは除去
            sN = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[14].Trim().Replace("-", ""), 2)).ToString();
            if (sN != global.FLGOFF)
            {
                r.残業理由2 = sN;
            }
            else
            {
                r.残業理由2 = string.Empty;
            }

            r.残業時2 = Utility.GetStringSubMax(stCSV[15].Trim().Replace("-", ""), 2);
            r.残業分2 = Utility.GetStringSubMax(stCSV[16].Trim().Replace("-", ""), 1);
            r.取消 = global.FLGOFF;
            r.データ領域名 = string.Empty;
            r.更新年月日 = DateTime.Now;
            r.帰宅後勤務ID = string.Empty;
            r.社員名 = string.Empty;
            
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

            // 社員番号：先頭ゼロは除去
            string sN = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[2].Trim().Replace("-", ""), 6)).ToString();
            if (sN != global.FLGOFF)
            {
                r.社員番号 = sN;
            }
            else
            {
                r.社員番号 = string.Empty;
            }

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

                // 残業理由１：先頭ゼロは除去
                sN = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[6].Trim().Replace("-", ""), 2)).ToString();
                if (sN != global.FLGOFF)
                {
                    r.残業理由1 = sN;
                }
                else
                {
                    r.残業理由1 = string.Empty;
                }

                r.残業時1 = Utility.GetStringSubMax(stCSV[7].Trim().Replace("-", ""), 2);
                r.残業分1 = Utility.GetStringSubMax(stCSV[8].Trim().Replace("-", ""), 1);


                // 残業理由２：先頭ゼロは除去
                sN = Utility.StrtoInt(Utility.GetStringSubMax(stCSV[9].Trim().Replace("-", ""), 2)).ToString();
                if (sN != global.FLGOFF)
                {
                    r.残業理由2 = sN;
                }
                else
                {
                    r.残業理由2 = string.Empty;
                }

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
        ///     勤怠データエラーチェックメイン処理。
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
        public Boolean errCheckMain(int sIx, int eIx, Form frm, DataSet1 dts, string[] cID)
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

            // 奉行SQLServer接続文字列取得
            string sc = sqlControl.obcConnectSting.get(_dbName);

            // 奉行SQLServer接続
            sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

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
                    //DataSet1.勤務票ヘッダRow r = (DataSet1.勤務票ヘッダRow)dts.勤務票ヘッダ.Rows[i];
                    DataSet1.勤務票ヘッダRow r = dts.勤務票ヘッダ.Single(a => a.ID == cID[i]);

                    // エラーチェック実施
                    eCheck = errCheckData(sdCon, dts, r);

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

            // 奉行SQLServer接続コネクション閉じる
            sdCon.Close();

            return eCheck;
        }

        ///--------------------------------------------------------------------------------------------------
        /// <summary>
        ///     過去データエラーチェックメイン処理。
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
        public Boolean errCheckMain(string dID, DataSet1 dts)
        {
            // 出勤簿データ読み出し
            Boolean eCheck = true;

            // 奉行SQLServer接続文字列取得
            string sc = sqlControl.obcConnectSting.get(_dbName);

            // 奉行SQLServer接続
            sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            // 過去勤務票ヘッダ行のコレクションを取得します
            DataSet1.過去勤務票ヘッダRow r = dts.過去勤務票ヘッダ.Single(a => a.ID == dID);

            // エラーチェック実施
            eCheck = errCheckData(sdCon, dts, r);

            if (!eCheck)　//エラーがあったとき
            {
                _errHeaderIndex = 0;     // エラーとなったヘッダRowIndex
            }
            
            // 奉行SQLServer接続コネクション閉じる
            sdCon.Close();

            return eCheck;
        }


        ///--------------------------------------------------------------------------------------------------
        /// <summary>
        ///     応援移動票エラーチェックメイン処理。
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
        public Boolean errCheckMainOuen(int sIx, int eIx, Form frm, DataSet1 dts, string [] hArray, string [] cID)
        {
            int rCnt = 0;

            // オーナーフォームを無効にする
            frm.Enabled = false;

            // プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = frm;
            frmP.Show();

            // レコード件数取得
            int cTotal = dts.応援移動票ヘッダ.Rows.Count;

            // 応援移動票データ読み出し
            Boolean eCheck = true;

            // 奉行SQLServer接続文字列取得
            string sc = sqlControl.obcConnectSting.get(_dbName);
            sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

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
                    // 応援移動票ヘッダ行のコレクションを取得します
                    //DataSet1.応援移動票ヘッダRow r = (DataSet1.応援移動票ヘッダRow)dts.応援移動票ヘッダ.Rows[i];
                    DataSet1.応援移動票ヘッダRow r = dts.応援移動票ヘッダ.Single(a => a.ID == cID[i]);

                    // エラーチェック実施
                    eCheck = errCheckOuen(sdCon, dts, r, hArray);

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

            // コネクション閉じる
            sdCon.Close();

            return eCheck;
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
        public Boolean errCheckMainOuen(string dID, DataSet1 dts, string[] hArray)
        {
            Boolean eCheck = true;

            // 奉行SQLServer接続文字列取得
            string sc = sqlControl.obcConnectSting.get(_dbName);
            sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            // 過去応援移動票ヘッダ行のコレクションを取得します
            DataSet1.過去応援移動票ヘッダRow r = dts.過去応援移動票ヘッダ.Single(a => a.ID == dID);

            // エラーチェック実施
            eCheck = errCheckOuen(sdCon, dts, r, hArray);

            if (!eCheck)　//エラーがあったとき
            {
                _errHeaderIndex = 0;     // エラーとなったヘッダRowIndex
            }

            // コネクション閉じる
            sdCon.Close();

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
        public Boolean errCheckData(sqlControl.DataControl sdCon, DataSet1 dts, DataSet1.勤務票ヘッダRow r)
        {
            string sDate;
            DateTime eDate;
            DataSet1 _dts = dts;

            // 確認チェック
            if (r.確認 == global.flgOff)
            {
                setErrStatus(eDataCheck, 0, "未確認の出勤簿です");
                return false;
            }

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

            // 部署コードチェック
            string dCode = getDepartmentCode(r.部署コード.ToString());
            if (!chkDepartmentCode(dCode, sdCon))
            {
                setErrStatus(eBushoCode, 0, "マスター未登録の部署コードです");
                return false;
            }

            // 勤務体系コード就業奉行登録チェック
            if (!chkSftCode(sdCon, r.シフトコード.ToString()))
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

                // 登録済み社員番号マスター検証 : 出勤簿日付で判断 2017/09/28
                if (!chkShainCode(m.社員番号, sdCon, eDate))
                {
                    setErrStatus(eShainNo, iX - 1, "未登録もしくは" + eDate.ToShortDateString() + "現在、在籍していない社員番号です");
                    return false;
                }

                // 同一作業日報内で同じ社員番号が複数記入されているとエラー 2015/04/15
                if (!getSameNumber(mList, m.社員番号))
                {
                    setErrStatus(eShainNo, iX - 1, "同じ社員番号のデータが複数あります");
                    return false;
                }

                // 同日・同社員番号が存在するときエラー 2017/10/20
                if (!getSameDateNumber(dts, r.ID, r.年, r.月, r.日, m.社員番号))
                {
                    setErrStatus(eShainNo, iX - 1, "別の出勤簿に同日付で同じ社員番号のデータが存在します");
                    return false;
                }

                // 明細記入チェック
                if (!errCheckRow(m, "勤怠データＩ／Ｐ票内容", iX)) return false;

                // シフト通りチェック
                if (m.シフト通り == global.FLGON && m.シフトコード != string.Empty)
                {
                    setErrStatus(eChksft, iX - 1, "「シフト通り」で変更シフトコードが記入されています");
                    return false;
                }

                //// シフト通りではない場合：2017/07/19
                //if (m.シフト通り == global.FLGOFF)
                //{
                //    if (m.シフトコード == string.Empty && 
                //        m.出勤時 == string.Empty && m.出勤分 == string.Empty && 
                //        m.退勤時 == string.Empty && m.退勤分 == string.Empty)
                //    {
                //        setErrStatus(eChksft, iX - 1, "「シフト通りではない」とき変更シフトコードまたは勤務時間の記入が必要です");
                //        return false;
                //    }
                //}
                
                // シフト通りと事由
                if (!errCheckSftJiyu(m))
                {
                    setErrStatus(eChksft, iX - 1, "「シフト通り」で事由が記入されています");
                    return false;
                }

                // シフト通りと残業
                if (!errCheckSftZangyo(m))
                {
                    setErrStatus(eChksft, iX - 1, "「シフト通り」で残業が記入されています");
                    return false;
                }

                // シフト通りと休出：2018/02/06
                if (!errCheckSftKyushutsu(r, m))
                {
                    setErrStatus(eChksft, iX - 1, "休日出勤で「シフト通り」は記入できません");
                    return false;
                }

                //// 休出：残業チェック 2018/02/02 コメント化
                //if (!errKyushutsuZangyo(r, m))
                //{
                //    setErrStatus(eZanRe1, iX - 1, "「休日出勤」で残業が記入されていません");
                //    return false;
                //}

                // 休出：事由チェック
                int jNum = 0;
                if (!errKyushutsuJiyu(r, m, out jNum))
                {
                    int eNum = 0;

                    if (jNum == 1)
                    {
                        eNum = eJiyu1;
                    }
                    else if (jNum == 2)
                    {
                        eNum = eJiyu2;
                    }
                    else if (jNum == 3)
                    {
                        eNum = eJiyu3;
                    }

                    setErrStatus(eNum, iX - 1, "「休日出勤」で事由が記入されています");
                    return false;
                }

                // 事由マスター登録チェック
                string[] mJiyu = { m.事由1, m.事由2, m.事由3 };
                string[] eJiyu = { eJiyu1.ToString(), eJiyu2.ToString(), eJiyu3.ToString() };
                int errNum = 0;

                // 事由マスター登録チェック
                OCR.clsJiyuHas jiyu = new OCR.clsJiyuHas(mJiyu);
                if (!jiyu.isHasRows(out errNum, sdCon))
                {
                    setErrStatus(Utility.StrtoInt(eJiyu[errNum]), iX - 1, "マスター未登録の事由です");
                    return false;
                }

                // 前半欠勤または後半欠勤が単独記入されていたらエラー : 2018/02/15
                OCR.clsJiyuHankekkin jiyuHanke = new OCR.clsJiyuHankekkin(mJiyu);
                if (!jiyuHanke.isHankeAnotherDay(out errNum))
                {
                    setErrStatus(Utility.StrtoInt(eJiyu[errNum]), iX - 1, "前半欠勤または後半欠勤の単独記入は出来ません");
                    return false;
                }

                // 取得単位が終日の事由チェック
                OCR.clsJiyuAllDay jiyu2 = new OCR.clsJiyuAllDay(mJiyu);
                if (!jiyu2.isAllDayAnotherDay(out errNum, sdCon))
                {
                    setErrStatus(Utility.StrtoInt(eJiyu[errNum]), iX - 1, "取得単位が「終日」の事由と他の事由は同時に記入出来ません");
                    return false;
                }


                // 該当社員の雇用区分を取得：2017/09/27
                int kKbn = 0;
                getEmployee(sdCon, m.社員番号, ref kKbn);

                // 「土曜特休」はパート社員、契約社員、アルバイトのみ記入可 2017/11/10
                if (kKbn != KOYOU_PART && kKbn != KOYOU_PART8 && kKbn != KOYOU_PART9)
                {
                    for (int i = 0; i < mJiyu.Length; i++)
                    {
                        if (mJiyu[i] == global.SFT_DOTOKKYU.ToString())
                        {
                            setErrStatus(Utility.StrtoInt(eJiyu[i]), iX - 1, "パート以外は「土曜特休」は記入できません");
                            return false;
                        }
                    }
                }

                // 雇用区分「１０」派遣社員が使用可能な事由は「１０：通常欠勤」のみとする 2017/11/21
                // 「１１：休業欠勤」も対象とする 2018/10/30
                // 「２０：遅刻早退」も対象とする 2020/04/01
                if (kKbn == KOYOU_HAKEN)
                {
                    for (int i = 0; i < mJiyu.Length; i++)
                    {
                        if (mJiyu[i] != string.Empty && 
                            mJiyu[i] != global.SFT_TSUJYOKEKKIN.ToString() &&
                            mJiyu[i] != global.JIYU_KYUGYOKEKKIN.ToString() &&
                            mJiyu[i] != global.JIYU_KYUGYOCHISOU.ToString())
                        {
                            setErrStatus(Utility.StrtoInt(eJiyu[i]), iX - 1, "派遣社員は「10：通常欠勤」「11：休業欠勤」「20:休業遅早」以外は記入できません");
                            return false;
                        }
                    }
                }


                string[] jj = new string[3];
                jj[0] = m.事由1;
                jj[1] = m.事由2;
                jj[2] = m.事由3;

                //if (!Utility.chkJiyu(jj, _dbName))
                //{
                //    setErrStatus(eJiyu1, iX - 1, "取得単位が「終日」の事由と他の事由は同時に記入出来ません");
                //    return false;
                //}

                // 取得単位が終日のときシフト以外記入チェック
                jNum = 0;
                OCR.clsAlldayAnotherData jiyu3 = new OCR.clsAlldayAnotherData(mJiyu);
                if (!jiyu3.isAlldayAnotherData(m, sdCon, out jNum))
                {
                    int eNum = 0;

                    if (jNum == 1)
                    {
                        eNum = eSH;
                    }

                    if (jNum == 2)
                    {
                        eNum = eSM;
                    }

                    if (jNum == 3)
                    {
                        eNum = eEH;
                    }

                    if (jNum == 4)
                    {
                        eNum = eEM;
                    }

                    if (jNum == 5)
                    {
                        eNum = eZanRe1;
                    }

                    if (jNum == 6)
                    {
                        eNum = eZanH1;
                    }

                    if (jNum == 7)
                    {
                        eNum = eZanM1;
                    }

                    if (jNum == 8)
                    {
                        eNum = eZanRe2;
                    }

                    if (jNum == 9)
                    {
                        eNum = eZanH2;
                    }

                    if (jNum == 10)
                    {
                        eNum = eZanM2;
                    }

                    if (jNum == 11)
                    {
                        eNum = eChkOuen;
                    }

                    if (jNum == 12)
                    {
                        eNum = eSftCode;
                    }

                    setErrStatus(eNum, iX - 1, "取得単位が「終日」の事由で他の項目が記入されています");
                    return false;
                }

                // 取得単位が終日で休出のとき記入チェック
                OCR.clsAllDayOffWork jiyu4 = new OCR.clsAllDayOffWork(mJiyu);
                if (!jiyu4.isAllDayOffWork(r, m, sdCon))
                {
                    setErrStatus(Utility.StrtoInt(eJiyu[0]), iX - 1, "休出のとき取得単位が「終日」の事由は使用できません");
                    return false;
                }

                // 半休事由の重複記入チェック
                int divCnt = 0;
                OCR.clsJiyuDiv jiyu5 = new OCR.clsJiyuDiv(mJiyu);
                if (!jiyu5.isJiyuDiv(sdCon, out divCnt))
                {
                    for (int i = 0; i < mJiyu.Length; i++)
                    {
                        if (mJiyu[i] != string.Empty)
                        {
                            setErrStatus(Utility.StrtoInt(eJiyu[i]), iX - 1, "半休事由の記入が正しくありません");
                            break;
                        }
                    }
                    return false;
                }
                else
                {
                    // 半休のとき（前半、または後半休のみのとき）
                    if (divCnt == 1)
                    {
                        // 半休のとき始業時刻・終業時刻は必須入力
                        if (m.出勤時 == string.Empty && m.出勤分 == string.Empty &&
                            m.退勤時 == string.Empty && m.退勤分 == string.Empty)
                        {
                            setErrStatus(eSH, iX - 1, "半休で出退勤時刻が記入されていません");
                            return false;
                        }
                        else
                        {
                            if (!errCheckTime(m, "出退時間", tanMin1, iX)) return false;
                        }

                        // 半休と出退勤時刻
                        string msg = "";
                        int zenkou = 0;
                        if (!chkHankyuTime(sdCon, r, m, jj, out msg, out zenkou))
                        {
                            if (zenkou == 1)
                            {
                                // 前半休のとき
                                setErrStatus(eSH, iX - 1, msg);
                            }
                            else if (zenkou == 2)
                            {
                                // 後半休のとき
                                setErrStatus(eEH, iX - 1, msg);
                            }

                            return false;
                        }
                    }
                    else if (divCnt == 2)
                    {
                        // 前半、後半組み合わせで終日休暇のときシフト以外記入エラー
                        if (m.出勤時 != string.Empty || m.出勤分 != string.Empty || m.退勤時 != string.Empty || m.退勤分 != string.Empty ||
                            m.残業理由1 != string.Empty || m.残業時1 != string.Empty || m.残業分1 != string.Empty ||
                            m.残業理由2 != string.Empty || m.残業時2 != string.Empty || m.残業分2 != string.Empty ||
                            m.応援 == global.FLGON)
                        {
                            setErrStatus(eJiyu1, iX - 1, "前半休＋後半休で終日休みのため他の項目の記入は不要です");
                            return false;
                        }
                    }
                }

                // 前半休＋後半休で終日休みのときはチェック不要 2017/11/13
                if (divCnt != 2)
                {
                    // シフト通りではないとき : 2017/07/26
                    OCR.clsNotAlldayShift eShift = new OCR.clsNotAlldayShift(mJiyu);
                    if (!eShift.isNotAlldayShift(m, sdCon, out jNum))
                    {
                        setErrStatus(eChksft, iX - 1, "「シフト通りではない」とき変更シフトコードまたは勤務時間の記入が必要です");
                        return false;
                    }
                }
                

                //// 変更シフトコード
                //if (m.シフト通り == string.Empty && m.シフトコード == string.Empty)
                //{
                //    setErrStatus(eSftCode, iX - 1, "変更シフトコードが未記入です");
                //    return false;
                //}

                // 勤務体系（シフト）コード奉行登録チェック
                if (m.シフトコード != string.Empty)
                {
                    if (!chkSftCode(sdCon, m.シフトコード))
                    {
                        setErrStatus(eSftCode, iX - 1, "就業奉行に登録されていないシフトコードです");
                        return false;
                    }

                    // 以下、コメント化 2018/02/05
                    //// 休日以外にシフトコード振替出勤は使用不可 2018/02/03
                    //if (m.勤務票ヘッダRow.シフトコード != global.SFT_KYUSHUTSU && 
                    //    Utility.StrtoInt(m.シフトコード) == global.SFT_FURIKAESHUKKIN)
                    //{
                    //    setErrStatus(eSftCode, iX - 1, "休日以外で振替出勤は使用できません");
                    //    return false;
                    //}

                    // 2018/02/03 出退勤時刻が無記入のときエラー
                    if (m.出勤時 == string.Empty && m.出勤分 == string.Empty &&
                        m.退勤時 == string.Empty && m.退勤分 == string.Empty)
                    {
                        setErrStatus(eSH, iX - 1, "シフト変更で出退勤時刻が無記入です");
                        return false;
                    }


                    //// 部署別勤務体系Excelシート登録チェック
                    //if (!bs.getBushoSft(out sftName, r.部署コード.ToString(), m.シフトコード.ToString(), sHol))
                    //{
                    //    setErrStatus(eSftCode, iX - 1, "該当部署に登録されていないか勤務日・休日に該当しないシフトコードです");
                    //    return false;
                    //}


                    /* 休日以外の日に変更シフト欄に[31]休出・休憩なしが記入されているときエラーとしない為、
                     * コメント化 2018/05/28 */
                    // 休日以外の日に変更シフト欄に[31]休出・休憩なしが記入されているとき : 2017/10/19
                    // 休日以外の日に変更シフト欄に[32]休出・休憩ありが記入されているとき : 2018/02/05
                    //if (sHol != global.FLGON && 
                    //    (Utility.StrtoInt(m.シフトコード) == global.SFT_KYUSHUTSU ||
                    //     Utility.StrtoInt(m.シフトコード) == global.SFT_KYUKEI_KYUSHUTSU))
                    //{
                    //    // パートタイマー以外はエラー : 2017/10/19
                    //    if (kKbn != KOYOU_PART && kKbn != KOYOU_PART8 && kKbn != KOYOU_PART9)
                    //    {
                    //        setErrStatus(eSftCode, iX - 1, "休日以外に休出が記入されています");
                    //        return false;
                    //    }
                    //    else
                    //    {
                    //        // パートタイマーでも土曜以外は土曜特休とみなさない
                    //        if (eDate.ToString("ddd") != "土")
                    //        {
                    //            setErrStatus(eSftCode, iX - 1, "休出は記入できません");
                    //            return false;
                    //        }
                    //    }
                    //}
                }

                // 始業時刻・終業時刻チェック : 出退勤は30分単位 2017/08/30
                //if (!errCheckTime(m, "出退時間", tanMin1, iX)) return false;
                if (!errCheckTime(m, "出退時間", tanMin30, iX)) return false;

                // 残業理由
                if (!Utility.chkZangyoRe(m.残業理由1, m.残業時1, m.残業分1))
                {
                    setErrStatus(eZanRe1, iX - 1, "残業理由が未記入です");
                    return false;
                }

                if (!Utility.chkZangyoRe2(m.残業理由1, m.残業時1, m.残業分1))
                {
                    setErrStatus(eZanH1, iX - 1, "残業時間が未記入です");
                    return false;
                }


                // 休業遅早チェック：2018/10/31
                if (mJiyu[0] == global.JIYU_KYUGYOCHISOU.ToString() ||
                    mJiyu[1] == global.JIYU_KYUGYOCHISOU.ToString() ||
                    mJiyu[2] == global.JIYU_KYUGYOCHISOU.ToString())
                {
                    string sftCode = string.Empty;

                    if (m.シフト通り == global.FLGOFF)
                    {
                        if (m.シフトコード != string.Empty)
                        {
                            // 変更シフトコードあり
                            sftCode = m.シフトコード;
                        }
                        else
                        {
                            // 標準シフトコード
                            sftCode = r.シフトコード.ToString();
                        }
                    }
                    else
                    {
                        // 標準シフトコード
                        sftCode = r.シフトコード.ToString();
                    }
                    
                    sftCode = sftCode.PadLeft(4, '0');

                    // 休業遅早チェック
                    if (!isGreaterSftStartTime(sftCode, m.出勤時, m.出勤分, m.退勤時, m.退勤分))
                    {
                        setErrStatus(eSH, iX - 1, "休業遅早で出勤退勤時刻が正しくありません");
                        return false;
                    }
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

                if (!Utility.chkZangyoRe(m.残業理由2, m.残業時2, m.残業分2))
                {
                    setErrStatus(eZanRe2, iX - 1, "残業理由が未記入です");
                    return false;
                }

                if (!Utility.chkZangyoRe2(m.残業理由2, m.残業時2, m.残業分2))
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

                // 残業と出退勤時刻記入
                if (!errCheckZanShEh(m))
                {
                    setErrStatus(eSH, iX - 1, "残業があるときは出退勤時刻の記入が必要です");
                    return false;
                }

                // 休出日の残業時間チェック
                double uZan = 0;
                if (!errKyushutsuZanTime(_dts, r, m, out uZan))
                {
                    string msg = "残業時間が勤務時間を超えています";
                    msg += "（応援残業：" + (uZan / 60).ToString("##0.0") + "時間）";

                    setErrStatus(eZanH1, iX - 1, msg);
                    return false;
                }

                // 残業時間チェック
                double z = 0;
                double zk = 0;

                // 2017/09/27
                if (!errCheckZan(sdCon, r, m, dts, out z, out zk, kKbn))
                {
                    setErrStatus(eZanH1, iX - 1, "残業時間が正しくありません（計算値：" + z.ToString("##0.0") + "時間, 記入合計：" + zk.ToString("##0.0") + "時間）");
                    return false;
                }

                // 応援チェック
                // 応援チェックありで応援依頼票がないとき
                if (!chkOuenData(r, m, dts))
                {
                    setErrStatus(eChkOuen, iX - 1, "「応援チェックあり」で該当する応援依頼票データが存在しません");
                    return false;
                }

                // 応援チェックなしで応援依頼票があるとき
                if (!chkOuenData2(r, m, dts))
                {
                    setErrStatus(eChkOuen, iX - 1, "「応援チェックなし」で応援依頼票データが存在します");
                    return false;
                }
            }

            return true;
        }

        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     記入開始終了時刻が対象のシフトコードで遅刻早退に該当するか調べる </summary>
        /// <param name="sftCode">
        ///     シフトコード </param>
        /// <param name="sh">
        ///     記入開始時刻・時</param>
        /// <param name="sm">
        ///     記入開始時刻・分</param>
        /// <param name="eh">
        ///     記入終了時刻・時</param>
        /// <param name="em">
        ///     記入終了時刻・分</param>
        ///----------------------------------------------------------------------------------
        private bool isGreaterSftStartTime(string sftCode, string sh, string sm, string eh, string em)
        {
            bool rtn = false;

            string sG = Utility.NulltoStr(sh) + Utility.NulltoStr(sm);

            // 出勤時間空白は対象外とする
            if (sh != string.Empty && sm != string.Empty && eh != string.Empty && em != string.Empty)
            {
                DateTime sDt = DateTime.Now;
                DateTime eDt = DateTime.Now;
                DateTime cDt = DateTime.Now;

                // 有効なシフトコードが存在するとき
                if (sftCode != string.Empty)
                {
                    // 奉行SQLServer接続文字列取得
                    string sc = sqlControl.obcConnectSting.get(_dbName);
                    sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

                    // 勤務体系（シフト）コード情報取得
                    StringBuilder sb = new StringBuilder();
                    sb.Clear();
                    sb.Append("SELECT tbLaborSystem.LaborSystemID,tbLaborSystem.LaborSystemCode,");
                    sb.Append("tbLaborSystem.LaborSystemName,tbLaborTimeSpanRule.StartTime,");
                    sb.Append("tbLaborTimeSpanRule.EndTime, tbLaborSystem.DayChangeTime ");
                    sb.Append("FROM tbLaborSystem inner join tbLaborTimeSpanRule ");
                    sb.Append("on tbLaborSystem.LaborSystemID = tbLaborTimeSpanRule.LaborSystemID ");
                    sb.Append("where tbLaborTimeSpanRule.LaborTimeSpanRuleType = 1 ");
                    sb.Append("and tbLaborSystem.LaborSystemCode = '").Append(sftCode).Append("'");

                    SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

                    bool bn = false;

                    while (dR.Read())
                    {
                        bn = true;
                        sDt = DateTime.Parse(dR["StartTime"].ToString());       // 開始時刻
                        eDt = DateTime.Parse(dR["EndTime"].ToString());         // 終了時刻
                        cDt = DateTime.Parse(dR["DayChangeTime"].ToString());   // 日替時刻
                        break;
                    }

                    dR.Close();
                    sdCon.Close();

                    int sS = sDt.Hour * 100 + sDt.Minute;   // 開始時刻
                    int eS = eDt.Hour * 100 + eDt.Minute;   // 終了時刻
                    int cS = cDt.Hour * 100 + cDt.Minute;   // 日替時刻

                    // 日跨ぎ時間帯のとき
                    if (sS > eS)
                    {
                        eS += 2400;
                    }


                    // 開始時刻
                    int sK = Utility.StrtoInt(Utility.NulltoStr(sh)) * 100 + Utility.StrtoInt(Utility.NulltoStr(sm));

                    // 終了時刻
                    int sE = Utility.StrtoInt(Utility.NulltoStr(eh)) * 100 + Utility.StrtoInt(Utility.NulltoStr(em));

                    // 記入時刻が日替時刻より小さいときは翌日とみなす
                    if (sK < cS)
                    {
                        // 記入開始時刻
                        sK += 2400;
                    }

                    if (sE < cS)
                    {
                        // 記入終了時刻
                        sE += 2400;
                    }

                    if (!bn)
                    {
                        rtn = false;
                    }
                    else if (sS < sK || sE < eS)
                    {
                        // 遅刻もしくは早退とみなされるとき
                        rtn = true;
                    }
                    else
                    {
                        rtn = false;
                    }
                }
            }

            return rtn;
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
        public Boolean errCheckData(sqlControl.DataControl sdCon, DataSet1 dts, DataSet1.過去勤務票ヘッダRow r)
        {
            string sDate;
            DateTime eDate;
            DataSet1 _dts = dts;

            //// 確認チェック
            //if (r.確認 == global.flgOff)
            //{
            //    setErrStatus(eDataCheck, 0, "未確認の出勤簿です");
            //    return false;
            //}

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
            
            // 部署コードチェック
            string dCode = getDepartmentCode(r.部署コード.ToString());
            if (!chkDepartmentCode(dCode, sdCon))
            {
                setErrStatus(eBushoCode, 0, "マスター未登録の部署コードです");
                return false;
            }

            // 勤務体系コード就業奉行登録チェック
            if (!chkSftCode(sdCon, r.シフトコード.ToString()))
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
            List<DataSet1.過去勤務票明細Row> mList = dts.過去勤務票明細.Where(a => a.ヘッダID == r.ID).OrderBy(a => a.ID).ToList();

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

                // 登録済み社員番号マスター検証 : 出勤簿日付で判断 2017/09/28
                if (!chkShainCode(m.社員番号, sdCon, eDate))
                {
                    setErrStatus(eShainNo, iX - 1, "未登録もしくは" + eDate.ToShortDateString() + "現在、在籍していない社員番号です");
                    return false;
                }

                // 同一作業日報内で同じ社員番号が複数記入されているとエラー 2015/04/15
                if (!getSameNumber(mList, m.社員番号))
                {
                    setErrStatus(eShainNo, iX - 1, "同じ社員番号のデータが複数あります");
                    return false;
                }

                // 明細記入チェック
                if (!errCheckRow(m, "過去勤怠データＩ／Ｐ票内容", iX)) return false;
                
                // シフト通りチェック
                if (m.シフト通り == global.FLGON && m.シフトコード != string.Empty)
                {
                    setErrStatus(eChksft, iX - 1, "「シフト通り」で変更シフトコードが記入されています");
                    return false;
                }

                // シフト通りと事由
                if (!errCheckSftJiyu(m))
                {
                    setErrStatus(eChksft, iX - 1, "「シフト通り」で事由が記入されています");
                    return false;
                }

                // シフト通りと残業
                if (!errCheckSftZangyo(m))
                {
                    setErrStatus(eChksft, iX - 1, "「シフト通り」で残業が記入されています");
                    return false;
                }
                
                // シフト通りと休出：2018/02/06
                if (!errCheckSftKyushutsu(r, m))
                {
                    setErrStatus(eChksft, iX - 1, "休日出勤で「シフト通り」は記入できません");
                    return false;
                }

                //// 休出：残業チェック 2018/02/02 コメント化
                //if (!errKyushutsuZangyo(r, m))
                //{
                //    setErrStatus(eZanRe1, iX - 1, "「休日出勤」で残業が記入されていません");
                //    return false;
                //}

                // 休出：事由チェック
                int jNum = 0;
                if (!errKyushutsuJiyu(r, m, out jNum))
                {
                    int eNum = 0;

                    if (jNum == 1)
                    {
                        eNum = eJiyu1;
                    }
                    else if (jNum == 2)
                    {
                        eNum = eJiyu2;
                    }
                    else if (jNum == 3)
                    {
                        eNum = eJiyu3;
                    }

                    setErrStatus(eNum, iX - 1, "「休日出勤」で事由が記入されています");
                    return false;
                }

                // 事由マスター登録チェック
                string[] mJiyu = { m.事由1, m.事由2, m.事由3 };
                string[] eJiyu = { eJiyu1.ToString(), eJiyu2.ToString(), eJiyu3.ToString() };
                int errNum = 0;

                // 事由マスター登録チェック
                OCR.clsJiyuHas jiyu = new OCR.clsJiyuHas(mJiyu);
                if (!jiyu.isHasRows(out errNum, sdCon))
                {
                    setErrStatus(Utility.StrtoInt(eJiyu[errNum]), iX - 1, "マスター未登録の事由です");
                    return false;
                }

                // 前半欠勤または後半欠勤が単独記入されていたらエラー : 2018/02/15
                OCR.clsJiyuHankekkin jiyuHanke = new OCR.clsJiyuHankekkin(mJiyu);
                if (!jiyuHanke.isHankeAnotherDay(out errNum))
                {
                    setErrStatus(Utility.StrtoInt(eJiyu[errNum]), iX - 1, "前半欠勤または後半欠勤の単独記入は出来ません");
                    return false;
                }

                // 取得単位が終日の事由チェック
                OCR.clsJiyuAllDay jiyu2 = new OCR.clsJiyuAllDay(mJiyu);
                if (!jiyu2.isAllDayAnotherDay(out errNum, sdCon))
                {
                    setErrStatus(Utility.StrtoInt(eJiyu[errNum]), iX - 1, "取得単位が「終日」の事由と他の事由は同時に記入出来ません");
                    return false;
                }


                string[] jj = new string[3];
                jj[0] = m.事由1;
                jj[1] = m.事由2;
                jj[2] = m.事由3;

                //if (!Utility.chkJiyu(jj, _dbName))
                //{
                //    setErrStatus(eJiyu1, iX - 1, "取得単位が「終日」の事由と他の事由は同時に記入出来ません");
                //    return false;
                //}

                // 取得単位が終日のときシフト以外記入チェック
                jNum = 0;
                OCR.clsAlldayAnotherData jiyu3 = new OCR.clsAlldayAnotherData(mJiyu);
                if (!jiyu3.isAlldayAnotherData(m, sdCon, out jNum))
                {
                    int eNum = 0;

                    if (jNum == 1)
                    {
                        eNum = eSH;
                    }

                    if (jNum == 2)
                    {
                        eNum = eSM;
                    }

                    if (jNum == 3)
                    {
                        eNum = eEH;
                    }

                    if (jNum == 4)
                    {
                        eNum = eEM;
                    }

                    if (jNum == 5)
                    {
                        eNum = eZanRe1;
                    }
                    
                    if (jNum == 6)
                    {
                        eNum = eZanH1;
                    }

                    if (jNum == 7)
                    {
                        eNum = eZanM1;
                    }

                    if (jNum == 8)
                    {
                        eNum = eZanRe2;
                    }

                    if (jNum == 9)
                    {
                        eNum = eZanH2;
                    }

                    if (jNum == 10)
                    {
                        eNum = eZanM2;
                    }

                    if (jNum == 11)
                    {
                        eNum = eChkOuen;
                    }

                    if (jNum == 12)
                    {
                        eNum = eSftCode;
                    } 

                    setErrStatus(eNum, iX - 1, "取得単位が「終日」の事由で他の項目が記入されています");
                    return false;
                }

                // 取得単位が終日で休出のとき記入チェック
                OCR.clsAllDayOffWork jiyu4 = new OCR.clsAllDayOffWork(mJiyu);
                if (!jiyu4.isAllDayOffWork(r, m, sdCon))
                {
                    setErrStatus(Utility.StrtoInt(eJiyu[0]), iX - 1, "休出のとき取得単位が「終日」の事由は使用できません");
                    return false;
                }

                // 半休事由の重複記入チェック
                int divCnt = 0;
                OCR.clsJiyuDiv jiyu5 = new OCR.clsJiyuDiv(mJiyu);
                if (!jiyu5.isJiyuDiv(sdCon, out divCnt))
                {
                    for (int i = 0; i < mJiyu.Length; i++)
                    {
                        if (mJiyu[i] != string.Empty)
                        {
                            setErrStatus(Utility.StrtoInt(eJiyu[i]), iX - 1, "半休事由の記入が正しくありません");
                            break;
                        }
                    }
                    return false;
                }
                else
                {
                    // 半休のとき（前半、または後半休のみのとき）
                    if (divCnt == 1)
                    {
                        // 半休のとき始業時刻・終業時刻は必須入力
                        if (m.出勤時 == string.Empty && m.出勤分 == string.Empty &&
                            m.退勤時 == string.Empty && m.退勤分 == string.Empty)
                        {
                            setErrStatus(eSH, iX - 1, "半休で出退勤時刻が記入されていません");
                            return false;
                        }
                        else
                        {
                            if (!errCheckTime(m, "出退時間", tanMin1, iX)) return false;
                        }

                        // 半休と出退勤時刻
                        string msg = "";
                        int zenkou = 0;
                        if (!chkHankyuTime(sdCon, r, m, jj, out msg, out zenkou))
                        {
                            if (zenkou == 1)
                            {
                                // 前半休のとき
                                setErrStatus(eSH, iX - 1, msg);
                            }
                            else if (zenkou == 2)
                            {
                                // 後半休のとき
                                setErrStatus(eEH, iX - 1, msg);
                            }

                            return false;
                        }
                    }
                    else if (divCnt == 2)
                    {
                        // 前半、後半組み合わせで終日休暇のときシフト以外記入エラー
                        if (m.出勤時 != string.Empty || m.出勤分 != string.Empty || m.退勤時 != string.Empty || m.退勤分 != string.Empty ||
                            m.残業理由1 != string.Empty || m.残業時1 != string.Empty || m.残業分1 != string.Empty ||
                            m.残業理由2 != string.Empty || m.残業時2 != string.Empty || m.残業分2 != string.Empty ||
                            m.応援 == global.FLGON)
                        {
                            setErrStatus(eJiyu1, iX - 1, "前半休＋後半休で終日休みのため他の項目の記入は不要です");
                            return false;
                        }
                    }
                }
                
                // 前半休＋後半休で終日休みのときはチェック不要 2017/11/13
                if (divCnt != 2)
                {
                    // シフト通りではないとき : 2017/07/26
                    OCR.clsNotAlldayShift eShift = new OCR.clsNotAlldayShift(mJiyu);
                    if (!eShift.isNotAlldayShift(m, sdCon, out jNum))
                    {
                        setErrStatus(eChksft, iX - 1, "「シフト通りではない」とき変更シフトコードまたは勤務時間の記入が必要です");
                        return false;
                    }
                }

                //// 変更シフトコード
                //if (m.シフト通り == string.Empty && m.シフトコード == string.Empty)
                //{
                //    setErrStatus(eSftCode, iX - 1, "変更シフトコードが未記入です");
                //    return false;
                //}

                // 該当社員の雇用区分を取得：2017/09/27
                int kKbn = 0;
                getEmployee(sdCon, m.社員番号, ref kKbn);

                // 「土曜特休」はパート社員、契約社員、アルバイトのみ記入可 2017/11/10
                if (kKbn != KOYOU_PART && kKbn != KOYOU_PART8 && kKbn != KOYOU_PART9)
                {
                    for (int i = 0; i < mJiyu.Length; i++)
                    {
                        if (mJiyu[i] == global.SFT_DOTOKKYU.ToString())
                        {
                            setErrStatus(Utility.StrtoInt(eJiyu[i]), iX - 1, "パート以外は「土曜特休」は記入できません");
                            return false;
                        }
                    }
                }

                // 雇用区分「１０」派遣社員が使用可能な事由は「１０：通常欠勤」のみとする 2017/11/21
                // 「１１：休業欠勤」も対象とする 2018/10/30
                // 「２０：遅刻早退」も対象とする 2020/04/01
                if (kKbn == KOYOU_HAKEN)
                {
                    for (int i = 0; i < mJiyu.Length; i++)
                    {
                        if (mJiyu[i] != string.Empty && 
                            mJiyu[i] != global.SFT_TSUJYOKEKKIN.ToString() &&
                            mJiyu[i] != global.JIYU_KYUGYOKEKKIN.ToString() &&
                            mJiyu[i] != global.JIYU_KYUGYOCHISOU.ToString())
                        {
                            setErrStatus(Utility.StrtoInt(eJiyu[i]), iX - 1, "派遣社員は「10：通常欠勤」「11：休業欠勤」「20：休業遅早」以外は記入できません");
                            return false;
                        }
                    }
                }


                // 勤務体系（シフト）コード奉行登録チェック
                if (m.シフトコード != string.Empty)
                {
                    if (!chkSftCode(sdCon, m.シフトコード))
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
                    
                    // 2018/02/03 出退勤時刻が無記入のときエラー
                    if (m.出勤時 == string.Empty && m.出勤分 == string.Empty &&
                        m.退勤時 == string.Empty && m.退勤分 == string.Empty)
                    {
                        setErrStatus(eSH, iX - 1, "シフト変更で出退勤時刻が無記入です");
                        return false;
                    }

                    /* 休日以外の日に変更シフト欄に[31]休出・休憩なしが記入されているときエラーとしない為、
                     * コメント化 2018/05/28 */
                    // 休日以外の日に変更シフト欄に[31]休出が記入されているとき : 2017/10/19
                    //if (sHol != global.FLGON && 
                    //    (Utility.StrtoInt(m.シフトコード) == global.SFT_KYUSHUTSU || 
                    //     Utility.StrtoInt(m.シフトコード) == global.SFT_KYUKEI_KYUSHUTSU))
                    //{
                    //    // パートタイマー以外はエラー : 2017/10/19
                    //    if (kKbn != KOYOU_PART && kKbn != KOYOU_PART8 && kKbn != KOYOU_PART9)
                    //    {
                    //        setErrStatus(eSftCode, iX - 1, "休日以外に休出が記入されています");
                    //        return false;
                    //    }
                    //    else
                    //    {
                    //        // パートタイマーでも土曜以外は土曜特休とみなさない
                    //        if (eDate.ToString("ddd") != "土")
                    //        {
                    //            setErrStatus(eSftCode, iX - 1, "休出は記入できません");
                    //            return false;
                    //        }
                    //    }
                    //}
                }

                // 始業時刻・終業時刻チェック : 出退勤は30分単位 2017/08/30
                //if (!errCheckTime(m, "出退時間", tanMin1, iX)) return false;
                if (!errCheckTime(m, "出退時間", tanMin30, iX)) return false;

                // 残業理由
                if (!Utility.chkZangyoRe(m.残業理由1, m.残業時1, m.残業分1))
                {
                    setErrStatus(eZanRe1, iX - 1, "残業理由が未記入です");
                    return false;
                }

                if (!Utility.chkZangyoRe2(m.残業理由1, m.残業時1, m.残業分1))
                {
                    setErrStatus(eZanH1, iX - 1, "残業時間が未記入です");
                    return false;
                }


                // 休業遅早チェック：2018/10/31
                if (mJiyu[0] == global.JIYU_KYUGYOCHISOU.ToString() ||
                    mJiyu[1] == global.JIYU_KYUGYOCHISOU.ToString() ||
                    mJiyu[2] == global.JIYU_KYUGYOCHISOU.ToString())
                {
                    string sftCode = string.Empty;

                    if (m.シフト通り == global.FLGOFF)
                    {
                        if (m.シフトコード != string.Empty)
                        {
                            // 変更シフトコードあり
                            sftCode = m.シフトコード;
                        }
                        else
                        {
                            // 標準シフトコード
                            sftCode = r.シフトコード.ToString();
                        }
                    }
                    else
                    {
                        // 標準シフトコード
                        sftCode = r.シフトコード.ToString();
                    }

                    sftCode = sftCode.PadLeft(4, '0');

                    // 休業遅早チェック
                    if (!isGreaterSftStartTime(sftCode, m.出勤時, m.出勤分, m.退勤時, m.退勤分))
                    {
                        setErrStatus(eSH, iX - 1, "休業遅早で出勤退勤時刻が正しくありません");
                        return false;
                    }
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

                if (!Utility.chkZangyoRe(m.残業理由2, m.残業時2, m.残業分2))
                {
                    setErrStatus(eZanRe2, iX - 1, "残業理由が未記入です");
                    return false;
                }

                if (!Utility.chkZangyoRe2(m.残業理由2, m.残業時2, m.残業分2))
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
                
                // 残業と出退勤時刻記入
                if (!errCheckZanShEh(m))
                {
                    setErrStatus(eSH, iX - 1, "残業があるときは出退勤時刻の記入が必要です");
                    return false;
                }
                
                // 休出日の残業時間チェック
                double uZan = 0;
                if (!errKyushutsuZanTime(_dts, r, m, out uZan))
                {
                    string msg = "残業時間が勤務時間を超えています";
                    msg += "（応援残業：" + (uZan / 60).ToString("##0.0") + "時間）";

                    setErrStatus(eZanH1, iX - 1, msg);
                    return false;
                }

                // 残業時間チェック
                double z = 0;
                double zk = 0;

                // 2017/09/27
                if (!errCheckZan(sdCon, r, m, dts, out z, out zk, kKbn))
                {
                    setErrStatus(eZanH1, iX - 1, "残業時間が正しくありません（計算値：" + z.ToString("##0.0") + "時間, 記入合計：" + zk.ToString("##0.0") + "時間）");
                    return false;
                }

                // 応援チェック
                // 応援チェックありで応援依頼票がないとき
                if (!chkOuenData(r, m, dts))
                {
                    setErrStatus(eChkOuen, iX - 1, "「応援チェックあり」で該当する過去応援依頼票データが存在しません");
                    return false;
                }

                // 応援チェックなしで応援依頼票があるとき
                if (!chkOuenData2(r, m, dts))
                {
                    setErrStatus(eChkOuen, iX - 1, "「応援チェックなし」で過去応援依頼票データが存在します");
                    return false;
                }
            }

            return true;
        }

        //private bool chkJiyuHas(DataSet1.勤務票明細Row m, string mJiyu)
        //{
        //    OCR.clsJiyuHas jiyu = new OCR.clsJiyuHas(mJiyu, _dbName);

        //    // マスター登録チェック
        //    if (!jiyu.isHasRows())
        //    {
        //        return false;
        //    }
        //    else
        //    {
        //        return true;
        //    }
        //}


        ///------------------------------------------------------------------------------
        /// <summary>
        ///     半休と出退勤時刻チェック </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControlオブジェクト</param>
        /// <param name="r">
        ///     DataSet1.勤務票ヘッダRow</param>
        /// <param name="m">
        ///     DataSet1.勤務票明細Row</param>
        /// <param name="jj">
        ///     事由配列</param>
        /// <param name="msg">
        ///     エラーメッセージ</param>
        /// <param name="zenkou">
        ///     前半(1)・後半(2)</param>
        /// <returns>
        ///     true:エラーなし、false:エラー有り</returns>
        ///------------------------------------------------------------------------------
        private bool chkHankyuTime(sqlControl.DataControl sdCon, DataSet1.勤務票ヘッダRow r, DataSet1.勤務票明細Row m, string[] jj, out string msg, out int zenkou)
        {
            //
            msg = string.Empty;
            zenkou = 0;
            string sftCode = string.Empty;

            // 開始終了時間
            if (m.シフトコード != string.Empty)
            {
                sftCode = m.シフトコード;
            }
            else
            {
                sftCode = r.シフトコード.ToString();
            }

            //// 奉行SQLServer接続文字列取得
            //string sc = sqlControl.obcConnectSting.get(_dbName);
            //sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            // 勤務体系（シフト）取得
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select tbLaborSystem.LaborSystemCode, tbLaborSystem.LatterHalfStartTime, tbLaborSystem.FirstHalfEndTime,");
            sb.Append("tbLaborSystem.DayChangeTime,tbLaborTimeSpanRule.StartTime,tbLaborTimeSpanRule.EndTime ");
            sb.Append("from tbLaborSystem inner join tbLaborTimeSpanRule ");
            sb.Append("on tbLaborSystem.LaborSystemID = tbLaborTimeSpanRule.LaborSystemID ");
            sb.Append("where tbLaborSystem.LaborSystemCode = '" + sftCode.ToString().PadLeft(4, '0') + "' and ");
            sb.Append("tbLaborTimeSpanRule.LaborTimeItemID = 1 ");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            DateTime dtStart = DateTime.Now;        // 開始時刻
            DateTime dtEnd = DateTime.Now;          // 終了時刻
            DateTime dtLatterStart = DateTime.Now;  // 後半開始時刻
            DateTime dtFirstEnd = DateTime.Now;     // 前半終了時刻
            DateTime dtChange = DateTime.Now;       // 日替わり時刻
            bool dm = false;

            while (dR.Read())
            {
                dm = true;
                dtStart = (DateTime)dR["StartTime"];
                dtEnd = (DateTime)dR["EndTime"];
                dtLatterStart = (DateTime)dR["LatterHalfStartTime"];
                dtFirstEnd = (DateTime)dR["FirstHalfEndTime"];
                dtChange = (DateTime)dR["DayChangeTime"];
                break;
            }

            dR.Close();

            string sKbn = string.Empty;

            if (dm)
            {
                for (int i = 0; i < 3; i++)
                {
                    if (jj[i] == string.Empty)
                    {
                        continue;
                    }

                    // 事由・取得区分（前半・後半）取得
                    sb.Clear();
                    sb.Append("select LaborReasonCode,AcquireUnit,AcquireDivision from tbLaborReason ");
                    sb.Append("where IsValid = 1 and LaborReasonCode = '" + jj[i].PadLeft(2, '0') + "'");

                    dR = sdCon.free_dsReader(sb.ToString());

                    while (dR.Read())
                    {
                        // 取得単位
                        if (Utility.NulltoStr(dR["AcquireUnit"]) == global.FLGON)
                        {
                            // 半日「1」
                            dm = true;  // 終日あり
                            sKbn = Utility.NulltoStr(dR["AcquireDivision"]); // 取得区分
                        }

                        break;
                    }

                    dR.Close();
                    break;
                }

                //sdCon.Close();

                int lTm = 0;
                int sTm = 0;
                int nextTime = dtChange.Hour * 100 + dtChange.Minute;

                if (sKbn == global.FLGOFF)
                {
                    sTm = Utility.StrtoInt(m.出勤時) * 100 + Utility.StrtoInt(m.出勤分);

                    // 勤務開始時刻が日替わり時刻以前のとき翌日とみなす
                    if (sTm < nextTime)
                    {
                        sTm += 2400;
                    }

                    // 前半終了時刻 2017/09/19
                    if (dtStart.Day < dtFirstEnd.Day)
                    {
                        lTm = (dtFirstEnd.Hour + 24) * 100 + dtFirstEnd.Minute;
                    }
                    else
                    {
                        lTm = dtFirstEnd.Hour * 100 + dtFirstEnd.Minute;
                    }

                    // 前半休のとき：出勤時刻が前半終了時刻以前のときエラー
                    if (sTm < lTm)
                    {
                        zenkou = 1;
                        msg = "前半休で勤務時間が正しくありません。（前半：" +
                            dtStart.Hour.ToString().PadLeft(2, '0') + ":" + dtStart.Minute.ToString().PadLeft(2, '0') + "～" +
                            dtFirstEnd.Hour.ToString().PadLeft(2, '0') + ":" + dtFirstEnd.Minute.ToString().PadLeft(2, '0') +
                            "、後半：" +
                            dtLatterStart.Hour.ToString().PadLeft(2, '0') + ":" + dtLatterStart.Minute.ToString().PadLeft(2, '0') +
                            "～" +
                            dtEnd.Hour.ToString().PadLeft(2, '0') + ":" + dtEnd.Minute.ToString().PadLeft(2, '0') + "）";

                        return false;
                    }
                }
                else
                {
                    sTm = Utility.StrtoInt(m.退勤時) * 100 + Utility.StrtoInt(m.退勤分);

                    // 退勤時刻が日替わり時刻以前のとき翌日とみなす
                    if (sTm < nextTime)
                    {
                        sTm += 2400;
                    }

                    // 後半開始時刻 2017/09/19
                    if (dtStart.Day < dtLatterStart.Day)
                    {
                        // 後半開始が翌日のとき
                        lTm = (dtLatterStart.Hour + 24) * 100 + dtLatterStart.Minute;
                    }
                    else
                    {
                        lTm = dtLatterStart.Hour * 100 + dtLatterStart.Minute;
                    }

                    // 後半休のとき：退勤時刻が後半開始時刻以降のときエラー
                    if (sTm > lTm)
                    {
                        zenkou = 2;
                        msg = "後半休で勤務時間が正しくありません。（前半：" +
                            dtStart.Hour.ToString().PadLeft(2, '0') + ":" + dtStart.Minute.ToString().PadLeft(2, '0') + "～" +
                            dtFirstEnd.Hour.ToString().PadLeft(2, '0') + ":" + dtFirstEnd.Minute.ToString().PadLeft(2, '0') +
                            "、後半：" +
                            dtLatterStart.Hour.ToString().PadLeft(2, '0') + ":" + dtLatterStart.Minute.ToString().PadLeft(2, '0') +
                            "～" +
                            dtEnd.Hour.ToString().PadLeft(2, '0') + ":" + dtEnd.Minute.ToString().PadLeft(2, '0') + "）";

                        return false;
                    }
                }

                return true;
            }
            else
            {
                return true;
            }
        }


        ///------------------------------------------------------------------------------
        /// <summary>
        ///     半休と出退勤時刻チェック </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControlオブジェクト</param>
        /// <param name="r">
        ///     DataSet1.過去勤務票ヘッダRow</param>
        /// <param name="m">
        ///     DataSet1.過去勤務票明細Row</param>
        /// <param name="jj">
        ///     事由配列</param>
        /// <param name="msg">
        ///     エラーメッセージ</param>
        /// <param name="zenkou">
        ///     前半(1)・後半(2)</param>
        /// <returns>
        ///     true:エラーなし、false:エラー有り</returns>
        ///------------------------------------------------------------------------------
        private bool chkHankyuTime(sqlControl.DataControl sdCon, DataSet1.過去勤務票ヘッダRow r, DataSet1.過去勤務票明細Row m, string[] jj, out string msg, out int zenkou)
        {
            //
            msg = string.Empty;
            zenkou = 0;
            string sftCode = string.Empty;

            // 開始終了時間
            if (m.シフトコード != string.Empty)
            {
                sftCode = m.シフトコード;
            }
            else
            {
                sftCode = r.シフトコード.ToString();
            }

            //// 奉行SQLServer接続文字列取得
            //string sc = sqlControl.obcConnectSting.get(_dbName);
            //sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            // 勤務体系（シフト）取得
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select tbLaborSystem.LaborSystemCode, tbLaborSystem.LatterHalfStartTime, tbLaborSystem.FirstHalfEndTime,");
            sb.Append("tbLaborSystem.DayChangeTime,tbLaborTimeSpanRule.StartTime,tbLaborTimeSpanRule.EndTime ");
            sb.Append("from tbLaborSystem inner join tbLaborTimeSpanRule ");
            sb.Append("on tbLaborSystem.LaborSystemID = tbLaborTimeSpanRule.LaborSystemID ");
            sb.Append("where tbLaborSystem.LaborSystemCode = '" + sftCode.ToString().PadLeft(4, '0') + "' and ");
            sb.Append("tbLaborTimeSpanRule.LaborTimeItemID = 1 ");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            DateTime dtStart = DateTime.Now;        // 開始時刻
            DateTime dtEnd = DateTime.Now;          // 終了時刻
            DateTime dtLatterStart = DateTime.Now;  // 後半開始時刻
            DateTime dtFirstEnd = DateTime.Now;     // 前半終了時刻
            DateTime dtChange = DateTime.Now;       // 日替わり時刻
            bool dm = false;

            while (dR.Read())
            {
                dm = true;
                dtStart = (DateTime)dR["StartTime"];
                dtEnd = (DateTime)dR["EndTime"];
                dtLatterStart = (DateTime)dR["LatterHalfStartTime"];
                dtFirstEnd = (DateTime)dR["FirstHalfEndTime"];
                dtChange = (DateTime)dR["DayChangeTime"];
                break;
            }

            dR.Close();

            string sKbn = string.Empty;

            if (dm)
            {
                for (int i = 0; i < 3; i++)
                {
                    if (jj[i] == string.Empty)
                    {
                        continue;
                    }

                    // 事由・取得区分（前半・後半）取得
                    sb.Clear();
                    sb.Append("select LaborReasonCode,AcquireUnit,AcquireDivision from tbLaborReason ");
                    sb.Append("where IsValid = 1 and LaborReasonCode = '" + jj[i].PadLeft(2, '0') + "'");

                    dR = sdCon.free_dsReader(sb.ToString());

                    while (dR.Read())
                    {
                        // 取得単位
                        if (Utility.NulltoStr(dR["AcquireUnit"]) == global.FLGON)
                        {
                            // 半日「1」
                            dm = true;  // 終日あり
                            sKbn = Utility.NulltoStr(dR["AcquireDivision"]); // 取得区分
                        }

                        break;
                    }

                    dR.Close();
                    break;
                }

                //sdCon.Close();

                int lTm = 0;
                int sTm = 0;
                int nextTime = dtChange.Hour * 100 + dtChange.Minute;

                if (sKbn == global.FLGOFF)
                {
                    sTm = Utility.StrtoInt(m.出勤時) * 100 + Utility.StrtoInt(m.出勤分);

                    // 勤務開始時刻が日替わり時刻以前のとき翌日とみなす
                    if (sTm < nextTime)
                    {
                        sTm += 2400;
                    }

                    // 前半終了時刻 2017/09/19
                    if (dtStart.Day < dtFirstEnd.Day)
                    {
                        lTm = (dtFirstEnd.Hour + 24) * 100 + dtFirstEnd.Minute;
                    }
                    else
                    {
                        lTm = dtFirstEnd.Hour * 100 + dtFirstEnd.Minute;
                    }

                    // 前半休のとき：出勤時刻が前半終了時刻以前のときエラー
                    if (sTm < lTm)
                    {
                        zenkou = 1;
                        msg = "前半休で勤務時間が正しくありません。（前半：" +
                            dtStart.Hour.ToString().PadLeft(2, '0') + ":" + dtStart.Minute.ToString().PadLeft(2, '0') + "～" +
                            dtFirstEnd.Hour.ToString().PadLeft(2, '0') + ":" + dtFirstEnd.Minute.ToString().PadLeft(2, '0') +
                            "、後半：" +
                            dtLatterStart.Hour.ToString().PadLeft(2, '0') + ":" + dtLatterStart.Minute.ToString().PadLeft(2, '0') +
                            "～" +
                            dtEnd.Hour.ToString().PadLeft(2, '0') + ":" + dtEnd.Minute.ToString().PadLeft(2, '0') + "）";

                        return false;
                    }
                }
                else
                {
                    sTm = Utility.StrtoInt(m.退勤時) * 100 + Utility.StrtoInt(m.退勤分);

                    // 退勤時刻が日替わり時刻以前のとき翌日とみなす
                    if (sTm < nextTime)
                    {
                        sTm += 2400;
                    }

                    // 後半開始時刻 2017/09/19
                    if (dtStart.Day < dtLatterStart.Day)
                    {
                        // 後半開始が翌日のとき
                        lTm = (dtLatterStart.Hour + 24) * 100 + dtLatterStart.Minute;
                    }
                    else
                    {
                        lTm = dtLatterStart.Hour * 100 + dtLatterStart.Minute;
                    }

                    // 後半休のとき：退勤時刻が後半開始時刻以降のときエラー
                    if (sTm > lTm)
                    {
                        zenkou = 2;
                        msg = "後半休で勤務時間が正しくありません。（前半：" +
                            dtStart.Hour.ToString().PadLeft(2, '0') + ":" + dtStart.Minute.ToString().PadLeft(2, '0') + "～" +
                            dtFirstEnd.Hour.ToString().PadLeft(2, '0') + ":" + dtFirstEnd.Minute.ToString().PadLeft(2, '0') +
                            "、後半：" +
                            dtLatterStart.Hour.ToString().PadLeft(2, '0') + ":" + dtLatterStart.Minute.ToString().PadLeft(2, '0') +
                            "～" +
                            dtEnd.Hour.ToString().PadLeft(2, '0') + ":" + dtEnd.Minute.ToString().PadLeft(2, '0') + "）";

                        return false;
                    }
                }

                return true;
            }
            else
            {
                return true;
            }
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
                if (!Utility.checkHourSpan(m.出勤時))
                {
                    setErrStatus(eSH, iX - 1, tittle + "が正しくありません");
                    return false;
                }

                if (!Utility.checkMinSpan(m.出勤分, Tani))
                {
                    setErrStatus(eSM, iX - 1, tittle + "が正しくありません");
                    return false;
                }

                if (!Utility.checkHourSpan(m.退勤時))
                {
                    setErrStatus(eEH, iX - 1, tittle + "が正しくありません");
                    return false;
                }

                if (!Utility.checkMinSpan(m.退勤分, Tani))
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

        private bool errCheckTime(DataSet1.過去勤務票明細Row m, string tittle, int Tani, int iX)
        {
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
                if (!Utility.checkHourSpan(m.出勤時))
                {
                    setErrStatus(eSH, iX - 1, tittle + "が正しくありません");
                    return false;
                }

                if (!Utility.checkMinSpan(m.出勤分, Tani))
                {
                    setErrStatus(eSM, iX - 1, tittle + "が正しくありません");
                    return false;
                }

                if (!Utility.checkHourSpan(m.退勤時))
                {
                    setErrStatus(eEH, iX - 1, tittle + "が正しくありません");
                    return false;
                }

                if (!Utility.checkMinSpan(m.退勤分, Tani))
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
        
        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     対象のシフトコードの残業時間帯を取得する </summary>
        /// <param name="r">
        ///     MultiRowの行インデックス </param>
        ///----------------------------------------------------------------------------------
        private int getSftZanTimeXXX(DataSet1.勤務票ヘッダRow r, DataSet1.勤務票明細Row m)
        {
            double zanTime = 0;

            string sG = m.出勤時 + m.出勤分;

            // 出勤時間空白は対象外とする
            if (sG != string.Empty)
            {
                // 勤務開始時刻
                DateTime wSTM = DateTime.Parse(m.出勤時 + ":" + m.出勤分);

                // 勤務終了時刻
                DateTime wETM = DateTime.Now;

                int st = Utility.StrtoInt(m.出勤時) * 100 + Utility.StrtoInt(m.出勤分);
                int et = Utility.StrtoInt(m.退勤時) * 100 + Utility.StrtoInt(m.退勤分);

                if (st > et)
                {
                    // 退勤時刻が翌日のとき
                    wETM = DateTime.Parse(m.退勤時 + ":" + m.退勤分).AddDays(1);
                }
                else
                {
                    // 当日のとき
                    wETM = DateTime.Parse(m.退勤時 + ":" + m.退勤分);
                }

                // 対象のシフトコード取得する
                string sftCode = string.Empty;
                DateTime zSTM = DateTime.Now;
                DateTime zETM = DateTime.Now;

                // 休憩時刻
                DateTime rSTM = DateTime.Now;
                DateTime rETM = DateTime.Now;

                if (m.シフトコード != string.Empty)
                {
                    // 変更シフトコードあり
                    sftCode = m.シフトコード.PadLeft(4, '0');
                }
                else if (r.シフトコード.ToString() != string.Empty)
                {
                    // 標準シフトコード
                    sftCode = r.シフトコード.ToString().PadLeft(4, '0');
                }

                // 有効なシフトコードが存在するとき
                if (sftCode != string.Empty)
                {
                    // 奉行SQLServer接続文字列取得
                    string sc = sqlControl.obcConnectSting.get(_dbName);
                    sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

                    // 指定した勤務体系（シフト）コードの残業開始、終了時刻を取得
                    StringBuilder sb = new StringBuilder();
                    sb.Clear();
                    sb.Append("SELECT tbLaborSystem.LaborSystemID,tbLaborSystem.LaborSystemCode,");
                    sb.Append("tbLaborSystem.LaborSystemName,tbLaborTimeSpanRule.StartTime,");
                    sb.Append("tbLaborTimeSpanRule.EndTime,");
                    sb.Append("tbRestTimeSpanRule.StartTime as restStartTime,");
                    sb.Append("tbRestTimeSpanRule.EndTime as restEndTime ");
                    sb.Append("FROM tbLaborSystem inner join tbLaborTimeSpanRule ");
                    sb.Append("on tbLaborSystem.LaborSystemID = tbLaborTimeSpanRule.LaborSystemID ");
                    sb.Append("left join tbRestTimeSpanRule ");
                    sb.Append("on tbLaborSystem.LaborSystemID = tbRestTimeSpanRule.LaborSystemID ");
                    sb.Append("where tbLaborTimeSpanRule.LaborTimeSpanRuleType = 4 ");
                    sb.Append("and tbLaborSystem.LaborSystemCode = '").Append(sftCode).Append("' and ");
                    sb.Append("tbRestTimeSpanRule.TimeOrder = 0");

                    SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

                    while (dR.Read())
                    {
                        zSTM = DateTime.Parse(dR["StartTime"].ToString());
                        zETM = DateTime.Parse(dR["EndTime"].ToString());

                        // 今日の日付に変換
                        zSTM = DateTime.Parse(zSTM.Hour.ToString() + ":" + zSTM.Minute.ToString() + ":" + zSTM.Second.ToString());

                        if ((zSTM.Hour * 100 + zSTM.Minute) > (zETM.Hour * 100 + zETM.Minute))
                        {
                            // 翌日のとき
                            zETM = DateTime.Parse(zETM.Hour.ToString() + ":" + zETM.Minute.ToString() + ":" + zETM.Second.ToString()).AddDays(1);
                        }
                        else
                        {
                            // 当日のとき
                            zETM = DateTime.Parse(zETM.Hour.ToString() + ":" + zETM.Minute.ToString() + ":" + zETM.Second.ToString());
                        }
                        
                        // 退勤時刻が残業開始時刻以降のとき、残業計算を行う
                        if (wETM > zSTM)
                        {
                            DateTime dS = DateTime.Now;
                            DateTime dE = DateTime.Now;

                            // 残業開始時間の決定
                            if (wSTM < zSTM)
                            {
                                dS = zSTM;
                            }
                            else
                            {
                                dS = wSTM;
                            }

                            // 残業終了時間を決定
                            if (wETM < zETM)
                            {
                                dE = wETM;
                            }
                            else
                            {
                                dE = zETM;
                            }

                            // 残業時間（分単位）
                            zanTime += Utility.GetTimeSpan(dS, dE).TotalMinutes;

                            // 休憩時間を取得
                            rSTM = DateTime.Parse(dR["restStartTime"].ToString());
                            rETM = DateTime.Parse(dR["restEndTime"].ToString());

                            // 今日の日付に変換
                            rSTM = DateTime.Parse(rSTM.Hour.ToString() + ":" + rSTM.Minute.ToString() + ":" + rSTM.Second.ToString());

                            if ((rSTM.Hour * 100 + rSTM.Minute) > (rETM.Hour * 100 + rETM.Minute))
                            {
                                // 翌日のとき
                                rETM = DateTime.Parse(rETM.Hour.ToString() + ":" + rETM.Minute.ToString() + ":" + rETM.Second.ToString()).AddDays(1);
                            }
                            else
                            {
                                // 当日のとき
                                rETM = DateTime.Parse(rETM.Hour.ToString() + ":" + rETM.Minute.ToString() + ":" + rETM.Second.ToString());
                            }
       
                            DateTime rS = DateTime.Now;
                            DateTime rE = DateTime.Now;
                            bool restType = false;

                            // 休憩開始時間の決定
                            if (dS < rSTM && dE > rETM)
                            {
                                rS = rSTM;
                                rE = rETM;
                                restType = true;
                            }

                            if (dS >= rSTM && dE >= rETM)
                            {
                                rS = dS;
                                rE = rETM;
                                restType = true;
                            }

                            if (dS <= rSTM && dE <= rETM)
                            {
                                rS = rSTM;
                                rE = dE;
                                restType = true;
                            }

                            if (dS > rSTM && dE < rETM)
                            {
                                rS = dS;
                                rE = dE;
                                restType = true;
                            }

                            // 休憩時間を差し引く
                            if (restType)
                            {
                                zanTime -= Utility.GetTimeSpan(rS, rE).TotalMinutes;
                            }
                        }
                    }

                    dR.Close();
                    sdCon.Close();
                }
            }

            return (int)zanTime;
        }

        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     対象のシフトコードの残業時間を取得する </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl</param>
        /// <param name="m">
        ///     DataSet1.勤務票ヘッダRow</param>
        /// <param name="r">
        ///     DataSet1.勤務票明細Row </param>
        ///----------------------------------------------------------------------------------
        private int getSftZanTime(sqlControl.DataControl sdCon, DataSet1.勤務票ヘッダRow r, DataSet1.勤務票明細Row m)
        {
            double zanTime = 0;

            // 出勤時間空白は対象外とする
            if (m.出勤時 == string.Empty && m.出勤分 == string.Empty)
            {
                return (int)zanTime;
            }
            
            bool bn = false;
            DateTime dtStart = DateTime.Now;        // 開始時刻
            DateTime dtEnd = DateTime.Now;          // 終了時刻
            DateTime dtChange = DateTime.Now;       // 日替わり時刻

            // 対象のシフトコード取得する
            string sftCode = string.Empty;

            if (m.シフトコード != string.Empty)
            {
                // 変更シフトコードあり
                sftCode = m.シフトコード.PadLeft(4, '0');
            }
            else if (r.シフトコード.ToString() != string.Empty)
            {
                // 標準シフトコード
                sftCode = r.シフトコード.ToString().PadLeft(4, '0');
            }

            // 有効なシフトコードが存在しないとき
            if (sftCode == string.Empty)
            {
                return (int)zanTime;
            }

            //// 奉行SQLServer接続文字列取得
            //string sc = sqlControl.obcConnectSting.get(_dbName);
            //sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);
            
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select tbLaborSystem.LaborSystemCode, LaborSystemName, tbLaborSystem.LatterHalfStartTime, tbLaborSystem.FirstHalfEndTime,");
            sb.Append("tbLaborSystem.DayChangeTime,a.StartTime,a.EndTime ");
            sb.Append("from tbLaborSystem left join ");
            sb.Append("(select * from tbLaborTimeSpanRule where LaborTimeItemID = 1) as a ");
            sb.Append("on tbLaborSystem.LaborSystemID = a.LaborSystemID ");
            sb.Append("where tbLaborSystem.LaborSystemCode = '").Append(sftCode).Append("'");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                if (!(dR["StartTime"] is DBNull))
                {
                    bn = true;
                    dtStart = (DateTime)dR["StartTime"];
                }

                if (!(dR["EndTime"] is DBNull))
                {
                    bn = true;
                    dtEnd = (DateTime)dR["EndTime"];
                    dtChange = (DateTime)dR["DayChangeTime"];
                }
            }

            dR.Close();

            //sdCon.Close();

            // 開始時刻、終了時刻がないときは戻る
            if (!bn)
            {
                return (int)zanTime;
            }
            else
            {
                // 今日の日付に変換
                dtStart = DateTime.Parse(dtStart.Hour.ToString() + ":" + dtStart.Minute.ToString() + ":0");
                dtEnd = DateTime.Parse(dtEnd.Hour.ToString() + ":" + dtEnd.Minute.ToString() + ":0");
            }

            // 勤務体系の終了時刻が翌日のとき（夜勤など）
            if (dtStart > dtEnd)
            {
                // 日付を翌日にする
                dtEnd = dtEnd.AddDays(1);   
            }

            // 出勤時刻と退勤時刻
            DateTime wStartTm = DateTime.Parse(m.出勤時 + ":" + m.出勤分 + ":0");
            DateTime wEndTm = DateTime.Parse(m.退勤時 + ":" + m.退勤分 + ":0");

            // 日替わり時刻 : 2017/09/19
            int nextTime = dtChange.Hour * 100 + dtChange.Minute;

            // 2017/09/19
            int sTm = Utility.StrtoInt(m.出勤時) * 100 + Utility.StrtoInt(m.出勤分);

            // 勤務開始時刻が日替わり時刻以前のとき翌日とみなす : 2017/09/19
            if (sTm < nextTime)
            {
                wStartTm = wStartTm.AddDays(1);  // 日付を翌日にする
            }

            // 2017/09/19
            int eTm = Utility.StrtoInt(m.退勤時) * 100 + Utility.StrtoInt(m.退勤分);

            // 勤務終了時刻が日替わり時刻以前のとき、または退勤が翌日のとき翌日とみなす : 2017/09/19
            if (wStartTm > wEndTm || eTm < nextTime)
            {
                wEndTm = wEndTm.AddDays(1);  // 日付を翌日にする
            }

            //// 退勤が翌日のとき
            //if (wStartTm > wEndTm)
            //{
            //    wEndTm = wEndTm.AddDays(1);  // 日付を翌日にする
            //}

            // 退勤時刻が勤務終了時刻以降のとき超過時間計算を行う
            if (wEndTm > dtEnd)
            {
                if (wStartTm < dtEnd)
                {
                    /* 出勤時刻が勤務体系の終了時刻以前のとき
                     * 例）勤務体系8:00-17:00で勤務時間が8:00-21:00等 */
                    zanTime += Utility.GetTimeSpan(dtEnd, wEndTm).TotalMinutes;
                }
                else
                {
                    /* 出勤時刻が勤務体系の終了時刻以降のとき
                     * 例）勤務体系8:00-17:00で勤務時間が18:00-23:00等 */
                    zanTime += Utility.GetTimeSpan(wStartTm, wEndTm).TotalMinutes;
                }
            }

            // 出勤時刻が勤務開始時刻以前のとき早出時間計算を行う
            if (wStartTm < dtStart)
            {
                if (dtStart < wEndTm)
                {
                    /* 退勤時刻時刻が勤務体系の開始時刻以降のとき
                     * 例）勤務体系8:00-17:00で勤務時間が6:00-17:00等 */
                    zanTime += Utility.GetTimeSpan(wStartTm, dtStart).TotalMinutes;
                }
                else
                {
                    /* 退勤時刻時刻が勤務体系の開始時刻以前のとき
                     * 例）勤務体系8:00-17:00で勤務時間が6:00-7:00等 */
                    zanTime += Utility.GetTimeSpan(wStartTm, wEndTm).TotalMinutes;
                }
            }

            return (int)zanTime;
        }

        private int getSftZanTime(sqlControl.DataControl sdCon, DataSet1.過去勤務票ヘッダRow r, DataSet1.過去勤務票明細Row m)
        {
            double zanTime = 0;

            // 出勤時間空白は対象外とする
            if (m.出勤時 == string.Empty && m.出勤分 == string.Empty)
            {
                return (int)zanTime;
            }

            bool bn = false;
            DateTime dtStart = DateTime.Now;        // 開始時刻
            DateTime dtEnd = DateTime.Now;          // 終了時刻
            DateTime dtChange = DateTime.Now;       // 日替わり時刻

            // 対象のシフトコード取得する
            string sftCode = string.Empty;

            if (m.シフトコード != string.Empty)
            {
                // 変更シフトコードあり
                sftCode = m.シフトコード.PadLeft(4, '0');
            }
            else if (r.シフトコード.ToString() != string.Empty)
            {
                // 標準シフトコード
                sftCode = r.シフトコード.ToString().PadLeft(4, '0');
            }

            // 有効なシフトコードが存在しないとき
            if (sftCode == string.Empty)
            {
                return (int)zanTime;
            }

            //// 奉行SQLServer接続文字列取得
            //string sc = sqlControl.obcConnectSting.get(_dbName);
            //sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select tbLaborSystem.LaborSystemCode, LaborSystemName, tbLaborSystem.LatterHalfStartTime, tbLaborSystem.FirstHalfEndTime,");
            sb.Append("tbLaborSystem.DayChangeTime,a.StartTime,a.EndTime ");
            sb.Append("from tbLaborSystem left join ");
            sb.Append("(select * from tbLaborTimeSpanRule where LaborTimeItemID = 1) as a ");
            sb.Append("on tbLaborSystem.LaborSystemID = a.LaborSystemID ");
            sb.Append("where tbLaborSystem.LaborSystemCode = '").Append(sftCode).Append("'");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                if (!(dR["StartTime"] is DBNull))
                {
                    bn = true;
                    dtStart = (DateTime)dR["StartTime"];
                }

                if (!(dR["EndTime"] is DBNull))
                {
                    bn = true;
                    dtEnd = (DateTime)dR["EndTime"];
                    dtChange = (DateTime)dR["DayChangeTime"];
                }
            }

            dR.Close();

            //sdCon.Close();

            // 開始時刻、終了時刻がないときは戻る
            if (!bn)
            {
                return (int)zanTime;
            }
            else
            {
                // 今日の日付に変換
                dtStart = DateTime.Parse(dtStart.Hour.ToString() + ":" + dtStart.Minute.ToString() + ":0");
                dtEnd = DateTime.Parse(dtEnd.Hour.ToString() + ":" + dtEnd.Minute.ToString() + ":0");
            }

            // 勤務体系の終了時刻が翌日のとき（夜勤など）
            if (dtStart > dtEnd)
            {
                // 日付を翌日にする
                dtEnd = dtEnd.AddDays(1);
            }

            // 出勤時刻と退勤時刻
            DateTime wStartTm = DateTime.Parse(m.出勤時 + ":" + m.出勤分 + ":0");
            DateTime wEndTm = DateTime.Parse(m.退勤時 + ":" + m.退勤分 + ":0");

            // 日替わり時刻 : 2017/09/19
            int nextTime = dtChange.Hour * 100 + dtChange.Minute;

            // 2017/09/19
            int sTm = Utility.StrtoInt(m.出勤時) * 100 + Utility.StrtoInt(m.出勤分);

            // 勤務開始時刻が日替わり時刻以前のとき翌日とみなす : 2017/09/19
            if (sTm < nextTime)
            {
                wStartTm = wStartTm.AddDays(1);  // 日付を翌日にする
            }

            // 2017/09/19
            int eTm = Utility.StrtoInt(m.退勤時) * 100 + Utility.StrtoInt(m.退勤分);

            // 勤務終了時刻が日替わり時刻以前のとき、または退勤が翌日のとき翌日とみなす : 2017/09/19
            if (wStartTm > wEndTm || eTm < nextTime)
            {
                wEndTm = wEndTm.AddDays(1);  // 日付を翌日にする
            }

            //// 退勤が翌日のとき
            //if (wStartTm > wEndTm)
            //{
            //    wEndTm = wEndTm.AddDays(1);  // 日付を翌日にする
            //}

            // 退勤時刻が勤務終了時刻以降のとき超過時間計算を行う
            if (wEndTm > dtEnd)
            {
                if (wStartTm < dtEnd)
                {
                    /* 出勤時刻が勤務体系の終了時刻以前のとき
                     * 例）勤務体系8:00-17:00で勤務時間が8:00-21:00等 */
                    zanTime += Utility.GetTimeSpan(dtEnd, wEndTm).TotalMinutes;
                }
                else
                {
                    /*出勤時刻が勤務体系の終了時刻以降のとき
                     * 例）勤務体系8:00-17:00で勤務時間が18:00-23:00等 */
                    zanTime += Utility.GetTimeSpan(wStartTm, wEndTm).TotalMinutes;
                }
            }

            // 出勤時刻が勤務開始時刻以前のとき早出時間計算を行う
            if (wStartTm < dtStart)
            {
                if (dtStart < wEndTm)
                {
                    /* 退勤時刻時刻が勤務体系の開始時刻以降のとき
                     * 例）勤務体系8:00-17:00で勤務時間が6:00-17:00等 */
                    zanTime += Utility.GetTimeSpan(wStartTm, dtStart).TotalMinutes;
                }
                else
                {
                    /* 退勤時刻時刻が勤務体系の開始時刻以前のとき
                     * 例）勤務体系8:00-17:00で勤務時間が6:00-7:00等 */
                    zanTime += Utility.GetTimeSpan(wStartTm, wEndTm).TotalMinutes;
                }
            }

            return (int)zanTime;
        }

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     「事由」が記入されている場合「シフト外」</summary>
        /// <param name="m">
        ///     DataSet1.勤務票明細Row </param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///-----------------------------------------------------------------------
        private bool errCheckSftJiyu(DataSet1.勤務票明細Row m)
        {
            bool rtn = true;

            if (m.シフト通り == global.FLGON)
            {
                if (m.事由1 != string.Empty || m.事由2 != string.Empty || m.事由3 != string.Empty)
                {
                    return false;
                }
            }

            return rtn;
        }

        private bool errCheckSftJiyu(DataSet1.過去勤務票明細Row m)
        {
            bool rtn = true;

            if (m.シフト通り == global.FLGON)
            {
                if (m.事由1 != string.Empty || m.事由2 != string.Empty || m.事由3 != string.Empty)
                {
                    return false;
                }
            }

            return rtn;
        }

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     「残業」が記入されている場合「シフト外」</summary>
        /// <param name="r">
        ///     DataSet1.勤務票ヘッダRow</param>
        /// <param name="m">
        ///     DataSet1.勤務票明細Row </param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///-----------------------------------------------------------------------
        private bool errCheckSftZangyo(DataSet1.勤務票明細Row m)
        {
            bool rtn = true;

            if (m.シフト通り == global.FLGON)
            {
                if (m.残業理由1 != string.Empty || m.残業時1 != string.Empty || m.残業分1 != string.Empty ||
                    m.残業理由2 != string.Empty || m.残業時2 != string.Empty || m.残業分2 != string.Empty)
                {
                    return false;
                }
            }

            return rtn;
        }

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     「休出」でシフト通りはエラー： 休憩ありを含む</summary>
        /// <param name="r">
        ///     DataSet1.勤務票ヘッダRow</param>
        /// <param name="m">
        ///     DataSet1.勤務票明細Row </param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///-----------------------------------------------------------------------
        private bool errCheckSftKyushutsu(DataSet1.勤務票ヘッダRow r, DataSet1.勤務票明細Row m)
        {
            bool rtn = true;

            if (m.シフト通り == global.FLGON)
            {
                // 休日出勤のときエラー :
                if (Utility.StrtoInt(m.シフトコード) == global.SFT_KYUSHUTSU || r.シフトコード == global.SFT_KYUSHUTSU ||
                    Utility.StrtoInt(m.シフトコード) == global.SFT_KYUKEI_KYUSHUTSU || r.シフトコード == global.SFT_KYUKEI_KYUSHUTSU)
                {
                    return false;
                }
            }

            return rtn;
        }
        
        ///-----------------------------------------------------------------------
        /// <summary>
        ///     「休出」でシフト通りはエラー： 休憩ありを含む</summary>
        /// <param name="r">
        ///     DataSet1.過去勤務票ヘッダRow</param>
        /// <param name="m">
        ///     DataSet1.過去勤務票明細Row </param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///-----------------------------------------------------------------------
        private bool errCheckSftKyushutsu(DataSet1.過去勤務票ヘッダRow r, DataSet1.過去勤務票明細Row m)
        {
            bool rtn = true;

            if (m.シフト通り == global.FLGON)
            {
                // 休日出勤のときエラー :
                if (Utility.StrtoInt(m.シフトコード) == global.SFT_KYUSHUTSU || r.シフトコード == global.SFT_KYUSHUTSU ||
                    Utility.StrtoInt(m.シフトコード) == global.SFT_KYUKEI_KYUSHUTSU || r.シフトコード == global.SFT_KYUKEI_KYUSHUTSU)
                {
                    return false;
                }
            }

            return rtn;
        }

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     「残業」が記入されている場合「シフト外」</summary>
        /// <param name="r">
        ///     DataSet1.過去勤務票明細Row</param>
        /// <param name="m">
        ///     DataSet1.過去勤務票明細Row </param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///-----------------------------------------------------------------------
        private bool errCheckSftZangyo(DataSet1.過去勤務票明細Row m)
        {
            bool rtn = true;

            if (m.シフト通り == global.FLGON)
            {
                if (m.残業理由1 != string.Empty || m.残業時1 != string.Empty || m.残業分1 != string.Empty ||
                    m.残業理由2 != string.Empty || m.残業時2 != string.Empty || m.残業分2 != string.Empty)
                {
                    return false;
                }
            }

            return rtn;
        }
        ///-----------------------------------------------------------------------
        /// <summary>
        ///     休出のとき「残業」無記入のときエラー</summary>
        /// <param name="r">
        ///     DataSet1.勤務票ヘッダRow </param>
        /// <param name="m">
        ///     DataSet1.勤務票明細Row </param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///-----------------------------------------------------------------------
        private bool errKyushutsuZangyo(DataSet1.勤務票ヘッダRow r,　DataSet1.勤務票明細Row m)
        {
            bool rtn = true;
            int sft = global.flgOff;

            if (m.シフトコード != string.Empty)
            {
                sft = Utility.StrtoInt(m.シフトコード);
            }
            else
            {
                sft = r.シフトコード;
            }

            // 2018/02/05 休出・休憩ありを追加
            if (sft == global.SFT_KYUSHUTSU || sft == global.SFT_KYUKEI_KYUSHUTSU)
            {
                if (m.残業理由1 == string.Empty && m.残業時1 == string.Empty && m.残業分1 == string.Empty && 
                    m.残業理由2 == string.Empty && m.残業時2 == string.Empty && m.残業分2 == string.Empty)
                {
                    rtn = false;
                }
            }

            return rtn;
        }

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     休出の残業時間チェック : 休憩あり休出を条件に追加 2018/02/04</summary>
        /// <param name="r">
        ///     DataSet1.過去勤務票ヘッダRow </param>
        /// <param name="m">
        ///     DataSet1.過去勤務票明細Row </param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///-----------------------------------------------------------------------
        private bool errKyushutsuZangyo(DataSet1.過去勤務票ヘッダRow r, DataSet1.過去勤務票明細Row m)
        {
            bool rtn = true;
            int sft = global.flgOff;

            if (m.シフトコード != string.Empty)
            {
                sft = Utility.StrtoInt(m.シフトコード);
            }
            else
            {
                sft = r.シフトコード;
            }

            // 2018/02/05 休出・休憩ありを条件に追加
            if (sft == global.SFT_KYUSHUTSU || sft == global.SFT_KYUKEI_KYUSHUTSU)
            {
                if (m.残業理由1 == string.Empty && m.残業時1 == string.Empty && m.残業分1 == string.Empty &&
                    m.残業理由2 == string.Empty && m.残業時2 == string.Empty && m.残業分2 == string.Empty)
                {
                    rtn = false;
                }
            }

            return rtn;
        }

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     休出の残業時間チェック : 休憩あり休出を条件に追加 2018/02/04</summary>
        /// <param name="r">
        ///     DataSet1.勤務票ヘッダRow </param>
        /// <param name="m">
        ///     DataSet1.勤務票明細Row </param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///-----------------------------------------------------------------------
        private bool errKyushutsuZanTime(DataSet1 dts, DataSet1.勤務票ヘッダRow r, DataSet1.勤務票明細Row m, out double uZan)
        {
            bool rtn = true;
            uZan = 0;
            
            int sft = global.flgOff;

            if (m.シフトコード != string.Empty)
            {
                sft = Utility.StrtoInt(m.シフトコード);
            }
            else
            {
                sft = r.シフトコード;
            }

            // 2018/02/04 休憩あり休出を条件に追加
            if (sft != global.SFT_KYUSHUTSU && sft != global.SFT_KYUKEI_KYUSHUTSU)
            {
                return true;
            }

            DateTime sDT = DateTime.Parse(m.出勤時 + ":" + m.出勤分 + ":0");
            DateTime eDT = DateTime.Parse(m.退勤時 + ":" + m.退勤分 + ":0");

            double sp = Utility.GetTimeSpan(sDT, eDT).TotalMinutes;

            // 休出・休憩ありのとき 2018/02/04
            if (sft == global.SFT_KYUKEI_KYUSHUTSU)
            {
                // 開始から終業が４時間超のとき休憩１時間とする 2018/02/04
                if (sp > 240)
                {
                    sp -= 60;
                }
            }

            double zanTm = Utility.StrtoDouble(m.残業時1) * 60 + (Utility.StrtoDouble(m.残業分1) * 60 / 10) + Utility.StrtoDouble(m.残業時2) * 60 + (Utility.StrtoDouble(m.残業分2) * 60 / 10);

            // 応援ありのときは応援残業含める
            if (m.応援 == global.FLGON)
            {
                // 勤怠データＩ／Ｐ票に応援チェックがあるが該当する応援移動票がないとき
                var s = dts.応援移動票明細.Where(a => a.応援移動票ヘッダRow.年 == r.年 &&
                                               a.応援移動票ヘッダRow.月 == r.月 &&
                                               a.応援移動票ヘッダRow.日 == r.日 &&
                                               a.社員番号.PadLeft(6, '0') == m.社員番号.PadLeft(6, '0') && 
                                               a.データ区分 == 2 && 
                                               a.取消 == global.FLGOFF);

                foreach (var t in s)
                {
                    uZan += Utility.StrtoDouble(t.残業時1) * 60 + (Utility.StrtoDouble(t.残業分1) * 60 / 10) + Utility.StrtoDouble(t.残業時2) * 60 + (Utility.StrtoDouble(t.残業分2) * 60 / 10);
                }

                // 応援残業を加算
                zanTm += uZan;
            }

            if (sp < zanTm)
            {
                rtn = false;
            }

            return rtn;
        }

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     休出の残業時間チェック : 休憩あり休出を条件に追加 2018/02/05</summary>
        /// <param name="r">
        ///     DataSet1.勤務票ヘッダRow </param>
        /// <param name="m">
        ///     DataSet1.勤務票明細Row </param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///-----------------------------------------------------------------------
        private bool errKyushutsuZanTime(DataSet1 dts, DataSet1.過去勤務票ヘッダRow r, DataSet1.過去勤務票明細Row m, out double uZan)
        {
            bool rtn = true;
            uZan = 0;

            int sft = global.flgOff;

            if (m.シフトコード != string.Empty)
            {
                sft = Utility.StrtoInt(m.シフトコード);
            }
            else
            {
                sft = r.シフトコード;
            }

            // 2018/02/05 休出・休憩ありを条件に追加
            if (sft != global.SFT_KYUSHUTSU || sft != global.SFT_KYUKEI_KYUSHUTSU)
            {
                return true;
            }
            
            DateTime sDT = DateTime.Parse(m.出勤時 + ":" + m.出勤分 + ":0");
            DateTime eDT = DateTime.Parse(m.退勤時 + ":" + m.退勤分 + ":0");

            double sp = Utility.GetTimeSpan(sDT, eDT).TotalMinutes;

            // 休出・休憩ありのとき 2018/02/04
            if (sft == global.SFT_KYUKEI_KYUSHUTSU)
            {
                // 開始から終業が４時間超のとき休憩１時間とする 2018/02/04
                if (sp > 240)
                {
                    sp -= 60;
                }
            }

            double zanTm = Utility.StrtoDouble(m.残業時1) * 60 + (Utility.StrtoDouble(m.残業分1) * 60 / 10) + Utility.StrtoDouble(m.残業時2) * 60 + (Utility.StrtoDouble(m.残業分2) * 60 / 10);

            // 応援ありのときは応援残業含める
            if (m.応援 == global.FLGON)
            {
                // 勤怠データＩ／Ｐ票に応援チェックがあるが該当する応援移動票がないとき
                var s = dts.応援移動票明細.Where(a => a.応援移動票ヘッダRow.年 == r.年 &&
                                               a.応援移動票ヘッダRow.月 == r.月 &&
                                               a.応援移動票ヘッダRow.日 == r.日 &&
                                               a.社員番号.PadLeft(6, '0') == m.社員番号.PadLeft(6, '0') &&
                                               a.データ区分 == 2 &&
                                               a.取消 == global.FLGOFF);

                foreach (var t in s)
                {
                    uZan += Utility.StrtoDouble(t.残業時1) * 60 + (Utility.StrtoDouble(t.残業分1) * 60 / 10) + Utility.StrtoDouble(t.残業時2) * 60 + (Utility.StrtoDouble(t.残業分2) * 60 / 10);
                }

                // 応援残業を加算
                zanTm += uZan;
            }

            if (sp < zanTm)
            {
                rtn = false;
            }

            return rtn;
        }

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     休出のとき「事由」記入のときエラー
        ///     : 2018/02/04 休憩あり・休出を条件に追加</summary>
        /// <param name="r">
        ///     DataSet1.勤務票ヘッダRow </param>
        /// <param name="m">
        ///     DataSet1.勤務票明細Row </param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///-----------------------------------------------------------------------
        private bool errKyushutsuJiyu(DataSet1.勤務票ヘッダRow r, DataSet1.勤務票明細Row m, out int jNum)
        {
            jNum = 0;
            bool rtn = true;
            int sft = global.flgOff;

            if (m.シフトコード != string.Empty)
            {
                sft = Utility.StrtoInt(m.シフトコード);
            }
            else
            {
                sft = r.シフトコード;
            }

            // 休日出勤 2018/02/04 休憩あり・休出を条件に追加
            if (sft == global.SFT_KYUSHUTSU || sft == global.SFT_KYUKEI_KYUSHUTSU)
            {
                // 事由：30は許容 2018/02/04
                if (m.事由1 != string.Empty && Utility.StrtoInt(m.事由1) != global.JIYU_YOBIDASHI)
                {
                    jNum = 1;
                    rtn = false;
                }       // 事由：30は許容 2018/02/04
                else if (m.事由2 != string.Empty && Utility.StrtoInt(m.事由2) != global.JIYU_YOBIDASHI)
                {
                    jNum = 2;
                    rtn = false;
                }       // 事由：30は許容 2018/02/04
                else if (m.事由3 != string.Empty && Utility.StrtoInt(m.事由3) != global.JIYU_YOBIDASHI)
                {
                    jNum = 3;
                    rtn = false;
                }
            }

            return rtn;
        }

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     休出のとき「事由」記入のときエラー
        ///     : 2018/02/04 休憩あり・休出を条件に追加</summary>
        /// <param name="r">
        ///     DataSet1.過去勤務票ヘッダRow </param>
        /// <param name="m">
        ///     DataSet1.過去勤務票明細Row </param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///-----------------------------------------------------------------------
        private bool errKyushutsuJiyu(DataSet1.過去勤務票ヘッダRow r, DataSet1.過去勤務票明細Row m, out int jNum)
        {
            jNum = 0;
            bool rtn = true;
            int sft = global.flgOff;

            if (m.シフトコード != string.Empty)
            {
                sft = Utility.StrtoInt(m.シフトコード);
            }
            else
            {
                sft = r.シフトコード;
            }

            // 休日出勤 2018/02/04 休憩あり・休出を条件に追加
            if (sft == global.SFT_KYUSHUTSU || sft == global.SFT_KYUKEI_KYUSHUTSU)
            {
                // 事由：30は許容 2018/02/04
                if (m.事由1 != string.Empty && Utility.StrtoInt(m.事由1) != global.JIYU_YOBIDASHI)
                {
                    jNum = 1;
                    rtn = false;
                }       // 事由：30は許容 2018/02/04
                else if (m.事由2 != string.Empty && Utility.StrtoInt(m.事由2) != global.JIYU_YOBIDASHI)
                {
                    jNum = 2;
                    rtn = false;
                }       // 事由：30は許容 2018/02/04
                else if (m.事由3 != string.Empty && Utility.StrtoInt(m.事由3) != global.JIYU_YOBIDASHI)
                {
                    jNum = 3;
                    rtn = false;
                }
            }

            return rtn;
        }

        ///-----------------------------------------------------------------------------------------------
        /// <summary>
        ///     応援移動票項目別エラーチェック。
        ///     エラーのときヘッダ行インデックス、フィールド番号、明細行インデックス、エラーメッセージが記録される </summary>
        /// <param name="dts">
        ///     データセット</param>
        /// <param name="r">
        ///     応援移動票ヘッダ行コレクション</param>
        /// <returns>
        ///     エラーなし：true, エラー有り：false</returns>
        ///-----------------------------------------------------------------------------------------------
        /// 
        private Boolean errCheckOuen(sqlControl.DataControl sdCon, DataSet1 dts, DataSet1.応援移動票ヘッダRow r, string[] hArray)
        {
            string sDate;
            DateTime eDate;

            const string DIVID_LINE = "1002";
            const string DIVID_BUMON = "1003";
            const string DIVID_SEIHIN = "1004";

            // 確認チェック
            if (r.確認 == global.flgOff)
            {
                setErrStatus(eDataCheck, 0, "未確認の応援移動票です");
                return false;
            }

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

            // 部署コードチェック
            string dCode = getDepartmentCode(r.部署コード.ToString());
            if (!chkDepartmentCode(dCode, sdCon))
            {
                setErrStatus(eBushoCode, 0, "マスター未登録の部署コードです");
                return false;
            }

            //
            // 社員別勤怠記入欄データ
            //

            int iX = 0;
            string k = string.Empty;    // 特別休暇記号
            string yk = string.Empty;   // 有給記号

            // 応援移動票明細データ行を取得
            List<DataSet1.応援移動票明細Row> mList = dts.応援移動票明細.Where(a => a.ヘッダID == r.ID).OrderBy(a => a.ID).ToList();

            foreach (var m in mList)
            {
                // 行数
                iX++;

                // 無記入の行はチェック対象外とする
                if (m.社員番号 == string.Empty && m.ライン == string.Empty &&
                    m.部門 == string.Empty && m.製品群 == string.Empty &&
                    m.応援時 == string.Empty && m.応援分 == string.Empty && 
                    m.残業理由1 == string.Empty &&
                    m.残業時1 == string.Empty && m.残業分1 == string.Empty &&
                    m.残業理由2 == string.Empty && 
                    m.残業時2 == string.Empty && m.残業分2 == string.Empty)
                {
                    continue;
                }

                // 取消行はチェック対象外とする
                if (m.取消 == global.FLGON)
                {
                    continue;
                }

                // 社員番号：数字以外のとき
                if (!Utility.NumericCheck(Utility.NulltoStr(m.社員番号)))
                {
                    if (m.データ区分 == 1)
                    {
                        setErrStatus(eShainNo, iX - 1, "社員番号が入力されていません");
                    }
                    else
                    {
                        setErrStatus(eShainNo2, iX - 6, "社員番号が入力されていません");
                    }

                    return false;
                }

                // 登録済み社員番号マスター検証 : 出勤簿日付で判断 2017/09/28
                if (!chkShainCode(m.社員番号, sdCon, eDate))
                {
                    if (m.データ区分 == 1)
                    {
                        setErrStatus(eShainNo, iX - 1, "未登録もしくは" + eDate.ToShortDateString() + "現在、在籍していない社員番号です");
                    }
                    else
                    {
                        setErrStatus(eShainNo2, iX - 6, "未登録もしくは" + eDate.ToShortDateString() + "現在、在籍していない社員番号です");
                    }
                    return false;
                }

                // 日中・残業応援で同じ社員番号が複数記入されていてもエラーとしない：2017/08/30
                //// 日中応援で同じ社員番号が複数記入されているとエラー
                //if (!getSameNumber(mList, m.社員番号, 1))
                //{
                //    setErrStatus(eShainNo, iX - 1, "同じ社員番号のデータが複数あります");
                //    return false;
                //}

                //// 残業応援で同じ社員番号が複数記入されているとエラー
                //if (!getSameNumber(mList, m.社員番号, 2))
                //{
                //    setErrStatus(eShainNo2, iX - 6, "同じ社員番号のデータが複数あります");
                //    return false;
                //}

                // 明細記入チェック
                if (!errCheckRow(m, "応援移動票内容", iX)) return false;
                
                // ラインチェック
                string sCode = string.Empty;

                if (m.ライン.Trim() == string.Empty)
                {
                    if (m.データ区分 == 1)
                    {
                        setErrStatus(eLine, iX - 1, "ラインが無記入です");
                    }
                    else
                    {
                        setErrStatus(eLine2, iX - 6, "ラインが無記入です");
                    }
                    return false;
                }

                if (Utility.NumericCheck(m.ライン))
                {
                    sCode = m.ライン.Trim().PadLeft(10, '0');
                }
                else
                {
                    sCode = m.ライン.Trim().PadRight(10, ' ');
                }

                if (!Utility.getHisCategory(hArray, sCode, DIVID_LINE))
                {
                    if (m.データ区分 == 1)
                    {
                        setErrStatus(eLine, iX - 1, "存在しないラインです");
                    }
                    else
                    {
                        setErrStatus(eLine2, iX - 6, "存在しないラインです");
                    }
                    return false;
                }

                // 部門チェック
                if (m.部門.Trim() == string.Empty)
                {
                    if (m.データ区分 == 1)
                    {
                        setErrStatus(eBmn, iX - 1, "部門が無記入です");
                    }
                    else
                    {
                        setErrStatus(eBmn2, iX - 6, "部門が無記入です");
                    }
                    return false;
                }

                if (Utility.NumericCheck(m.部門))
                {
                    sCode = m.部門.Trim().PadLeft(10, '0');
                }
                else
                {
                    sCode = m.部門.Trim().PadRight(10, ' ');
                }

                if (!Utility.getHisCategory(hArray, sCode, DIVID_BUMON))
                {
                    if (m.データ区分 == 1)
                    {
                        setErrStatus(eBmn, iX - 1, "存在しない部門です");
                    }
                    else
                    {
                        setErrStatus(eBmn2, iX - 6, "存在しない部門です");
                    }
                    return false;
                }

                // 製品群チェック
                // 無記入を是とする：2017/07/18

                //if (m.製品群.Trim() == string.Empty)
                //{
                //    if (m.データ区分 == 1)
                //    {
                //        setErrStatus(eHin, iX - 1, "製品群が無記入です");
                //    }
                //    else
                //    {
                //        setErrStatus(eHin2, iX - 6, "製品群が無記入です");
                //    }
                //    return false;
                //}

                // 記入ありのときエラーチェックを実施する：2017/07/18
                if (m.製品群.Trim() != string.Empty)
                {
                    if (Utility.NumericCheck(m.製品群))
                    {
                        sCode = m.製品群.Trim().PadLeft(10, '0');
                    }
                    else
                    {
                        sCode = m.製品群.Trim().PadRight(10, ' ');
                    }

                    if (!Utility.getHisCategory(hArray, sCode, DIVID_SEIHIN))
                    {
                        if (m.データ区分 == 1)
                        {
                            setErrStatus(eHin, iX - 1, "存在しない製品群です");
                        }
                        else
                        {
                            setErrStatus(eHin2, iX - 6, "存在しない製品群です");
                        }
                        return false;
                    }
                }

                // 日中応援のとき
                if (m.データ区分 == 1)
                {
                    // 応援分単位
                    if (!chkZangyoMin(m.応援分))
                    {
                        setErrStatus(eOuenM, iX - 1, "応援時間分単位は０または５です");
                        return false;
                    }
                }

                // 残業応援のとき
                if (m.データ区分 == 2)
                {
                    // 残業理由
                    if (!Utility.chkZangyoRe(m.残業理由1, m.残業時1, m.残業分1))
                    {
                        setErrStatus(eZanRe1, iX - 6, "残業理由が未記入です");
                        return false;
                    }

                    if (!Utility.chkZangyoRe2(m.残業理由1, m.残業時1, m.残業分1))
                    {
                        setErrStatus(eZanH1, iX - 6, "残業時間が未記入です");
                        return false;
                    }

                    // 部署別残業理由Excelシート登録チェック
                    string reName = string.Empty;
                    if (m.残業理由1 != string.Empty)
                    {
                        if (!bs.getBushoZanRe(out reName, r.部署コード.ToString(), m.残業理由1.ToString()))
                        {
                            setErrStatus(eZanRe1, iX - 6, "該当部署に登録されていない残業理由です");
                            return false;
                        }
                    }

                    if (!Utility.chkZangyoRe(m.残業理由2, m.残業時2, m.残業分2))
                    {
                        setErrStatus(eZanRe2, iX - 6, "残業理由が未記入です");
                        return false;
                    }

                    if (!Utility.chkZangyoRe2(m.残業理由2, m.残業時2, m.残業分2))
                    {
                        setErrStatus(eZanH2, iX - 6, "残業時間が未記入です");
                        return false;
                    }

                    // 部署別残業理由Excelシート登録チェック
                    if (m.残業理由2 != string.Empty)
                    {
                        if (!bs.getBushoZanRe(out reName, r.部署コード.ToString(), m.残業理由2.ToString()))
                        {
                            setErrStatus(eZanRe2, iX - 6, "該当部署に登録されていない残業理由です");
                            return false;
                        }
                    }

                    // 残業分単位
                    if (!chkZangyoMin(m.残業分1))
                    {
                        setErrStatus(eZanM1, iX - 6, "残業分単位は０または５です");
                        return false;
                    }

                    if (!chkZangyoMin(m.残業分2))
                    {
                        setErrStatus(eZanM2, iX - 6, "残業分単位は０または５です");
                        return false;
                    }
                }

                /* 勤怠データＩ／Ｐ票データが存在するか？
                 * 勤怠データＩ／Ｐ票データに応援チェックされているか？ */
                //string eMsg = string.Empty; 
                //if (!chkOuenCheck(dts, m, out eMsg))
                //{
                //    if (m.データ区分 == 1)
                //    {
                //        setErrStatus(eOuenIP, iX - 1, eMsg);
                //    }
                //    else
                //    {
                //        setErrStatus(eOuenIP2, iX - 6, eMsg);
                //    }
                //    return false;
                //}
            }

            return true;
        }

        ///-----------------------------------------------------------------------------------------------
        /// <summary>
        ///     過去応援移動票項目別エラーチェック。
        ///     エラーのときヘッダ行インデックス、フィールド番号、明細行インデックス、エラーメッセージが記録される </summary>
        /// <param name="dts">
        ///     データセット</param>
        /// <param name="r">
        ///     応援移動票ヘッダ行コレクション</param>
        /// <returns>
        ///     エラーなし：true, エラー有り：false</returns>
        ///-----------------------------------------------------------------------------------------------
        /// 
        private Boolean errCheckOuen(sqlControl.DataControl sdCon, DataSet1 dts, DataSet1.過去応援移動票ヘッダRow r, string[] hArray)
        {
            string sDate;
            DateTime eDate;

            const string DIVID_LINE = "1002";
            const string DIVID_BUMON = "1003";
            const string DIVID_SEIHIN = "1004";

            //// 確認チェック
            //if (r.確認 == global.flgOff)
            //{
            //    setErrStatus(eDataCheck, 0, "未確認の応援移動票です");
            //    return false;
            //}

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

            // 部署コードチェック
            string dCode = getDepartmentCode(r.部署コード.ToString());
            if (!chkDepartmentCode(dCode, sdCon))
            {
                setErrStatus(eBushoCode, 0, "マスター未登録の部署コードです");
                return false;
            }

            //
            // 社員別勤怠記入欄データ
            //

            int iX = 0;                 // 日中応援
            int iZ = 0;                 // 残業応援
            string k = string.Empty;    // 特別休暇記号
            string yk = string.Empty;   // 有給記号

            // 過去応援移動票明細データ行を取得
            List<DataSet1.過去応援移動票明細Row> mList = dts.過去応援移動票明細.Where(a => a.ヘッダID == r.ID).OrderBy(a => a.ID).ToList();

            foreach (var m in mList)
            {
                // 無記入の行はチェック対象外とする
                if (m.社員番号 == string.Empty && m.ライン == string.Empty &&
                    m.部門 == string.Empty && m.製品群 == string.Empty &&
                    m.応援時 == string.Empty && m.応援分 == string.Empty &&
                    m.残業理由1 == string.Empty &&
                    m.残業時1 == string.Empty && m.残業分1 == string.Empty &&
                    m.残業理由2 == string.Empty &&
                    m.残業時2 == string.Empty && m.残業分2 == string.Empty)
                {
                    continue;
                }

                // 取消行はチェック対象外とする
                if (m.取消 == global.FLGON)
                {
                    continue;
                }

                // 行数：2017/09/20
                if (m.データ区分 == 1)
                {
                    iX++;
                }
                else
                {
                    iZ++;
                }

                // 社員番号：数字以外のとき
                if (!Utility.NumericCheck(Utility.NulltoStr(m.社員番号)))
                {
                    if (m.データ区分 == 1)
                    {
                        setErrStatus(eShainNo, iX - 1, "社員番号が入力されていません");
                    }
                    else
                    {
                        setErrStatus(eShainNo2, iZ - 1, "社員番号が入力されていません");
                    }

                    return false;
                }

                // 登録済み社員番号マスター検証 : 出勤簿日付で判断 2017/09/28
                if (!chkShainCode(m.社員番号, sdCon, eDate))
                {
                    if (m.データ区分 == 1)
                    {
                        setErrStatus(eShainNo, iX - 1, "未登録もしくは " + eDate.ToShortDateString() + " 現在、在籍していない社員番号です");
                    }
                    else
                    {
                        setErrStatus(eShainNo2, iZ - 1, "未登録もしくは " + eDate.ToShortDateString() + " 現在、在籍していない社員番号です");
                    }
                    return false;
                }
                
                // 日中・残業応援で同じ社員番号が複数記入されていてもエラーとしない：2017/08/30
                //// 日中応援で同じ社員番号が複数記入されているとエラー
                //if (!getSameNumber(mList, m.社員番号, 1))
                //{
                //    setErrStatus(eShainNo, iX - 1, "同じ社員番号のデータが複数あります");
                //    return false;
                //}

                //// 残業応援で同じ社員番号が複数記入されているとエラー
                //if (!getSameNumber(mList, m.社員番号, 2))
                //{
                //    setErrStatus(eShainNo2, iX - 6, "同じ社員番号のデータが複数あります");
                //    return false;
                //}

                // 明細記入チェック
                if (!errCheckRow(m, "応援移動票内容", iX)) return false;

                // ラインチェック
                string sCode = string.Empty;

                if (m.ライン.Trim() == string.Empty)
                {
                    if (m.データ区分 == 1)
                    {
                        setErrStatus(eLine, iX - 1, "ラインが無記入です");
                    }
                    else
                    {
                        setErrStatus(eLine2, iZ - 1, "ラインが無記入です");
                    }
                    return false;
                }

                if (Utility.NumericCheck(m.ライン))
                {
                    sCode = m.ライン.Trim().PadLeft(10, '0');
                }
                else
                {
                    sCode = m.ライン.Trim().PadRight(10, ' ');
                }

                if (!Utility.getHisCategory(hArray, sCode, DIVID_LINE))
                {
                    if (m.データ区分 == 1)
                    {
                        setErrStatus(eLine, iX - 1, "存在しないラインです");
                    }
                    else
                    {
                        setErrStatus(eLine2, iZ - 1, "存在しないラインです");
                    }
                    return false;
                }

                // 部門チェック
                if (m.部門.Trim() == string.Empty)
                {
                    if (m.データ区分 == 1)
                    {
                        setErrStatus(eBmn, iX - 1, "部門が無記入です");
                    }
                    else
                    {
                        setErrStatus(eBmn2, iZ - 1, "部門が無記入です");
                    }
                    return false;
                }

                if (Utility.NumericCheck(m.部門))
                {
                    sCode = m.部門.Trim().PadLeft(10, '0');
                }
                else
                {
                    sCode = m.部門.Trim().PadRight(10, ' ');
                }

                if (!Utility.getHisCategory(hArray, sCode, DIVID_BUMON))
                {
                    if (m.データ区分 == 1)
                    {
                        setErrStatus(eBmn, iX - 1, "存在しない部門です");
                    }
                    else
                    {
                        setErrStatus(eBmn2, iZ - 1, "存在しない部門です");
                    }
                    return false;
                }

                // 製品群チェック
                //if (m.製品群.Trim() == string.Empty)
                //{
                //    if (m.データ区分 == 1)
                //    {
                //        setErrStatus(eHin, iX - 1, "製品群が無記入です");
                //    }
                //    else
                //    {
                //        setErrStatus(eHin2, iX - 6, "製品群が無記入です");
                //    }
                //    return false;
                //}

                // 記入ありのときチェック実施：2017/07/18
                if (m.製品群.Trim() != string.Empty)
                {
                    if (Utility.NumericCheck(m.製品群))
                    {
                        sCode = m.製品群.Trim().PadLeft(10, '0');
                    }
                    else
                    {
                        sCode = m.製品群.Trim().PadRight(10, ' ');
                    }

                    if (!Utility.getHisCategory(hArray, sCode, DIVID_SEIHIN))
                    {
                        if (m.データ区分 == 1)
                        {
                            setErrStatus(eHin, iX - 1, "存在しない製品群です");
                        }
                        else
                        {
                            setErrStatus(eHin2, iZ - 1, "存在しない製品群です");
                        }
                        return false;
                    }
                }

                // 日中応援のとき
                if (m.データ区分 == 1)
                {
                    // 応援分単位
                    if (!chkZangyoMin(m.応援分))
                    {
                        setErrStatus(eOuenM, iX - 1, "応援時間分単位は０または５です");
                        return false;
                    }
                }

                // 残業応援のとき
                if (m.データ区分 == 2)
                {
                    // 残業理由
                    if (!Utility.chkZangyoRe(m.残業理由1, m.残業時1, m.残業分1))
                    {
                        setErrStatus(eZanRe1, iZ - 1, "残業理由が未記入です");
                        return false;
                    }

                    if (!Utility.chkZangyoRe2(m.残業理由1, m.残業時1, m.残業分1))
                    {
                        setErrStatus(eZanH1, iZ - 1, "残業時間が未記入です");
                        return false;
                    }

                    // 部署別残業理由Excelシート登録チェック
                    string reName = string.Empty;
                    if (m.残業理由1 != string.Empty)
                    {
                        if (!bs.getBushoZanRe(out reName, r.部署コード.ToString(), m.残業理由1.ToString()))
                        {
                            setErrStatus(eZanRe1, iZ - 1, "該当部署に登録されていない残業理由です");
                            return false;
                        }
                    }

                    if (!Utility.chkZangyoRe(m.残業理由2, m.残業時2, m.残業分2))
                    {
                        setErrStatus(eZanRe2, iZ - 1, "残業理由が未記入です");
                        return false;
                    }

                    if (!Utility.chkZangyoRe2(m.残業理由2, m.残業時2, m.残業分2))
                    {
                        setErrStatus(eZanH2, iZ - 1, "残業時間が未記入です");
                        return false;
                    }

                    // 部署別残業理由Excelシート登録チェック
                    if (m.残業理由2 != string.Empty)
                    {
                        if (!bs.getBushoZanRe(out reName, r.部署コード.ToString(), m.残業理由2.ToString()))
                        {
                            setErrStatus(eZanRe2, iZ - 1, "該当部署に登録されていない残業理由です");
                            return false;
                        }
                    }

                    // 残業分単位
                    if (!chkZangyoMin(m.残業分1))
                    {
                        setErrStatus(eZanM1, iZ - 1, "残業分単位は０または５です");
                        return false;
                    }

                    if (!chkZangyoMin(m.残業分2))
                    {
                        setErrStatus(eZanM2, iZ - 1, "残業分単位は０または５です");
                        return false;
                    }
                }

                /* 勤怠データＩ／Ｐ票データが存在するか？
                 * 勤怠データＩ／Ｐ票データに応援チェックされているか？ */
                string eMsg = string.Empty;
                if (!chkOuenCheck(dts, m, out eMsg))
                {
                    if (m.データ区分 == 1)
                    {
                        setErrStatus(eOuenIP, iX - 1, eMsg);
                    }
                    else
                    {
                        setErrStatus(eOuenIP2, iZ - 1, eMsg);
                    }
                    return false;
                }
            }

            return true;
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     検索用DepartmentCodeを取得する </summary>
        /// <returns>
        ///     DepartmentCode</returns>
        ///----------------------------------------------------------
        private string getDepartmentCode(string bCode)
        {
            string strCode = "";

            // DepartmentCode（部署コード）
            if (Utility.NumericCheck(bCode))
            {
                strCode = bCode.PadLeft(15, '0');
            }
            else
            {
                strCode = bCode.PadRight(15, ' ');
            }

            return strCode;
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

        ///---------------------------------------------------------------------------
        /// <summary>
        ///     残業のとき出勤退時刻が記入されているか </summary>
        /// <param name="m">
        ///     DataSet1.勤務票明細Row　</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///---------------------------------------------------------------------------
        private bool errCheckZanShEh(DataSet1.勤務票明細Row m)
        {
            bool rtn = true;

            if (m.残業理由1 != string.Empty || m.残業時1 != string.Empty || m.残業分1 != string.Empty ||
                m.残業理由2 != string.Empty || m.残業時2 != string.Empty || m.残業分2 != string.Empty)
            {
                if (m.出勤時 == string.Empty || m.出勤分 == string.Empty ||
                    m.退勤時 == string.Empty || m.退勤分 == string.Empty)
                {
                    return false;
                }
            }

            return rtn;
        }

        private bool errCheckZanShEh(DataSet1.過去勤務票明細Row m)
        {
            bool rtn = true;

            if (m.残業理由1 != string.Empty || m.残業時1 != string.Empty || m.残業分1 != string.Empty ||
                m.残業理由2 != string.Empty || m.残業時2 != string.Empty || m.残業分2 != string.Empty)
            {
                if (m.出勤時 == string.Empty || m.出勤分 == string.Empty ||
                    m.退勤時 == string.Empty || m.退勤分 == string.Empty)
                {
                    return false;
                }
            }

            return rtn;
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     部署コードチェック </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl オブジェクト </param>
        /// <param name="j">
        ///     部署コード</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        private bool chkDepartmentCode(string s, sqlControl.DataControl sdCon)
        {
            bool dm = false;

            //// 奉行SQLServer接続文字列取得
            //string sc = sqlControl.obcConnectSting.get(_dbName);
            //sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            // 部署コード取得
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select DepartmentName from tbDepartment ");
            sb.Append("where DepartmentCode = '" + s + "'");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                dm = true;
                break;
            }

            dR.Close();
            //sdCon.Close();

            return dm;
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     社員コードチェック </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl オブジェクト </param>
        /// <param name="s">
        ///     社員番号</param>
        /// <param name="sDt">
        ///     基準年月日</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        private bool chkShainCode(string s, sqlControl.DataControl sdCon, DateTime sDt)
        {
            bool dm = false;

            //// 奉行SQLServer接続文字列取得
            //string sc = sqlControl.obcConnectSting.get(_dbName);
            //sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            // 社員コード取得
            /* 入社年月日 EnterCorpDate <= 基準年月日（出勤簿日付）
             * 退職年月日 RetireCorpDate >= 基準年月日（出勤簿日付）
             * で在籍を判断 : 2017/09/28
             */
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select EmployeeNo,RetireCorpDate from tbEmployeeBase ");
            sb.Append("where EmployeeNo = '" + s.PadLeft(10, '0') + "' ");
            sb.Append("and EnterCorpDate <= '" + sDt.ToShortDateString() + "' ");
            sb.Append("and RetireCorpDate >= '" + sDt.ToShortDateString()+ "' ");

            //sb.Append(" and BeOnTheRegisterDivisionID != 9");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                dm = true;
                break;
            }

            dR.Close();

            //sdCon.Close();

            return dm;
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     シフトコードチェック </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl オブジェクト </param>
        /// <param name="j">
        ///     シフトコード</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        private bool chkSftCode(sqlControl.DataControl sdCon, string s)
        {
            bool dm = false;

            //// 奉行SQLServer接続文字列取得
            //string sc = sqlControl.obcConnectSting.get(_dbName);
            //sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            // 勤務体系（シフト）コード取得
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select LaborSystemCode, LaborSystemName from tbLaborSystem ");
            sb.Append("where LaborSystemCode = '" + s.ToString().PadLeft(4, '0') + "' and ");
            sb.Append("IsValid = " + global.FLGON);
            
            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                dm = true;
                break;
            }

            dR.Close();

            //sdCon.Close();

            return dm;
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     応援移動票データ存在チェック </summary>
        /// <param name="r">
        ///     DataSet1.勤務票ヘッダRow </param>
        /// <param name="m">
        ///     DataSet1.勤務票明細Row </param>
        /// <returns>
        ///     true:データあり、false:データなし</returns>
        ///----------------------------------------------------------------------
        private bool chkOuenData(DataSet1.勤務票ヘッダRow r, DataSet1.勤務票明細Row m, DataSet1 dts)
        {
            if (m.応援 == global.FLGON)
            {
                // 勤怠データＩ／Ｐ票に応援チェックがあるが該当する応援移動票がないとき
                if (!dts.応援移動票明細.Any(a => a.応援移動票ヘッダRow.年 == r.年 &&
                                               a.応援移動票ヘッダRow.月 == r.月 &&
                                               a.応援移動票ヘッダRow.日 == r.日 &&
                                               a.社員番号.PadLeft(6, '0') == m.社員番号.PadLeft(6, '0') && 
                                               a.取消 == global.FLGOFF))
                {
                    return false;
                }
            }

            return true;
        }
        
        private bool chkOuenData(DataSet1.過去勤務票ヘッダRow r, DataSet1.過去勤務票明細Row m, DataSet1 dts)
        {
            if (m.応援 == global.FLGON)
            {
                // 勤怠データＩ／Ｐ票に応援チェックがあるが該当する応援移動票がないとき
                if (!dts.過去応援移動票明細.Any(a => a.過去応援移動票ヘッダRow.年 == r.年 &&
                                               a.過去応援移動票ヘッダRow.月 == r.月 &&
                                               a.過去応援移動票ヘッダRow.日 == r.日 &&
                                               a.社員番号.PadLeft(6, '0') == m.社員番号.PadLeft(6, '0') &&
                                               a.取消 == global.FLGOFF))
                {
                    return false;
                }
            }

            return true;
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     応援移動票データ存在チェック </summary>
        /// <param name="r">
        ///     DataSet1.勤務票ヘッダRow </param>
        /// <param name="m">
        ///     DataSet1.勤務票明細Row </param>
        /// <returns>
        ///     true:データあり、false:データなし</returns>
        ///----------------------------------------------------------------------
        private bool chkOuenData2(DataSet1.勤務票ヘッダRow r, DataSet1.勤務票明細Row m, DataSet1 dts)
        {
            if (m.応援 == global.FLGOFF)
            {
                // 勤怠データＩ／Ｐ票に応援チェックはないが同日、同社員番号の応援移動票が存在するとき
                if (dts.応援移動票明細.Any(a => a.応援移動票ヘッダRow.年 == r.年 &&
                                               a.応援移動票ヘッダRow.月 == r.月 &&
                                               a.応援移動票ヘッダRow.日 == r.日 &&
                                               a.社員番号.PadLeft(6, '0') == m.社員番号.PadLeft(6, '0') && 
                                               a.取消 == global.FLGOFF))
                {
                    return false;
                }
            }

            return true;
        }

        private bool chkOuenData2(DataSet1.過去勤務票ヘッダRow r, DataSet1.過去勤務票明細Row m, DataSet1 dts)
        {
            if (m.応援 == global.FLGOFF)
            {
                // 過去勤怠データＩ／Ｐ票に応援チェックはないが同日、同社員番号の応援移動票が存在するとき
                if (dts.過去応援移動票明細.Any(a => a.過去応援移動票ヘッダRow.年 == r.年 &&
                                               a.過去応援移動票ヘッダRow.月 == r.月 &&
                                               a.過去応援移動票ヘッダRow.日 == r.日 &&
                                               a.社員番号.PadLeft(6, '0') == m.社員番号.PadLeft(6, '0') &&
                                               a.取消 == global.FLGOFF))
                {
                    return false;
                }
            }

            return true;
        }
        
        ///-------------------------------------------------------------------------------------
        /// <summary>
        ///     応援移動票データに対応する応援「○」勤怠データＩ／Ｐ票明細データがあるか </summary>
        /// <param name="dts">
        ///     DataSet1 </param>
        /// <returns>
        ///     true:あり、false:なし</returns>
        ///-------------------------------------------------------------------------------------
        private bool chkOuenCheck(DataSet1 dts)
        {
            bool rtn = true;

            foreach (var t in dts.応援移動票明細.OrderBy(a => a.ID))
            {
                if (!dts.勤務票明細.Any(a => a.勤務票ヘッダRow.年 == t.応援移動票ヘッダRow.年 && 
                                           a.勤務票ヘッダRow.月 == t.応援移動票ヘッダRow.月 && 
                                           a.勤務票ヘッダRow.日 == t.応援移動票ヘッダRow.日 && 
                                           a.社員番号.PadLeft(6, '0') == t.社員番号.PadLeft(6, '0') && 
                                           a.応援 == global.FLGON))
                {
                    rtn = false;
                    break;
                }
            }

            return rtn;
        }

        ///-------------------------------------------------------------------------------------
        /// <summary>
        ///     応援移動票データに対応する勤怠データＩ／Ｐ票明細データがあるか </summary>
        /// <param name="dts">
        ///     DataSet1 </param>
        /// <returns>
        ///     true:あり、false:なし</returns>
        ///-------------------------------------------------------------------------------------
        private bool chkOuenCheck(DataSet1 dts, DataSet1.応援移動票明細Row m, out string msg)
        {
            msg = string.Empty;
            bool rtn = true;

            if (!dts.勤務票明細.Any(a => a.勤務票ヘッダRow.年 == m.応援移動票ヘッダRow.年 &&
                                       a.勤務票ヘッダRow.月 == m.応援移動票ヘッダRow.月 &&
                                       a.勤務票ヘッダRow.日 == m.応援移動票ヘッダRow.日 &&
                                       a.社員番号.PadLeft(6, '0') == m.社員番号.PadLeft(6, '0') && 
                                       a.取消 == global.FLGOFF))
            {
                msg = m.応援移動票ヘッダRow.年.ToString() + "/" + m.応援移動票ヘッダRow.月.ToString() + "/" + m.応援移動票ヘッダRow.日.ToString() + " の勤怠データＩ／Ｐ票が存在しません";
                rtn = false;
            }
            else
            {
                if (!dts.勤務票明細.Any(a => a.勤務票ヘッダRow.年 == m.応援移動票ヘッダRow.年 &&
                                           a.勤務票ヘッダRow.月 == m.応援移動票ヘッダRow.月 &&
                                           a.勤務票ヘッダRow.日 == m.応援移動票ヘッダRow.日 &&
                                           a.社員番号.PadLeft(6, '0') == m.社員番号.PadLeft(6, '0') && 
                                           a.取消 == global.FLGOFF && 
                                           a.応援 == global.FLGON))
                {
                    msg = "勤怠データＩ／Ｐ票に応援チェックされていません";
                    rtn = false;
                }
            }

            return rtn;
        }

        ///-------------------------------------------------------------------------------------
        /// <summary>
        ///     過去応援移動票データに対応する過去勤怠データＩ／Ｐ票明細データがあるか </summary>
        /// <param name="dts">
        ///     DataSet1 </param>
        /// <returns>
        ///     true:あり、false:なし</returns>
        ///-------------------------------------------------------------------------------------
        private bool chkOuenCheck(DataSet1 dts, DataSet1.過去応援移動票明細Row m, out string msg)
        {
            msg = string.Empty;
            bool rtn = true;

            if (!dts.過去勤務票明細.Any(a => a.過去勤務票ヘッダRow.年 == m.過去応援移動票ヘッダRow.年 &&
                                       a.過去勤務票ヘッダRow.月 == m.過去応援移動票ヘッダRow.月 &&
                                       a.過去勤務票ヘッダRow.日 == m.過去応援移動票ヘッダRow.日 &&
                                       a.社員番号.PadLeft(6, '0') == m.社員番号.PadLeft(6, '0') &&
                                       a.取消 == global.FLGOFF))
            {
                msg = m.過去応援移動票ヘッダRow.年.ToString() + "/" + m.過去応援移動票ヘッダRow.月.ToString() + "/" + m.過去応援移動票ヘッダRow.日.ToString() + " の勤怠データＩ／Ｐ票が存在しません";
                rtn = false;
            }
            else
            {
                if (!dts.過去勤務票明細.Any(a => a.過去勤務票ヘッダRow.年 == m.過去応援移動票ヘッダRow.年 &&
                                           a.過去勤務票ヘッダRow.月 == m.過去応援移動票ヘッダRow.月 &&
                                           a.過去勤務票ヘッダRow.日 == m.過去応援移動票ヘッダRow.日 &&
                                           a.社員番号.PadLeft(6, '0') == m.社員番号.PadLeft(6, '0') &&
                                           a.取消 == global.FLGOFF &&
                                           a.応援 == global.FLGON))
                {
                    msg = "過去勤怠データＩ／Ｐ票に応援チェックされていません";
                    rtn = false;
                }
            }

            return rtn;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     同日、同社員番号の件数を調べる </summary>
        /// <param name="dts">
        ///     </param>
        /// <param name="hID">
        ///     勤務票ヘッダID</param>
        /// <param name="syy">
        ///     年</param>
        /// <param name="smm">
        ///     月</param>
        /// <param name="sdd">
        ///     日</param>
        /// <param name="sNum">
        ///     社員番号</param>
        /// <returns>
        ///     なし:true, あり:false</returns>
        ///------------------------------------------------------------------------------------
        private bool getSameDateNumber(DataSet1 dts, string hID, int syy, int smm, int sdd, string sNum)
        {
            bool rtn = true;
            DateTime dt;

            if (!DateTime.TryParse(syy+ "/" + smm + "/" + sdd, out dt))
            {
                return true;
            }

            // 同日、同社員番号の件数を調べる
            foreach (var t in dts.勤務票ヘッダ.Where(a => a.ID != hID && a.年 == syy && a.月 == smm && a.日 == sdd))
	        {
                if (t.Get勤務票明細Rows().Count(a => a.社員番号 == sNum) > 0)
                {
                    return false;
                }
	        }

            return rtn;
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

        private bool getSameNumber(List<DataSet1.過去勤務票明細Row> mList, string sNum)
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

        private bool getSameNumber(List<DataSet1.応援移動票明細Row> mList, string sNum, int dKbn)
        {
            bool rtn = true;

            if (sNum == string.Empty) return rtn;

            // 指定した社員番号の件数を調べる
            if (mList.Count(a => Utility.StrtoInt(a.社員番号) == Utility.StrtoInt(sNum) && a.取消 == global.FLGOFF && a.データ区分 == dKbn) > 1)
            {
                rtn = false;
            }

            return rtn;
        }

        private bool getSameNumber(List<DataSet1.過去応援移動票明細Row> mList, string sNum, int dKbn)
        {
            bool rtn = true;

            if (sNum == string.Empty) return rtn;

            // 指定した社員番号の件数を調べる
            if (mList.Count(a => Utility.StrtoInt(a.社員番号) == Utility.StrtoInt(sNum) && a.取消 == global.FLGOFF && a.データ区分 == dKbn) > 1)
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

        private bool errCheckRow(DataSet1.過去勤務票明細Row m, string tittle, int iX)
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

        private bool errCheckRow(DataSet1.応援移動票明細Row m, string tittle, int iX)
        {
            // 社員番号以外に記入項目なしのときエラーとする
            if (m.社員番号 != string.Empty && m.ライン == string.Empty &&
                m.部門 == string.Empty && m.製品群 == string.Empty &&
                m.応援時 == string.Empty && m.応援分 == string.Empty &&
                m.残業理由1 == string.Empty && 
                m.残業時1 == string.Empty && m.残業分1 == string.Empty &&
                m.残業理由2 == string.Empty && 
                m.残業時2 == string.Empty && m.残業分2 == string.Empty)
            {
                setErrStatus(eSH, iX - 1, tittle + "が未入力です");
                return false;
            }

            return true;
        }

        private bool errCheckRow(DataSet1.過去応援移動票明細Row m, string tittle, int iX)
        {
            // 社員番号以外に記入項目なしのときエラーとする
            if (m.社員番号 != string.Empty && m.ライン == string.Empty &&
                m.部門 == string.Empty && m.製品群 == string.Empty &&
                m.応援時 == string.Empty && m.応援分 == string.Empty &&
                m.残業理由1 == string.Empty &&
                m.残業時1 == string.Empty && m.残業分1 == string.Empty &&
                m.残業理由2 == string.Empty &&
                m.残業時2 == string.Empty && m.残業分2 == string.Empty)
            {
                setErrStatus(eSH, iX - 1, tittle + "が未入力です");
                return false;
            }

            return true;
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     残業時間チェック </summary>
        /// <param name="r">
        ///     DataSet1.勤務票ヘッダRow </param>
        /// <param name="m">
        ///     DataSet1.勤務票明細Row </param>
        /// <param name="dts">
        ///     DataSet1</param>
        /// <param name="z">
        ///     計算残業時間</param>
        /// <param name="zk">
        ///     記入された残業時間</param>
        /// <returns>
        ///     エラーなし：true, エラーあり：false</returns>
        ///------------------------------------------------------------------------------------
        private bool errCheckZan(sqlControl.DataControl sdCon, DataSet1.勤務票ヘッダRow r, DataSet1.勤務票明細Row m, DataSet1 dts, out double z, out double zk, int kKbn)
        {
            z = 0;
            zk = 0;

            // 対象のシフトコードを取得する
            string sftCode = string.Empty;

            if (m.シフトコード != string.Empty)
            {
                // 変更シフトコードあり
                sftCode = m.シフトコード;
            }
            else if (r.シフトコード.ToString() != string.Empty)
            {
                // 標準シフトコード
                sftCode = r.シフトコード.ToString();
            }

            //// 休日出勤のときはチェックしない : 2017/09/20 休出時の残業計算を以下に追加
            //if (Utility.StrtoInt(sftCode) == global.SFT_KYUSHUTSU)
            //{
            //    return true;
            //}

            // 無記入なら終了
            //if (m.出勤時 == string.Empty && m.出勤分 == string.Empty) return true;


            // 休日出勤の残業時間を取得する　2017/09/20
            // 休憩あり休出を条件に追加　2018/02/04
            if (Utility.StrtoInt(sftCode) == global.SFT_KYUSHUTSU ||
                Utility.StrtoInt(sftCode) == global.SFT_KYUKEI_KYUSHUTSU)
            {
                // 出勤時刻と退勤時刻
                DateTime wStartTm = DateTime.Parse(m.出勤時 + ":" + m.出勤分 + ":0");
                DateTime wEndTm = DateTime.Parse(m.退勤時 + ":" + m.退勤分 + ":0");

                // 終了時刻が翌日か判断
                if (wStartTm > wEndTm)
                {
                    wEndTm = wEndTm.AddDays(1);
                }

                // 休日出勤（休憩無し）のとき：開始時刻～退出時刻を残業時間とする 2018/02/04
                z = (int)Utility.GetTimeSpan(wStartTm, wEndTm).TotalMinutes;

                // 休日出勤（休憩あり）のとき：開始時刻～退出時刻が４時間超のとき休憩１時間とする 2018/02/04
                if (Utility.StrtoInt(sftCode) == global.SFT_KYUKEI_KYUSHUTSU)
                {
                    if (z > 240)
                    {
                        z -= 60;
                    }
                }
            }
            else
            {
                // 休日出勤ではないとき、出退勤時刻とシフトコードから残業時間を求める
                // 派遣社員：2017/09/27
                if (kKbn == KOYOU_HAKEN)
                {
                    // 稼働時間を取得：2017/09/27
                    z = getWorkTime_Haken(kKbn, sdCon, m);

                    // 8時間超を残業時間とする
                    if (z > 480)
                    {
                        z = z - 480;
                    }
                    else
                    {
                        z = 0;
                    }
                }
                else
                {
                    // 派遣社員以外
                    z = getSftZanTime(sdCon, r, m);
                }
            }

            // 勤怠データＩ／Ｐ票に記入された残業時間
            zk = (Utility.StrtoInt(m.残業時1) * 60 + (int)(60 * Utility.StrtoDouble(m.残業分1) / 10)) + (Utility.StrtoInt(m.残業時2) * 60 + (int)(60 * Utility.StrtoDouble(m.残業分2) / 10));

            // 応援移動票の残業時間を加算する
            if (m.応援 == global.FLGON)
            {
                var s = dts.応援移動票明細.Where(a => a.応援移動票ヘッダRow.年 == r.年 &&
                                               a.応援移動票ヘッダRow.月 == r.月 &&
                                               a.応援移動票ヘッダRow.日 == r.日 &&
                                               a.社員番号.PadLeft(6, '0') == m.社員番号.PadLeft(6, '0') && 
                                               a.データ区分 == 2);
                foreach (var t in s)
                {
                    zk += (Utility.StrtoInt(t.残業時1) * 60 + (int)(60 * Utility.StrtoDouble(t.残業分1) / 10)) + (Utility.StrtoInt(t.残業時2) * 60 + (int)(60 * Utility.StrtoDouble(t.残業分2) / 10));
                }
            }
            
            //// 帰宅後勤務の残業を加算する
            //var k = dts.帰宅後勤務.Where(a => a.年 == r.年 && a.月 == r.月 && a.日 == r.日 &&
            //                               a.社員番号.PadLeft(6, '0') == m.社員番号.PadLeft(6, '0'));
            //foreach (var t in k)
            //{
            //    zk += (Utility.StrtoInt(t.残業時1) * 60 + (int)(60 * Utility.StrtoDouble(t.残業分1) / 10)) + (Utility.StrtoInt(t.残業時2) * 60 + (int)(60 * Utility.StrtoDouble(t.残業分2) / 10));
            //}
            
            // 計算値を30分単位に丸める
            z = (z - (z % 30));

            // 計算値と記入値が一致しているか？
            if (z == zk)
            {
                return true;
            }
            else
            {
                z = z / 60;
                zk = zk / 60;
                return false;
            }
        }

        ///---------------------------------------------------------------------------------
        /// <summary>
        ///     残業時間チェック </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl</param>
        /// <param name="r">
        ///     DataSet1.過去勤務票ヘッダRow</param>
        /// <param name="m">
        ///     DataSet1.過去勤務票明細Row</param>
        /// <param name="dts">
        ///     DataSet1</param>
        /// <param name="z">
        ///     残業時間</param>
        /// <param name="zk">
        ///     帰宅後勤務残業時間</param>
        /// <param name="kKbn">
        ///     雇用区分 : 2017/09/27</param>
        /// <returns>
        ///     true:エラーなし、false:エラー有り</returns>
        ///---------------------------------------------------------------------------------
        private bool errCheckZan(sqlControl.DataControl sdCon, DataSet1.過去勤務票ヘッダRow r, DataSet1.過去勤務票明細Row m, DataSet1 dts, out double z, out double zk, int kKbn)
        {
            z = 0;
            zk = 0;

            // 対象のシフトコードを取得する
            string sftCode = string.Empty;

            if (m.シフトコード != string.Empty)
            {
                // 変更シフトコードあり
                sftCode = m.シフトコード;
            }
            else if (r.シフトコード.ToString() != string.Empty)
            {
                // 標準シフトコード
                sftCode = r.シフトコード.ToString();
            }

            //// 休日出勤のときはチェックしない
            //if (Utility.StrtoInt(sftCode) == global.SFT_KYUSHUTSU)
            //{
            //    return true;
            //}

            // 無記入なら終了
            //if (m.出勤時 == string.Empty && m.出勤分 == string.Empty) return true;


            // 休憩あり休出を条件に追加　2018/02/04
            if (Utility.StrtoInt(sftCode) == global.SFT_KYUSHUTSU ||
                Utility.StrtoInt(sftCode) == global.SFT_KYUKEI_KYUSHUTSU)
            {
                // 出勤時刻と退勤時刻
                DateTime wStartTm = DateTime.Parse(m.出勤時 + ":" + m.出勤分 + ":0");
                DateTime wEndTm = DateTime.Parse(m.退勤時 + ":" + m.退勤分 + ":0");

                // 終了時刻が翌日か判断
                if (wStartTm > wEndTm)
                {
                    wEndTm = wEndTm.AddDays(1);
                }

                // 休日出勤（休憩無し）のとき：開始時刻～退出時刻を残業時間とする 2018/02/04
                z = (int)Utility.GetTimeSpan(wStartTm, wEndTm).TotalMinutes;

                // 休日出勤（休憩あり）のとき：開始時刻～退出時刻が４時間超のとき休憩１時間とする 2018/02/04
                if (Utility.StrtoInt(sftCode) == global.SFT_KYUKEI_KYUSHUTSU)
                {
                    if (z > 240)
                    {
                        z -= 60;
                    }
                }
            }
            else
            {
                // 休日出勤ではないとき、出退勤時刻とシフトコードから残業時間を求める
                // 派遣社員：2017/09/27
                if (kKbn == KOYOU_HAKEN)
                {
                    // 稼働時間を取得：2017/09/27
                    z = getWorkTime_Haken(kKbn, sdCon, m);

                    // 8時間超を残業時間とする
                    if (z > 480)
                    {
                        z = z - 480;
                    }
                    else
                    {
                        z = 0;
                    }
                }

                // 派遣社員以外
                z = getSftZanTime(sdCon, r, m);
            }


            // 勤怠データＩ／Ｐ票に記入された残業時間
            zk = (Utility.StrtoInt(m.残業時1) * 60 + (int)(60 * Utility.StrtoDouble(m.残業分1) / 10)) + (Utility.StrtoInt(m.残業時2) * 60 + (int)(60 * Utility.StrtoDouble(m.残業分2) / 10));

            // 応援移動票の残業時間を加算する
            if (m.応援 == global.FLGON)
            {
                var s = dts.過去応援移動票明細.Where(a => a.過去応援移動票ヘッダRow.年 == r.年 &&
                                               a.過去応援移動票ヘッダRow.月 == r.月 &&
                                               a.過去応援移動票ヘッダRow.日 == r.日 &&
                                               a.社員番号.PadLeft(6, '0') == m.社員番号.PadLeft(6, '0') &&
                                               a.データ区分 == 2);
                foreach (var t in s)
                {
                    zk += (Utility.StrtoInt(t.残業時1) * 60 + (int)(60 * Utility.StrtoDouble(t.残業分1) / 10)) + (Utility.StrtoInt(t.残業時2) * 60 + (int)(60 * Utility.StrtoDouble(t.残業分2) / 10));
                }
            }

            //// 帰宅後勤務の残業を加算する
            //var k = dts.帰宅後勤務.Where(a => a.年 == r.年 && a.月 == r.月 && a.日 == r.日 &&
            //                               a.社員番号.PadLeft(6, '0') == m.社員番号.PadLeft(6, '0'));
            //foreach (var t in k)
            //{
            //    zk += (Utility.StrtoInt(t.残業時1) * 60 + (int)(60 * Utility.StrtoDouble(t.残業分1) / 10)) + (Utility.StrtoInt(t.残業時2) * 60 + (int)(60 * Utility.StrtoDouble(t.残業分2) / 10));
            //}

            // 計算値を30分単位に丸める
            z = (z - (z % 30));

            // 計算値と記入値が一致しているか？
            if (z == zk)
            {
                return true;
            }
            else
            {
                z = z / 60;
                zk = zk / 60;
                return false;
            }
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


        ///-----------------------------------------------------------------------------
        /// <summary>
        ///     奉行の社員マスターより社員名と雇用区分を取得する </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl</param>
        /// <param name="sNum">
        ///     社員番号</param>
        /// <param name="sKbn">
        ///     雇用区分</param>
        ///-----------------------------------------------------------------------------
        private void getEmployee(sqlControl.DataControl sdCon, string sNum, ref int sKbn)
        {
            SqlDataReader dR = null;

            try
            {
                // 該当者の雇用区分を取得する
                StringBuilder sb = new StringBuilder();
                sb.Clear();
                sb.Append("select EmployeeNo,Name, tbHR_DivisionCategory.CategoryCode ");
                sb.Append("from tbEmployeeBase inner join tbHR_DivisionCategory ");
                sb.Append("on tbEmployeeBase.EmploymentDivisionID = tbHR_DivisionCategory.CategoryID ");
                sb.Append("where EmployeeNo = '" + sNum.PadLeft(10, '0') + "'");

                dR = sdCon.free_dsReader(sb.ToString());

                while (dR.Read())
                {
                    //sName = dR["Name"].ToString();
                    sKbn = Utility.StrtoInt(dR["CategoryCode"].ToString());
                    break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                dR.Close();
            }
        }

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     稼働時間を求める </summary>
        /// <param name="kKbn">
        ///     雇用区分</param>
        /// <param name="rtnArray">
        ///     出力配列</param>
        /// <param name="sdCon">
        ///     sqlControl.DataControl</param>
        /// <param name="t">
        ///     DataSet1.勤務票明細Row</param>
        /// <returns>
        /// 　　稼働時間</returns>
        ///------------------------------------------------------------------------------------
        private double getWorkTime_Haken(int kKbn, sqlControl.DataControl sdCon, DataSet1.勤務票明細Row t)
        {
            // 出勤時間計算
            DateTime cTm = DateTime.Now;
            DateTime sTime = DateTime.Now;
            DateTime eTime = DateTime.Now;
            DateTime restSTime = DateTime.Now;
            DateTime restETime = DateTime.Now;
            DateTime dayChangeTime = DateTime.Now;
            double wt = 0;

            // 派遣
            if (t.出勤時.Trim() != string.Empty)
            {
                // シフト通りではないとき（時刻記入のとき）
                if (DateTime.TryParse(t.出勤時 + ":" + t.出勤分, out cTm))
                {
                    sTime = cTm;

                    if (DateTime.TryParse(t.退勤時 + ":" + t.退勤分, out cTm))
                    {
                        eTime = cTm;

                        // シフトコードから休憩時刻を求める
                        getSftRestTime(sdCon, t.勤務票ヘッダRow.シフトコード.ToString(), t.シフトコード, ref restSTime, ref restETime, ref dayChangeTime);

                        // 出勤時間を求める
                        datetimeAdjust(ref sTime, ref eTime, ref restSTime, ref restETime, dayChangeTime);
                        wt = workTimePart(sTime, eTime, restSTime, restETime);
                        return wt;
                    }
                    else
                    {
                        return 0;
                    }
                }
                else
                {
                    return 0;
                }
            }
            else
            {
                // シフトコードから求める
                getSftTime(sdCon, t.勤務票ヘッダRow.シフトコード.ToString(), t.シフトコード, ref sTime, ref eTime, ref restSTime, ref restETime);
                wt = workTimePart(sTime, eTime, restSTime, restETime);
                return wt;
            }
        }


        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     稼働時間を求める </summary>
        /// <param name="kKbn">
        ///     雇用区分</param>
        /// <param name="rtnArray">
        ///     出力配列</param>
        /// <param name="sdCon">
        ///     sqlControl.DataControl</param>
        /// <param name="t">
        ///     DataSet1.過去勤務票明細Row</param>
        /// <returns>
        /// 　　稼働時間</returns>
        ///------------------------------------------------------------------------------------
        private double getWorkTime_Haken(int kKbn, sqlControl.DataControl sdCon, DataSet1.過去勤務票明細Row t)
        {
            // 出勤時間計算
            DateTime cTm = DateTime.Now;
            DateTime sTime = DateTime.Now;
            DateTime eTime = DateTime.Now;
            DateTime restSTime = DateTime.Now;
            DateTime restETime = DateTime.Now;
            DateTime dayChangeTime = DateTime.Now;
            double wt = 0;

            // 派遣
            if (t.出勤時.Trim() != string.Empty)
            {
                // シフト通りではないとき（時刻記入のとき）
                if (DateTime.TryParse(t.出勤時 + ":" + t.出勤分, out cTm))
                {
                    sTime = cTm;

                    if (DateTime.TryParse(t.退勤時 + ":" + t.退勤分, out cTm))
                    {
                        eTime = cTm;

                        // シフトコードから休憩時刻を求める
                        getSftRestTime(sdCon, t.過去勤務票ヘッダRow.シフトコード.ToString(), t.シフトコード, ref restSTime, ref restETime, ref dayChangeTime);

                        // 出勤時間を求める
                        datetimeAdjust(ref sTime, ref eTime, ref restSTime, ref restETime, dayChangeTime);
                        wt = workTimePart(sTime, eTime, restSTime, restETime);
                        return wt;
                    }
                    else
                    {
                        return 0;
                    }
                }
                else
                {
                    return 0;
                }
            }
            else
            {
                // シフトコードから求める
                getSftTime(sdCon, t.過去勤務票ヘッダRow.シフトコード.ToString(), t.シフトコード, ref sTime, ref eTime, ref restSTime, ref restETime);
                wt = workTimePart(sTime, eTime, restSTime, restETime);
                return wt;
            }
        }



        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     対象のシフトコードの休憩開始時刻・終了時刻を取得する </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl</param>
        /// <param name="sSftCode">
        ///     標準シフトコード</param>
        /// <param name="hSftCode">
        ///     変更シフトコード</param>
        /// <param name="restDtStart">
        ///     休憩開始時刻</param>
        /// <param name="restDtEnd">
        ///     休憩終了時刻</param>
        ///----------------------------------------------------------------------------------
        private void getSftRestTime(sqlControl.DataControl sdCon, string sSftCode, string hSftCode, ref DateTime restDtStart, ref DateTime restDtEnd, ref DateTime dayChange)
        {
            bool bn = false;
            DateTime dtChange = DateTime.Now;       // 日替わり時刻

            // 対象のシフトコードを取得する
            string sftCode = string.Empty;

            if (hSftCode != string.Empty)
            {
                // 変更シフトコードあり
                sftCode = hSftCode.PadLeft(4, '0');
            }
            else if (sSftCode.ToString() != string.Empty)
            {
                // 標準シフトコード
                sftCode = sSftCode.ToString().PadLeft(4, '0');
            }

            // 有効なシフトコードが存在しないとき
            if (sftCode == string.Empty)
            {
                return;
            }

            StringBuilder sb = new StringBuilder();

            // 休憩開始時刻、休憩終了時刻を求める
            sb.Clear();
            sb.Append("SELECT tbLaborSystem.LaborSystemID,LaborSystemCode,tbRestTimeSpanRule.StartTime,");
            sb.Append("tbRestTimeSpanRule.EndTime, tbLaborSystem.DayChangeTime ");
            sb.Append("from tbLaborSystem inner join tbRestTimeSpanRule ");
            sb.Append("on tbLaborSystem.LaborSystemID = tbRestTimeSpanRule.LaborSystemID ");
            sb.Append("where LaborSystemCode = '").Append(sftCode).Append("'");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                restDtStart = (DateTime)dR["StartTime"];
                restDtEnd = (DateTime)dR["EndTime"];
                dayChange = (DateTime)dR["DayChangeTime"];    // 2017/09/27
            }


            dR.Close();
        }


        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     対象のシフトコードの開始時刻・終了時刻を取得する </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl</param>
        /// <param name="sSftCode">
        ///     標準シフトコード</param>
        /// <param name="hSftCode">
        ///     変更シフトコード</param>
        /// <param name="dtStart">
        ///     開始時刻</param>
        /// <param name="dtEnd">
        ///     終了時刻</param>
        /// <param name="restDtStart">
        ///     休憩開始時刻</param>
        /// <param name="restDtEnd">
        ///     休憩終了時刻</param>
        ///----------------------------------------------------------------------------------
        private void getSftTime(sqlControl.DataControl sdCon, string sSftCode, string hSftCode, ref DateTime dtStart, ref DateTime dtEnd, ref DateTime restDtStart, ref DateTime restDtEnd)
        {
            bool bn = false;
            DateTime dtChange = DateTime.Now;       // 日替わり時刻

            // 対象のシフトコードを取得する
            string sftCode = string.Empty;

            if (hSftCode != string.Empty)
            {
                // 変更シフトコードあり
                sftCode = hSftCode.PadLeft(4, '0');
            }
            else if (sSftCode.ToString() != string.Empty)
            {
                // 標準シフトコード
                sftCode = sSftCode.ToString().PadLeft(4, '0');
            }

            // 有効なシフトコードが存在しないとき
            if (sftCode == string.Empty)
            {
                return;
            }

            // 開始時刻、終了時刻を取得する
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select tbLaborSystem.LaborSystemCode, LaborSystemName, tbLaborSystem.LatterHalfStartTime, tbLaborSystem.FirstHalfEndTime,");
            sb.Append("tbLaborSystem.DayChangeTime,a.StartTime,a.EndTime ");
            sb.Append("from tbLaborSystem left join ");
            sb.Append("(select * from tbLaborTimeSpanRule where LaborTimeItemID = 1) as a ");
            sb.Append("on tbLaborSystem.LaborSystemID = a.LaborSystemID ");
            sb.Append("where tbLaborSystem.LaborSystemCode = '").Append(sftCode).Append("'");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                if (!(dR["StartTime"] is DBNull))
                {
                    bn = true;
                    dtStart = (DateTime)dR["StartTime"];
                }

                if (!(dR["EndTime"] is DBNull))
                {
                    bn = true;
                    dtEnd = (DateTime)dR["EndTime"];
                    dtChange = (DateTime)dR["DayChangeTime"];
                }
            }

            dR.Close();

            // 休憩開始時刻、休憩終了時刻を求める
            sb.Clear();
            sb.Append("SELECT tbLaborSystem.LaborSystemID,LaborSystemCode,tbRestTimeSpanRule.StartTime,");
            sb.Append("tbRestTimeSpanRule.EndTime ");
            sb.Append("from tbLaborSystem inner join tbRestTimeSpanRule ");
            sb.Append("on tbLaborSystem.LaborSystemID = tbRestTimeSpanRule.LaborSystemID ");
            sb.Append("where LaborSystemCode = '").Append(sftCode).Append("'");

            dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                restDtStart = (DateTime)dR["StartTime"];
                restDtEnd = (DateTime)dR["EndTime"];
            }

            dR.Close();
        }
        
        ///-------------------------------------------------------------------
        /// <summary>
        ///     経過時間を求める </summary>
        /// <param name="sTime">
        ///     出勤時刻</param>
        /// <param name="eTime">
        ///     退勤時刻</param>
        /// <returns>
        ///     出勤時間</returns>
        /// <param name="restSTime">
        ///     休憩開始時刻</param>
        /// <param name="restETime">
        ///     休憩終了時刻</param>
        ///-------------------------------------------------------------------
        private double workTimePart(DateTime sTime, DateTime eTime, DateTime restSTime, DateTime restETime)
        {
            double wt = Utility.GetTimeSpan(sTime, eTime).TotalMinutes;
            double restTime = 0;

            // 勤務時間中に休憩時刻があるとき
            if (sTime <= restSTime && restETime <= eTime)
            {
                restTime = Utility.GetTimeSpan(restSTime, restETime).TotalMinutes;
            }
            else if (restSTime <= sTime && sTime <= restETime && restETime <= eTime)
            {
                // 一部休憩時間と被るとき１
                restTime = Utility.GetTimeSpan(sTime, restETime).TotalMinutes;
            }
            else if (sTime <= restSTime && restETime < eTime && eTime <= restETime)
            {
                // 一部休憩時間と被るとき２
                restTime = Utility.GetTimeSpan(restSTime, eTime).TotalMinutes;
            }

            // 休憩時間を差し引く
            wt = wt - restTime;

            //wt = (wt - (wt % 30));  // 計算値を30分単位に丸める
            //wt = wt / 60;

            return wt;
        }

        ///--------------------------------------------------------------------------
        /// <summary>
        ///     開始終了時間に当日、翌日の日付情報を与える </summary>
        /// <param name="sTime">
        ///     開始時刻</param>
        /// <param name="eTime">
        ///     終了時刻</param>
        /// <param name="restSTime">
        ///     時間帯開始時刻</param>
        /// <param name="restETime">
        ///     時間帯終了時刻</param>
        ///--------------------------------------------------------------------------
        private void datetimeAdjust(ref DateTime sTime, ref DateTime eTime, ref DateTime restSTime, ref DateTime restETime, DateTime dayChengeTime)
        {
            int st = sTime.Hour * 100 + sTime.Minute;
            int et = eTime.Hour * 100 + eTime.Minute;
            int rSt = restSTime.Hour * 100 + restSTime.Minute;
            int rEt = restETime.Hour * 100 + restETime.Minute;
            int nxt = dayChengeTime.Hour * 100 + dayChengeTime.Minute;  // 日替わり時刻 : 2017/09/27

            DateTime nDt = DateTime.Now;
            sTime = new DateTime(nDt.Year, nDt.Month, nDt.Day, sTime.Hour, sTime.Minute, 0); // 今日の日付とする

            // 開始時刻より終了時刻が小さいときは翌日
            if (st > et)
            {
                eTime = new DateTime(nDt.AddDays(1).Year, nDt.AddDays(1).Month, nDt.AddDays(1).Day, eTime.Hour, eTime.Minute, 0);
            }
            else
            {
                eTime = new DateTime(nDt.Year, nDt.Month, nDt.Day, eTime.Hour, eTime.Minute, 0);
            }

            // 開始時刻が翌日のとき : 2017/09/27
            if (st < nxt)
            {
                sTime = new DateTime(nDt.AddDays(1).Year, nDt.AddDays(1).Month, nDt.AddDays(1).Day, sTime.Hour, sTime.Minute, 0);
                eTime = new DateTime(nDt.AddDays(1).Year, nDt.AddDays(1).Month, nDt.AddDays(1).Day, eTime.Hour, eTime.Minute, 0);
            }

            
            //// 開始時刻より時間帯開始時刻が小さいときは翌日
            //if (st > rSt)
            //{
            //    restSTime = new DateTime(nDt.AddDays(1).Year, nDt.AddDays(1).Month, nDt.AddDays(1).Day, restSTime.Hour, restSTime.Minute, 0);
            //}
            //else
            //{
            //    restSTime = new DateTime(nDt.Year, nDt.Month, nDt.Day, restSTime.Hour, restSTime.Minute, 0);
            //}

            //// 休憩開始時刻より時間帯終了時刻が小さいときは翌日
            //if (rSt > rEt)
            //{
            //    restETime = new DateTime(restSTime.AddDays(1).Year, restSTime.AddDays(1).Month, restSTime.AddDays(1).Day, restETime.Hour, restETime.Minute, 0);
            //}
            //else
            //{
            //    restETime = new DateTime(restSTime.Year, restSTime.Month, restSTime.Day, restETime.Hour, restETime.Minute, 0);
            //}

            // 奉行マスタから取得した日付情報の日付が[2]当日、[3]翌日
            if (restSTime.Day == 3)
            {
                restSTime = new DateTime(nDt.AddDays(1).Year, nDt.AddDays(1).Month, nDt.AddDays(1).Day, restSTime.Hour, restSTime.Minute, 0);
            }
            else
            {
                restSTime = new DateTime(nDt.Year, nDt.Month, nDt.Day, restSTime.Hour, restSTime.Minute, 0);
            }

            if (restETime.Day == 3)
            {
                restETime = new DateTime(nDt.AddDays(1).Year, nDt.AddDays(1).Month, nDt.AddDays(1).Day, restETime.Hour, restETime.Minute, 0);
            }
            else
            {
                restETime = new DateTime(nDt.Year, nDt.Month, nDt.Day, restETime.Hour, restETime.Minute, 0);
            }
        }
        

    }
}
