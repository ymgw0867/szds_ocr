using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SZDS_TIMECARD
{
    class global
    {
        //public static string pblImagePath = "";

        #region 画像表示倍率（%）・座標
        public float miMdlZoomRate = 0f;     // 現在の表示倍率
        public float ZOOM_RATE = 0.32f;      // 標準倍率
        public float ZOOM_MAX = 2.00f;       // 最大倍率
        public float ZOOM_MIN = 0.05f;       // 最小倍率
        public float ZOOM_STEP = 0.05f;      // ステップ倍率
        public float ZOOM_NOW = 0.0f;        // 現在の倍率

        public int RECTD_NOW = 0;            // 現在の座標
        public int RECTS_NOW = 0;            // 現在の座標
        public int RECT_STEP = 20;           // ステップ座標
        #endregion

        //和暦西暦変換
        public const int rekiCnv = 1988;    //西暦、和暦変換

        #region 就業奉行汎用データヘッダ項目
        public const string H1 = @"""EBAS001""";    // 社員番号
        public const string H2 = @"""LTLT001""";    // 日付
        public const string H3 = @"""LTLT003""";    // 勤務体系コード（使用 2013/11/11）: "001"
        public const string H4 = @"""LTLT004""";    // 事由コード
        public const string H5 = @"""LTDT001""";    // 出勤時刻
        public const string H6 = @"""LTDT002""";    // 退出時刻
        public const string H7 = @"""LTDT003""";    // 外出時刻（未使用）
        public const string H8 = @"""LTDT004""";    // 戻入時刻（未使用）
        public const string H9 = @"""LTTC001""";    // 勤怠時間項目コード１：出勤時間
        public const string H10 = @"""LTTC002""";   // 勤怠時間項目コード２：休憩時間
        public const string H14 = @"""LTTC003""";   // 勤怠時間項目コード３：休日勤務時間
        public const string H11 = @"""LTTS001""";   // 時間１：出勤時間
        public const string H12 = @"""LTTS002""";   // 時間２：休憩時間
        public const string H15 = @"""LTTS003""";   // 時間３：休日勤務時間

        // 給与奉行汎用データヘッダ項目
        public const string H13 = @"""SPPM280""";   // 通勤手当
        #endregion

        #region ローカルMDB関連定数
        public const string MDBFILE = "SZDS.mdb";           // MDBファイル名
        public const string MDBTEMP = "SZDS_Temp.mdb";      // 最適化一時ファイル名
        public const string MDBBACK = "Backmdb.mdb";        // 最適化後バックアップファイル名
        #endregion

        #region フラグオン・オフ定数
        public const int flgOn = 1;            //フラグ有り(1)
        public const int flgOff = 0;           //フラグなし(0)
        public const string FLGON = "1";
        public const string FLGOFF = "0";
        #endregion

        public static int pblDenNum = 0;            // データ数

        public const int configKEY = 1;         // 環境設定データキー
        public const int mailKey = 1;           // メール設定データキー

        //ＯＣＲ処理ＣＳＶデータの検証要素
        public const int CSVLENGTH = 197;          //データフィールド数 2011/06/11
        public const int CSVFILENAMELENGTH = 21;   //ファイル名の文字数 2011/06/11  
 
        // 勤務記録表
        public const int STARTTIME = 8;            // 単位記入開始時間帯
        public const int ENDTIME = 22;             // 単位記入終了時間帯
        public const int TANNIMAX = 4;             // 単位最大値
        public const int WEEKLIMIT = 160;          // 週労働時間基準単位：40時間
        public const int DAYLIMIT = 32;            // 一日あたり労働時間基準単位：8時間

        #region 環境設定項目
        public int cnfYear = 0;                  // 対象年
        public int cnfMonth = 0;                 // 対象月
        public string cnfPath = string.Empty;    // 受け渡しデータ作成パス
        public string cnfImgPath = string.Empty; // 画像保存先パス
        public string cnfLogPath = string.Empty; // ログデータ作成パス
        public int cnfArchived = 0;              // データ保管期間（月数）
        #endregion

        #region コード桁数定数
        public const int ShozokuLength = 0;                 // 所属コード桁数
        public const int ShainLength = 0;                   // 社員コード桁数
        public const int ShozokuMaxLength = 4;              // 所属コードＭＡＸ桁数
        public const int ShainMaxLength = 4;                // 社員コードＭＡＸ桁数
        #endregion  

        #region 勤怠記号定数
        public const string K_SHUKIN = "1";                 // 休日出勤（デイリー）
        public const string K_KYUJITSUSHUKIN = "2";         // 休日出勤・代休無し
        public const string K_KYUJITSUSHUKIN_D = "3";       // 休日出勤・代休あり
        public const string K_YUKYU = "4";                  // 有休休暇
        public const string K_YUKYU_HAN = "5";              // 有休休暇
        public const string K_DAIKYU = "6";                 // 代休
        public const string K_TOKUBETSU_KYUKA = "7";        // 特別休暇
        public const string K_TOKUBETSU_KYUKA_MUKYU = "8";  // 特別休暇・無給（社員）
        public const string K_KOUKA = "8";                  // 公暇（パート）
        public const string K_KEKKIN = "9";                 // 欠勤
        public const string K_STOCK_KYUKA = "A";            // ストック休暇
        public const string K_KOUSHO = "B";                 // 公傷
        public const string K_SHUCCHOU = "C";               // 出張
        public const string K_KYUSHOKU = "D";               // 休職
        public const string K_SHITEI_KYUJITSU = "E";        // 振替休日
        #endregion

        #region 呼出コード定数
        public const int YOBICODE_1 = 1;                    // 呼出コード１
        public const int YOBICODE_2 = 2;                    // 呼出コード２
        #endregion

        #region 交替コード定数
        public const int KOUTAI_ASA = 1;                    // 朝番
        public const int KOUTAI_NAKA = 2;                   // 中番
        public const int KOUTAI_YORU = 3;                   // 夜番
        #endregion

        // 深夜時間帯チェック用
        public static DateTime dt2200 = DateTime.Parse("22:00");
        public static DateTime dt0000 = DateTime.Parse("0:00");
        public static DateTime dt0500 = DateTime.Parse("05:00");
        public static DateTime dt0800 = DateTime.Parse("08:00");
        public static DateTime dt2359 = DateTime.Parse("23:59");
        public const int TOUJITSU_SINYATIME = 120;      // 終了時刻が翌日のときの当日の深夜勤務時間

        // ChangeValueStatus
        public bool ChangeValueStatus = true;

        public const int MAX_GYO = 9;
        public const int MAX_MIN = 1;

        // 雇用区分
        public const string CATEGORY_SHAIN = "正社員";
        public const string CATEGORY_PART = "パート";
        public const string CATEGORY_ARBEIT = "アルバイト";

        public const int CATEGORY_SEISHAIN = 2;
        public const int CATEGORY_HAKEN = 4;

        // ＯＣＲモード
        public static string OCR_SCAN = "1";
        public static string OCR_IMAGE = "2";

        #region 勤務管理表種別ID定数
        public const string SHAIN_ID = "1";
        public const string PART_ID = "2";
        public const string SHUKKOU_ID = "3";
        #endregion

        public string[] arrayChohyoID = { "社員","パート","出向社員" };

        // データ作成画面datagridview表示行数
        public const int _MULTIGYO = 31;

        // フォーム登録モード
        public const int FORM_ADDMODE = 0;
        public const int FORM_EDITMODE = 1;

        // 社員マスター検索該当者なしの戻り値
        public const string NO_MASTER = "NonMaster";
        public const string NO_ZAISEKI = "NonZaiseki";
        public const string NO_TAISHOKU = "NonTaishoku";
        public const string NO_KYUSHOKU = "NonKyushoku";
        
        // 車種
        public string[,] arrStyle = new string[12, 3] { { "01", "実用", "0" }, { "02", "婦人", "0" }, { "03", "軽快", "0" }, { "04", "スポーツ", "0" }, { "05", "ＭＴＢ", "0" }, { "06", "ミニ", "0" },
                                                 { "07", "子供", "0" }, { "08", "電動", "0" }, { "09", "折畳", "0" }, { "21", "単車型", "1" }, { "22", "スクーター型", "1" }, { "90", "その他", "0" }};

        // 年月日未設定値
        public static DateTime NODATE = DateTime.Parse("1900/01/01");

        // データ区分
        public static int DATA_CYCLE = 0;
        public static int DATA_AUTO = 1;

        // 地区
        public static string DATA_CPA = "CPA";

        // ＣＳＶファイル名
        public static string CSV_LOG = "logdata";   // ログデータ

        // 勤務体系コード・事由
        public static int SFT_KYUSHUTSU = 31;           // 休日出勤（休憩無し）
        public static int SFT_KYUKEI_KYUSHUTSU = 32;    // 休日出勤（休憩あり）2018/02/04
        public static int JIYU_YOBIDASHI = 30;          // 事由・呼出回数 // 2018/02/04
        public static int SFT_FURIKAESHUKKIN = 43;      // 振替出勤 // 2018/02/03
        public static int SFT_DOTOKKYU = 40;            // 事由・土曜特休 2017/11/10
        public static int SFT_TSUJYOKEKKIN = 10;        // 事由・通常欠勤 2017/11/21
        public static int JIYU_KYUGYOKEKKIN = 11;       // 事由・休業欠勤 2018/10/30
        public static int JIYU_HANKETSU_AM = 17;        // 事由・前半欠勤 2018/02/13
        public static int JIYU_HANKETSU_PM = 18;        // 事由・後半欠勤 2018/02/13
        public static int JIYU_KYUGYOCHISOU = 20;       // 事由・休業遅早 2018/10/30

        // 部門
        public static string BMN_ALL = "全社";
        public static string BMN_SEIZOU = "製造部門";
        public static string BMN_KANSETSU = "間接部門";

        // 残業推移グラフ集計単位
        public static int CHART_HAN = 3;
        public static int CHART_KAKARI = 1;
        public static int CHART_KA = 2;
        public static int CHART_NOTHING = 0;
    }
}
