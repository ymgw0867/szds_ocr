using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using SZDS_TIMECARD.Common;

namespace SZDS_TIMECARD.OCR
{
    /// <summary>
    ///     事由エラーチェック基本クラス
    /// </summary>
    public class clsJiyu
    {
        public clsJiyu(string[] s)
        {
            // 事由配列
            mJiyu = s;
        }

        protected string[] mJiyu = null;

        //internal sqlControl.DataControl sdCon = null;

        protected SqlDataReader getLaborReason(string sCode, sqlControl.DataControl sdCon)
        {
            //// 奉行SQLServer接続文字列取得
            //string sc = sqlControl.obcConnectSting.get(dbName);
            //sdCon = new sqlControl.DataControl(sc);

            // 事由データ取得
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select LaborReasonCode,AcquireUnit,AcquireDivision from tbLaborReason ");
            sb.Append("where IsValid = 1 and LaborReasonCode = '" + sCode.PadLeft(2, '0') + "'");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            return dR;
        }
    }

    /// <summary>
    ///     事由コード登録チェッククラス：基本クラスを継承</summary>
    public class clsJiyuHas : clsJiyu
    {
        ///------------------------------------------------------------
        /// <summary>
        ///     事由コードチェッククラス</summary>
        /// <param name="s">
        ///     事由配列</param>
        /// <param name="db">
        ///     データベース名</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        public clsJiyuHas(string[] s)
            : base(s)
        {

        }

        ///------------------------------------------------------------
        /// <summary>
        ///     事由コードチェック </summary>
        /// <param name="errNum">
        ///     エラー事由番号</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        public bool isHasRows(out int errNum, sqlControl.DataControl sdCon)
        {
            bool dm = true;
            errNum = 0;

            for (int i = 0; i < mJiyu.Length; i++)
            {
                if (mJiyu[i].Trim() == string.Empty)
                {
                    continue;
                }

                // 事由データ取得
                SqlDataReader dR = getLaborReason(mJiyu[i], sdCon);

                if (!dR.HasRows)
                {
                    dm = false;
                    errNum = i;
                    dR.Close();
                    break;
                }

                dR.Close();
            }

            //if (sdCon != null)
            //{
            //    sdCon.Close();
            //}

            return dm;
        }
    }

    /// <summary>
    ///     「終日」事由と他の事由併記チェッククラス：基本クラスを継承
    /// </summary>
    public class clsJiyuAllDay : clsJiyu
    {
        ///------------------------------------------------------------
        /// <summary>
        ///     「終日」事由と他の事由併記チェッククラス </summary>
        /// <param name="s">
        ///     事由配列</param>
        /// <param name="_db">
        ///     データベース名</param>
        ///------------------------------------------------------------
        public clsJiyuAllDay(string[] s)
            : base(s)
        {

        }

        bool dm = false;

        ///------------------------------------------------------------
        /// <summary>
        ///     「終日」事由と他の事由併記チェック </summary>
        /// <param name="errNum">
        ///     エラー事由番号</param>
        /// <param name="sdCon">
        ///     sqlControl.DataControlオブジェクト</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        public bool isAllDayAnotherDay(out int errNum, sqlControl.DataControl sdCon)
        {
            errNum = 0;

            for (int i = 0; i < mJiyu.Length; i++)
            {
                if (mJiyu[i].Trim() == string.Empty)
                {
                    continue;
                }

                SqlDataReader dR = getLaborReason(mJiyu[i], sdCon);

                while (dR.Read())
                {
                    // 取得区分
                    if (Utility.NulltoStr(dR["AcquireUnit"]) == global.FLGOFF)
                    {
                        // 終日「０」
                        dm = true;  // 終日あり
                    }

                    break;
                }

                dR.Close();
            }

            //if (sdCon != null)
            //{
            //    sdCon.Close();
            //}

            // 終日事由があり、他の事由が併記されているときはエラ―
            if (!dm)
            {
                // 終日がない場合戻る
                return true;
            }
            else
            {
                int cnt = 0;
                for (int i = 0; i < mJiyu.Length; i++)
                {
                    if (mJiyu[i] != string.Empty)
                    {
                        // 事由記入あり
                        cnt++;
                        errNum = i;
                    }
                }

                // 終日を含んで2つ以上の事由が記入されているとエラー
                if (cnt > 1)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
        }
    }

    /// <summary>
    ///     「終日」事由とシフト以外の記入チェッククラス：基本クラスを継承
    /// </summary>
    public class clsAlldayAnotherData : clsJiyu
    {
        ///------------------------------------------------------------
        /// <summary>
        ///     「終日」事由とシフト以外の記入チェック </summary>
        /// <param name="s">
        ///     事由配列</param>
        /// <param name="_dbName">
        ///     データベース名</param>
        ///------------------------------------------------------------
        public clsAlldayAnotherData(string[] s)
            : base(s)
        {

        }

        bool dm = false;

        ///------------------------------------------------------------
        /// <summary>
        ///     「終日」事由とシフト以外の記入チェック </summary>
        /// <param name="m">
        ///     DataSet1.勤務票明細Row</param>
        /// <param name="sdCon">
        ///     sqlControl.DataControlオブジェクト</param>
        /// <param name="eNum">
        ///     エラー項目番号</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        public bool isAlldayAnotherData(DataSet1.勤務票明細Row m, sqlControl.DataControl sdCon, out int eNum)
        {
            eNum = 0;

            for (int i = 0; i < mJiyu.Length; i++)
            {
                if (mJiyu[i].Trim() == string.Empty)
                {
                    continue;
                }

                SqlDataReader dR = getLaborReason(mJiyu[i], sdCon);

                while (dR.Read())
                {
                    // 取得区分
                    if (Utility.NulltoStr(dR["AcquireUnit"]) == global.FLGOFF)
                    {
                        // 終日「０」
                        dm = true;  // 終日あり
                    }

                    break;
                }

                dR.Close();
            }

            //if (sdCon != null)
            //{
            //    sdCon.Close();
            //}

            if (!dm)
            {
                // 終日がない場合戻る
                return true;
            }
            else
            {
                // 終日でシフトコード以外記入があるとき
                if (m.出勤時 != string.Empty)
                {
                    eNum = 1;
                    return false;
                }
                if (m.出勤分 != string.Empty)
                {
                    eNum = 2;
                    return false;
                }

                if (m.退勤時 != string.Empty)
                {
                    eNum = 3;
                    return false;
                }

                if (m.退勤分 != string.Empty)
                {
                    eNum = 4;
                    return false;
                }

                if (m.残業理由1 != string.Empty)
                {
                    eNum = 5;
                    return false;
                }

                if (m.残業時1 != string.Empty)
                {
                    eNum = 6;
                    return false;
                }

                if (m.残業分1 != string.Empty)
                {
                    eNum = 7;
                    return false;
                }

                if (m.残業理由2 != string.Empty)
                {
                    eNum = 8;
                    return false;
                }

                if (m.残業時2 != string.Empty)
                {
                    eNum = 9;
                    return false;
                }

                if (m.残業分2 != string.Empty)
                {
                    eNum = 10;
                    return false;
                }

                if (m.応援 == global.FLGON)
                {
                    eNum = 11;
                    return false;
                }

                if (m.シフトコード != string.Empty)
                {
                    eNum = 12;
                    return false;
                }

                return true;
            }
        }

        public bool isAlldayAnotherData(DataSet1.過去勤務票明細Row m, sqlControl.DataControl sdCon, out int eNum)
        {
            eNum = 0;

            for (int i = 0; i < mJiyu.Length; i++)
            {
                if (mJiyu[i].Trim() == string.Empty)
                {
                    continue;
                }

                SqlDataReader dR = getLaborReason(mJiyu[i], sdCon);

                while (dR.Read())
                {
                    // 取得区分
                    if (Utility.NulltoStr(dR["AcquireUnit"]) == global.FLGOFF)
                    {
                        // 終日「０」
                        dm = true;  // 終日あり
                    }

                    break;
                }

                dR.Close();
            }

            //if (sdCon != null)
            //{
            //    sdCon.Close();
            //}

            if (!dm)
            {
                // 終日がない場合戻る
                return true;
            }
            else
            {
                // 終日でシフトコード以外記入があるとき
                if (m.出勤時 != string.Empty)
                {
                    eNum = 1;
                    return false;
                }
                if (m.出勤分 != string.Empty)
                {
                    eNum = 2;
                    return false;
                }

                if (m.退勤時 != string.Empty)
                {
                    eNum = 3;
                    return false;
                }

                if (m.退勤分 != string.Empty)
                {
                    eNum = 4;
                    return false;
                }

                if (m.残業理由1 != string.Empty)
                {
                    eNum = 5;
                    return false;
                }

                if (m.残業時1 != string.Empty)
                {
                    eNum = 6;
                    return false;
                }

                if (m.残業分1 != string.Empty)
                {
                    eNum = 7;
                    return false;
                }

                if (m.残業理由2 != string.Empty)
                {
                    eNum = 8;
                    return false;
                }

                if (m.残業時2 != string.Empty)
                {
                    eNum = 9;
                    return false;
                }

                if (m.残業分2 != string.Empty)
                {
                    eNum = 10;
                    return false;
                }

                if (m.応援 == global.FLGON)
                {
                    eNum = 11;
                    return false;
                }

                if (m.シフトコード != string.Empty)
                {
                    eNum = 12;
                    return false;
                }

                return true;
            }
        }
    }


    /// <summary>
    ///     「終日」事由と休出シフトの記入チェック：基本クラスを継承
    /// </summary>
    public class clsAllDayOffWork : clsJiyu
    {
        ///------------------------------------------------------------
        /// <summary>
        ///     「終日」事由と休出シフトの記入チェッククラス </summary>
        /// <param name="s">
        ///     事由配列</param>
        /// <param name="_dbName">
        ///     データベース名</param>
        ///------------------------------------------------------------
        public clsAllDayOffWork(string[] s):base(s)
        {

        }

        bool dm = false;

        ///------------------------------------------------------------
        /// <summary>
        ///     「終日」事由と休出シフトの記入チェック </summary>
        /// <param name="r">
        ///     DataSet1.勤務票ヘッダRow </param>
        /// <param name="m">
        ///     DataSet1.勤務票明細Row</param>
        /// <param name="sdCon">
        ///     sqlControl.DataControlオブジェクト</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        public bool isAllDayOffWork(DataSet1.勤務票ヘッダRow r, DataSet1.勤務票明細Row m, sqlControl.DataControl sdCon)
        {
            for (int i = 0; i < mJiyu.Length; i++)
            {
                if (mJiyu[i].Trim() == string.Empty)
                {
                    continue;
                }

                SqlDataReader dR = getLaborReason(mJiyu[i], sdCon);

                while (dR.Read())
                {
                    // 取得区分
                    if (Utility.NulltoStr(dR["AcquireUnit"]) == global.FLGOFF)
                    {
                        // 終日「０」
                        dm = true;  // 終日あり
                    }

                    break;
                }

                dR.Close();
            }

            //if (sdCon != null)
            //{
            //    sdCon.Close();
            //}

            if (!dm)
            {
                // 終日がない場合戻る
                return true;
            }
            else
            {
                // 休日出勤のときエラー : 休憩あり休出を条件に追加 2018/02/04
                if (Utility.StrtoInt(m.シフトコード) == global.SFT_KYUSHUTSU || r.シフトコード == global.SFT_KYUSHUTSU ||
                    Utility.StrtoInt(m.シフトコード) == global.SFT_KYUKEI_KYUSHUTSU || r.シフトコード == global.SFT_KYUKEI_KYUSHUTSU)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
        }

        public bool isAllDayOffWork(DataSet1.過去勤務票ヘッダRow r, DataSet1.過去勤務票明細Row m, sqlControl.DataControl sdCon)
        {
            for (int i = 0; i < mJiyu.Length; i++)
            {
                if (mJiyu[i].Trim() == string.Empty)
                {
                    continue;
                }

                SqlDataReader dR = getLaborReason(mJiyu[i], sdCon);

                while (dR.Read())
                {
                    // 取得区分
                    if (Utility.NulltoStr(dR["AcquireUnit"]) == global.FLGOFF)
                    {
                        // 終日「０」
                        dm = true;  // 終日あり
                    }

                    break;
                }

                dR.Close();
            }

            //if (sdCon != null)
            //{
            //    sdCon.Close();
            //}

            if (!dm)
            {
                // 終日がない場合戻る
                return true;
            }
            else
            {
                // 休日出勤のときエラー
                if (Utility.StrtoInt(m.シフトコード) == global.SFT_KYUSHUTSU || r.シフトコード == global.SFT_KYUSHUTSU)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
        }
    }

    /// <summary>
    ///     取得単位「半日」事由の取得区分の重複記入チェッククラス：基本クラスを継承
    /// </summary>
    public class clsJiyuDiv : clsJiyu
    {
        ///------------------------------------------------------------
        /// <summary>
        ///     取得単位「半日」事由の取得区分の重複記入チェッククラス </summary>
        /// <param name="s">
        ///     事由配列</param>
        /// <param name="_dbName">
        ///     データベース名</param>
        ///------------------------------------------------------------
        public clsJiyuDiv(string[] s)
            : base(s)
        {

        }

        bool dm = false;

        ///------------------------------------------------------------
        /// <summary>
        ///     取得単位「半日」事由の取得区分の重複記入チェック </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControlオブジェクト</param>
        /// <param name="dCnt">
        ///     半日事由の数</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        public bool isJiyuDiv(sqlControl.DataControl sdCon, out int dCnt)
        {
            dCnt = 0;            
            string[] div = new string[3];

            for (int i = 0; i < mJiyu.Length; i++)
            {
                if (mJiyu[i].Trim() == string.Empty)
                {
                    div[i] = string.Empty;
                    continue;
                }

                SqlDataReader dR = getLaborReason(mJiyu[i], sdCon);

                while (dR.Read())
                {
                    // 取得単位
                    if (Utility.NulltoStr(dR["AcquireUnit"]) == global.FLGON)
                    {
                        // 半日「1」
                        dm = true;  // 終日あり
                        div[i] = Utility.NulltoStr(dR["AcquireDivision"]); // 取得区分
                    }

                    break;
                }

                dR.Close();
            }

            //if (sdCon != null)
            //{
            //    sdCon.Close();
            //}

            if (!dm)
            {
                // 半日がない場合戻る
                return true;
            }
            else
            {
                int kbn = 0;

                for (int i = 0; i < 3; i++)
                {
                    if (div[i] == string.Empty)
                    {
                        continue;
                    }

                    kbn += Utility.StrtoInt(div[i]);
                    dCnt++;
                }

                if (dCnt == 3)
                {
                    // 半日事由が3つ記入されている
                    return false;
                }
                else if (dCnt == 2)
                {
                    // 半日事由が2つ記入されている
                    if (kbn != 1)
                    {
                        // 前半(0)・前半(0)、または後半(1)・後半(1)の組み合わせになっている
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
                else
                {
                    return true;
                }
            }
        }

        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     記入されている事由が終日単位か半日単位か調べる : 2017/11/21</summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl</param>
        /// <returns>
        ///     終日は1, 半日は0.5</returns>
        ///----------------------------------------------------------------------------------
        public double jiyuDivCount(sqlControl.DataControl sdCon)
        {
            double dCnt = 0;

            for (int i = 0; i < mJiyu.Length; i++)
            {
                if (mJiyu[i] == string.Empty)
                {
                    continue;
                }

                SqlDataReader dR = getLaborReason(mJiyu[i], sdCon);

                while (dR.Read())
                {
                    // 取得単位
                    if (Utility.NulltoStr(dR["AcquireUnit"]) == global.FLGOFF)
                    {
                        // 終日事由
                        dCnt++;
                    }
                    else if (Utility.NulltoStr(dR["AcquireUnit"]) == global.FLGON)
                    {
                        // 半日事由
                        dCnt += 0.5;
                    }

                    break;
                }

                dR.Close();
            }

            return dCnt;
        }
    }


    /// <summary>
    ///     「終日」事由以外で「シフト通りではない」とき変更シフトコードまたは勤務時間の記入が必要：基本クラスを継承
    /// </summary>
    public class clsNotAlldayShift : clsJiyu
    {
        ///------------------------------------------------------------
        /// <summary>
        ///     「終日」事由とシフト以外の記入チェック </summary>
        /// <param name="s">
        ///     事由配列</param>
        /// <param name="_dbName">
        ///     データベース名</param>
        ///------------------------------------------------------------
        public clsNotAlldayShift(string[] s)
            : base(s)
        {

        }

        bool dm = false;

        ///------------------------------------------------------------
        /// <summary>
        ///     「終日」事由とシフト以外の記入チェック </summary>
        /// <param name="m">
        ///     DataSet1.勤務票明細Row</param>
        /// <param name="sdCon">
        ///     sqlControl.DataControlオブジェクト</param>
        /// <param name="eNum">
        ///     エラー項目番号</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        public bool isNotAlldayShift(DataSet1.勤務票明細Row m, sqlControl.DataControl sdCon, out int eNum)
        {
            eNum = 0;

            for (int i = 0; i < mJiyu.Length; i++)
            {
                if (mJiyu[i].Trim() == string.Empty)
                {
                    continue;
                }

                SqlDataReader dR = getLaborReason(mJiyu[i], sdCon);

                while (dR.Read())
                {
                    // 取得区分
                    if (Utility.NulltoStr(dR["AcquireUnit"]) == global.FLGOFF)
                    {
                        // 終日「０」
                        dm = true;  // 終日あり
                    }

                    break;
                }

                dR.Close();
            }

            if (dm)
            {
                // 終日のとき戻る
                return true;
            }
            else
            {
                //  シフト通りでないとき
                if (m.シフト通り == global.FLGOFF)
                {
                    // 2018/02/03　エラー条件を以下のように各々独立させた
                    if (m.シフトコード == string.Empty &&
                        m.出勤時 == string.Empty && m.出勤分 == string.Empty &&
                        m.退勤時 == string.Empty && m.退勤分 == string.Empty)
                    {
                        return false;
                    }

                    //// 2018/02/03 変更シフトコードが無記入はエラー
                    //if (m.シフトコード == string.Empty)
                    //{
                    //    return false;
                    //}

                    //// 2018/02/03 出退勤時刻が無記入のときエラー
                    //if (m.シフトコード == string.Empty &&
                    //    m.出勤時 == string.Empty && m.出勤分 == string.Empty &&
                    //    m.退勤時 == string.Empty && m.退勤分 == string.Empty)
                    //{
                    //    return false;
                    //}
                }

                return true;
            }
        }

        public bool isNotAlldayShift(DataSet1.過去勤務票明細Row m, sqlControl.DataControl sdCon, out int eNum)
        {
            eNum = 0;

            for (int i = 0; i < mJiyu.Length; i++)
            {
                if (mJiyu[i].Trim() == string.Empty)
                {
                    continue;
                }

                SqlDataReader dR = getLaborReason(mJiyu[i], sdCon);

                while (dR.Read())
                {
                    // 取得区分
                    if (Utility.NulltoStr(dR["AcquireUnit"]) == global.FLGOFF)
                    {
                        // 終日「０」
                        dm = true;  // 終日あり
                    }

                    break;
                }

                dR.Close();
            }

            if (dm)
            {
                // 終日のとき戻る
                return true;
            }
            else
            {
                if (m.シフト通り == global.FLGOFF)
                {
                    if (m.シフトコード == string.Empty &&
                        m.出勤時 == string.Empty && m.出勤分 == string.Empty &&
                        m.退勤時 == string.Empty && m.退勤分 == string.Empty)
                    {
                        return false;
                    }
                }

                return true;
            }
        }
    }


    /// <summary>
    ///     「17:前半欠勤」「18:後半欠勤」事由と他の事由併記チェッククラス：基本クラスを継承
    /// </summary>
    public class clsJiyuHankekkin : clsJiyu
    {
        ///------------------------------------------------------------
        /// <summary>
        ///     「17:前半欠勤」「18:後半欠勤」と他の事由併記チェッククラス </summary>
        /// <param name="s">
        ///     事由配列</param>
        /// <param name="_db">
        ///     データベース名</param>
        ///------------------------------------------------------------
        public clsJiyuHankekkin(string[] s)
            : base(s)
        {

        }

        bool dm = false;

        ///------------------------------------------------------------
        /// <summary>
        ///     「17:前半欠勤」「18:後半欠勤」の単独記入チェック 
        ///     2018/02/15</summary>
        /// <param name="errNum">
        ///     エラー事由番号</param>
        /// <returns>
        ///     true:エラーなし、false:エラー</returns>
        ///------------------------------------------------------------
        public bool isHankeAnotherDay(out int errNum)
        {
            errNum = 0;

            for (int i = 0; i < mJiyu.Length; i++)
            {
                if (mJiyu[i].Trim() == string.Empty)
                {
                    continue;
                }

                // 事由に前半欠勤または後半欠勤が記入されているとき
                if (Utility.StrtoInt(mJiyu[i].Trim()) == global.JIYU_HANKETSU_AM ||
                    Utility.StrtoInt(mJiyu[i].Trim()) == global.JIYU_HANKETSU_PM)
                {
                    dm = true;
                }
            }

            // 前半欠勤または後半欠勤が単独記入されているときはエラー
            if (!dm)
            {
                // 前半欠勤または後半欠勤がない場合戻る
                return true;
            }
            else
            {
                int cnt = 0;
                for (int i = 0; i < mJiyu.Length; i++)
                {
                    if (mJiyu[i] != string.Empty)
                    {
                        // 事由記入あり
                        cnt++;
                        errNum = i;
                    }
                }

                // 前半欠勤または後半欠勤のみの記入のときエラー
                if (cnt > 1)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }
    }

}
