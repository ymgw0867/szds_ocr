using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;
//using Excel = Microsoft.Office.Interop.Excel;

namespace SZDS_TIMECARD.Common
{
    class Utility
    {
        /// <summary>
        /// ウィンドウ最小サイズの設定
        /// </summary>
        /// <param name="tempFrm">対象とするウィンドウオブジェクト</param>
        /// <param name="wSize">width</param>
        /// <param name="hSize">Height</param>
        public static void WindowsMinSize(Form tempFrm, int wSize, int hSize)
        {
            tempFrm.MinimumSize = new Size(wSize, hSize);
        }

        /// <summary>
        /// ウィンドウ最小サイズの設定
        /// </summary>
        /// <param name="tempFrm">対象とするウィンドウオブジェクト</param>
        /// <param name="wSize">width</param>
        /// <param name="hSize">height</param>
        public static void WindowsMaxSize(Form tempFrm, int wSize, int hSize)
        {
            tempFrm.MaximumSize = new Size(wSize, hSize);
        }

        /// <summary>
        /// 休日コンボボックスクラス
        /// </summary>
        public class comboHoliday
        {
            public string Date { get; set; }
            public string Name { get; set; }

            ///------------------------------------------------------------------------
            /// <summary>
            ///     休日コンボボックスデータロード</summary>
            /// <param name="tempBox">
            ///     ロード先コンボボックスオブジェクト名</param>
            ///------------------------------------------------------------------------
            public static void Load(ComboBox tempBox)
            {

                // 休日配列
                string[] sDay = {"01/01元旦", "     成人の日", "02/11建国記念の日", "     春分の日", "04/29昭和の日",
                            "05/03憲法記念日","05/04みどりの日","05/05こどもの日","08/12海の日","     敬老の日",
                            "     秋分の日","     体育の日","11/03文化の日","11/23勤労感謝の日","12/23天皇誕生日",
                            "     振替休日","     国民の休日","     土曜日","     年末年始休暇","     夏季休暇"}; 

                try
                {
                    comboHoliday cmb1;

                    tempBox.Items.Clear();
                    tempBox.DisplayMember = "Name";
                    tempBox.ValueMember = "Date";

                    foreach (var a in sDay)
                    {
                        cmb1 = new comboHoliday();
                        cmb1.Date = a.Substring(0, 5);
                        int s = a.Length;
                        cmb1.Name = a.Substring(5, s - 5);
                        tempBox.Items.Add(cmb1);
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "休日コンボボックスロード");
                }
            }


            ///------------------------------------------------------------------------
            /// <summary>
            ///     休日コンボ表示 </summary>
            /// <param name="tempBox">
            ///     コンボボックスオブジェクト</param>
            /// <param name="dt">
            ///     月日</param>
            ///------------------------------------------------------------------------
            public static void selectedIndex(ComboBox tempBox, string dt)
            {
                comboHoliday cmbS = new comboHoliday();
                Boolean Sh = false;

                for (int iX = 0; iX <= tempBox.Items.Count - 1; iX++)
                {
                    tempBox.SelectedIndex = iX;
                    cmbS = (comboHoliday)tempBox.SelectedItem;

                    if (cmbS.Date == dt)
                    {
                        Sh = true;
                        break;
                    }
                }

                if (Sh == false)
                {
                    tempBox.SelectedIndex = -1;
                }
            }
        }

        ///------------------------------------------------------------------------
        /// <summary>
        ///     文字列の値が数字かチェックする </summary>
        /// <param name="tempStr">
        ///     検証する文字列</param>
        /// <returns>
        ///     数字:true,数字でない:false</returns>
            ///------------------------------------------------------------------------
        public static bool NumericCheck(string tempStr)
        {
            double d;

            if (tempStr == null) return false;

            if (double.TryParse(tempStr, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                return false;

            return true;
        }

        ///------------------------------------------------------------------------
        /// <summary>
        ///     emptyを"0"に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、not nullのときそのまま値を返す</returns>
        ///------------------------------------------------------------------------
        public static string EmptytoZero(string tempStr)
        {
            if (tempStr == string.Empty)
            {
                return "0";
            }
            else
            {
                return tempStr;
            }
        }

        ///------------------------------------------------------------------------
        /// <summary>
        ///     Nullをstring.Empty("")に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、not nullのとき文字型値を返す</returns>
        ///------------------------------------------------------------------------
        public static string NulltoStr(string tempStr)
        {
            if (tempStr == null)
            {
                return string.Empty;
            }
            else
            {
                return tempStr;
            }
        }

        ///------------------------------------------------------------------------
        /// <summary>
        ///     Nullをstring.Empty("")に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、not nullのときそのまま値を返す</returns>
        ///------------------------------------------------------------------------
        public static string NulltoStr(object tempStr)
        {
            if (tempStr == null)
            {
                return string.Empty;
            }
            else
            {
                if (tempStr == DBNull.Value)
                {
                    return string.Empty;
                }
                else
                {
                    return (string)tempStr.ToString();
                }
            }
        }

        /// <summary>
        /// 文字型をIntへ変換して返す（数値でないときは０を返す）
        /// </summary>
        /// <param name="tempStr">文字型の値</param>
        /// <returns>Int型の値</returns>
        public static int StrtoInt(string tempStr)
        {
            if (NumericCheck(tempStr)) return int.Parse(tempStr);
            else return 0;
        }

        /// <summary>
        /// 文字型をDoubleへ変換して返す（数値でないときは０を返す）
        /// </summary>
        /// <param name="tempStr">文字型の値</param>
        /// <returns>double型の値</returns>
        public static double StrtoDouble(string tempStr)
        {
            if (NumericCheck(tempStr)) return double.Parse(tempStr);
            else return 0;
        }

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     経過時間を返す </summary>
        /// <param name="s">
        ///     開始時間</param>
        /// <param name="e">
        ///     終了時間</param>
        /// <returns>
        ///     経過時間</returns>
        ///-----------------------------------------------------------------------
        public static TimeSpan GetTimeSpan(DateTime s, DateTime e)
        {
            TimeSpan ts;
            if (s > e)
            {
                TimeSpan j = new TimeSpan(24, 0, 0);
                ts = e + j - s;
            }
            else
            {
                ts = e - s;
            }

            return ts;
        }

        /// ------------------------------------------------------------------------
        /// <summary>
        ///     指定した精度の数値に切り捨てします。</summary>
        /// <param name="dValue">
        ///     丸め対象の倍精度浮動小数点数。</param>
        /// <param name="iDigits">
        ///     戻り値の有効桁数の精度。</param>
        /// <returns>
        ///     iDigits に等しい精度の数値に切り捨てられた数値。</returns>
        /// ------------------------------------------------------------------------
        public static double ToRoundDown(double dValue, int iDigits)
        {
            double dCoef = System.Math.Pow(10, iDigits);

            return dValue > 0 ? System.Math.Floor(dValue * dCoef) / dCoef :
                                System.Math.Ceiling(dValue * dCoef) / dCoef;
        }


        // 部門コンボボックスクラス
        public class ComboBumon
        {
            public string ID { get; set; }
            public string DisplayName { get; set; }
            public string Name { get; set; }
            public string code { get; set; }

            ////部門マスターロード
            //public static void load(ComboBox tempObj, int tempLen, string dbName)
            //{
            //    try
            //    {
            //        ComboBumon cmb1;
            //        string sqlSTRING = string.Empty;
            //        dbControl.DataControl dCon = new dbControl.DataControl(dbName);
            //        OleDbDataReader dR;

            //        sqlSTRING += "select * from Bumon inner join ";
            //        sqlSTRING += "(select distinct BumonId as bumonid from Shain) as sbumon ";
            //        sqlSTRING += "on Bumon.Id = sbumon.bumonid ";
            //        sqlSTRING += "order by Code";

            //        //データリーダーを取得する
            //        dR = dCon.FreeReader(sqlSTRING);

            //        tempObj.Items.Clear();
            //        tempObj.DisplayMember = "DisplayName";
            //        tempObj.ValueMember = "code";

            //        while (dR.Read())
            //        {
            //            cmb1 = new ComboBumon();
            //            cmb1.ID = int.Parse(dR["Id"].ToString());
            //            cmb1.DisplayName = string.Format("{0:D" + tempLen.ToString() + "}", int.Parse(dR["Code"].ToString())) + " " + dR["Name"].ToString().Trim() + "";
            //            cmb1.Name = dR["Name"].ToString().Trim() + "";
            //            cmb1.code = dR["Code"].ToString() + "";
            //            tempObj.Items.Add(cmb1);
            //        }

            //        dR.Close();
            //        dCon.Close();
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show(ex.Message, "部門コンボボックスロード");
            //    }
            //}

            ///----------------------------------------------------------------
            /// <summary>
            ///     ＣＳＶデータから部門コンボボックスにロードする </summary>
            /// <param name="tempObj">
            ///     コンボボックスオブジェクト</param>
            /// <param name="fName">
            ///     ＣＳＶデータファイルパス</param>
            ///----------------------------------------------------------------
            public static void loadBusho(ComboBox tempObj, string dbName)
            {
                try
                {
                    ComboBumon cmb1;

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "DisplayName";
                    tempObj.ValueMember = "code";

                    // 奉行SQLServer接続文字列取得
                    string sc = sqlControl.obcConnectSting.get(dbName);
                    sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

                    StringBuilder sb = new StringBuilder();
                    sb.Clear();
                    sb.Append("select DepartmentCode,DepartmentName from tbDepartment ");
                    sb.Append("order by DepartmentCode");

                    SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

                    while (dR.Read())
                    {
                        string dCode = string.Empty;

                        // コンボボックスにセット
                        cmb1 = new ComboBumon();
                        cmb1.ID = string.Empty;

                        if (Utility.NumericCheck(dR["DepartmentCode"].ToString()))
                        {
                            dCode = Utility.StrtoInt(dR["DepartmentCode"].ToString()).ToString().PadLeft(5, '0');
                        }
                        else
                        {
                            dCode = dR["DepartmentCode"].ToString().Trim();
                        }

                        cmb1.DisplayName = dCode + " " + dR["DepartmentName"].ToString();

                        cmb1.Name = dR["DepartmentName"].ToString();
                        cmb1.code = dCode;
                        tempObj.Items.Add(cmb1);
                    }

                    dR.Close();
                    sdCon.Close();
                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "部門コンボボックスロード");
                }
            }


            ///----------------------------------------------------------------
            /// <summary>
            ///     ＣＳＶデータから部門コンボボックスに課名をロードする </summary>
            /// <param name="tempObj">
            ///     コンボボックスオブジェクト</param>
            /// <param name="fName">
            ///     ＣＳＶデータファイルパス</param>
            ///----------------------------------------------------------------
            public static void loadKa(ComboBox tempObj, string dbName, int stts)
            {
                try
                {
                    ComboBumon cmb1;

                    tempObj.Items.Clear();
                    tempObj.DisplayMember = "DisplayName";
                    tempObj.ValueMember = "code";

                    // 奉行SQLServer接続文字列取得
                    string sc = sqlControl.obcConnectSting.get(dbName);
                    sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

                    StringBuilder sb = new StringBuilder();
                    sb.Clear();
                    sb.Append("select DepartmentCode,DepartmentName from tbDepartment ");

                    if (stts == global.CHART_HAN)
                    {
                        // 班別
                        //sb.Append("where Right(DepartmentCode, 1) != '0'");
                    }
                    else if (stts == global.CHART_KAKARI)
                    {
                        // 係別
                        sb.Append("where Right(DepartmentCode, 2) != '00' and Right(DepartmentCode, 1) = '0'");
                    }
                    else if (stts == global.CHART_KA)
                    {
                        // 課別
                        sb.Append("where Right(DepartmentCode, 2) = '00'");
                    }
                    
                    sb.Append("order by DepartmentCode");

                    SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

                    while (dR.Read())
                    {
                        string dCode = string.Empty;

                        // コンボボックスにセット
                        cmb1 = new ComboBumon();
                        cmb1.ID = string.Empty;

                        if (Utility.NumericCheck(dR["DepartmentCode"].ToString()))
                        {
                            dCode = Utility.StrtoInt(dR["DepartmentCode"].ToString()).ToString().PadLeft(5, '0');
                        }
                        else
                        {
                            dCode = dR["DepartmentCode"].ToString().Trim();
                        }

                        cmb1.DisplayName = dCode + " " + dR["DepartmentName"].ToString();

                        cmb1.Name = dR["DepartmentName"].ToString();
                        cmb1.code = dCode;
                        tempObj.Items.Add(cmb1);
                    }

                    dR.Close();
                    sdCon.Close();

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "部門コンボボックスロード");
                }
            }
        }

        //// 社員コンボボックスクラス
        //public class ComboShain
        //{
        //    public int ID { get; set; }
        //    public string DisplayName { get; set; }
        //    public string Name { get; set; }
        //    public string code { get; set; }
        //    public int YakushokuType { get; set; }
        //    public string BumonName { get; set; }
        //    public string BumonCode { get; set; }

        //    // 社員マスターロード
        //    public static void load(ComboBox tempObj, string dbName)
        //    {
        //        try
        //        {
        //            ComboShain cmb1;
        //            string sqlSTRING = string.Empty;
        //            dbControl.DataControl dCon = new dbControl.DataControl(dbName);
        //            OleDbDataReader dR;

        //            sqlSTRING += "select Id,Code, Sei, Mei, YakushokuType from Shain ";
        //            sqlSTRING += "where Shurojokyo = 1 ";
        //            sqlSTRING += "order by Code";

        //            //データリーダーを取得する
        //            dR = dCon.FreeReader(sqlSTRING);

        //            tempObj.Items.Clear();
        //            tempObj.DisplayMember = "DisplayName";
        //            tempObj.ValueMember = "code";

        //            while (dR.Read())
        //            {
        //                cmb1 = new ComboShain();
        //                cmb1.ID = int.Parse(dR["Id"].ToString());
        //                cmb1.DisplayName = dR["Code"].ToString().Trim() + " " + dR["Sei"].ToString().Trim() + "　" + dR["Mei"].ToString().Trim();
        //                cmb1.Name = dR["Sei"].ToString().Trim() + "　" + dR["Mei"].ToString().Trim();
        //                cmb1.code = (dR["Code"].ToString() + "").Trim();
        //                cmb1.YakushokuType = int.Parse(dR["YakushokuType"].ToString());
        //                tempObj.Items.Add(cmb1);
        //            }

        //            dR.Close();
        //            dCon.Close();
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show(ex.Message, "社員コンボボックスロード");
        //        }

        //    }


        //    ///----------------------------------------------------------------
        //    /// <summary>
        //    ///     ＣＳＶデータから社員コンボボックスにロードする </summary>
        //    /// <param name="tempObj">
        //    ///     コンボボックスオブジェクト</param>
        //    /// <param name="fName">
        //    ///     ＣＳＶデータファイルパス</param>
        //    ///----------------------------------------------------------------
        //    public static void loadCsv(ComboBox tempObj, string fName)
        //    {
        //        string[] bArray = null;

        //        try
        //        {
        //            ComboShain cmb1;

        //            tempObj.Items.Clear();
        //            tempObj.DisplayMember = "DisplayName";
        //            tempObj.ValueMember = "code";

        //            // 社員名簿CSV読み込み
        //            bArray = System.IO.File.ReadAllLines(fName, Encoding.Default);

        //            System.Collections.ArrayList al = new System.Collections.ArrayList();

        //            foreach (var t in bArray)
        //            {
        //                string[] d = t.Split(',');

        //                if (d.Length < 4)
        //                {
        //                    continue;
        //                }

        //                string bn = d[1].PadLeft(5, '0') + "," + d[0] + "";
        //                al.Add(bn);
        //            }

        //            // 配列をソートします
        //            al.Sort();

        //            string alCode = string.Empty;

        //            foreach (var item in al)
        //            {
        //                string[] d = item.ToString().Split(',');

        //                // 重複社員はネグる
        //                if (alCode != string.Empty && alCode.Substring(0, 5) == d[0])
        //                {
        //                    continue;
        //                }

        //                // コンボボックスにセット
        //                cmb1 = new ComboShain();
        //                cmb1.ID = 0;
        //                cmb1.DisplayName = item.ToString().Replace(',', ' ');

        //                string[] cn = item.ToString().Split(',');
        //                cmb1.Name = cn[1] + "";
        //                cmb1.code = cn[0] + "";
        //                tempObj.Items.Add(cmb1);

        //                alCode = item.ToString();
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show(ex.Message, "社員コンボボックスロード");
        //        }

        //    }

        //    ///------------------------------------------------------------------------
        //    /// <summary>
        //    ///     常陽コンピュータサービスエクセル社員マスターコンボボックスロード </summary>
        //    /// <param name="fName">
        //    ///     エクセルファイル名</param>
        //    /// <param name="sheetNum">
        //    ///     シート名</param>
        //    /// <param name="tempObj">
        //    ///     コンボボックス</param>
        //    /// <param name="szStatus">
        //    ///     ０：部門情報含めない、１：部門情報含める</param>
        //    ///------------------------------------------------------------------------
        //    public static void xlsArrayLoad(string fName, string sheetNum, ComboBox tempObj, int szStatus)
        //    {
        //        string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

        //        Excel.Application oXls = new Excel.Application();

        //        Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(fName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //                                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //                                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //                                           Type.Missing, Type.Missing));

        //        Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[sheetNum];

        //        Excel.Range dRg;
        //        Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

        //        const int C_BCODE = 7;
        //        const int C_BNAME = 8;
        //        const int C_SCODE = 11;
        //        const int C_SEI = 24;
        //        const int C_MEI = 25;

        //        int iX = 0;

        //        System.Collections.ArrayList al = new System.Collections.ArrayList();

        //        try
        //        {
        //            int frmRow = 21;  // 開始行
        //            int toRow = oxlsSheet.UsedRange.Rows.Count;

        //            for (int i = frmRow; i <= toRow; i++)
        //            {
        //                // 社員番号
        //                dRg = (Excel.Range)oxlsSheet.Cells[i, C_SCODE];

        //                // 社員番号に有効値があること
        //                string sc = dRg.Text.ToString().Trim();
        //                if (Utility.StrtoInt(sc) == 0)
        //                {
        //                    continue;
        //                }

        //                // 社員姓
        //                dRg = (Excel.Range)oxlsSheet.Cells[i, C_SEI];
        //                string sei = dRg.Text.ToString().Trim();

        //                // 社員名
        //                dRg = (Excel.Range)oxlsSheet.Cells[i, C_MEI];
        //                string mei = dRg.Text.ToString().Trim();

        //                string bn = sc.ToString().PadLeft(5, '0') + "," + (sei + " " + mei) + "";

        //                if (szStatus == global.flgOn)
        //                {
        //                    // 組織コード
        //                    dRg = (Excel.Range)oxlsSheet.Cells[i, C_BCODE];
        //                    int bCode = Utility.StrtoInt(dRg.Text.ToString().Trim());

        //                    // 組織名称
        //                    dRg = (Excel.Range)oxlsSheet.Cells[i, C_BNAME];
        //                    string bName = dRg.Text.ToString().Trim();

        //                    bn += "," + bCode.ToString() + "," + (bName + "");

        //                }

        //                al.Add(bn);

        //                iX++;
        //            }
        //        }
        //        catch (Exception e)
        //        {
        //            MessageBox.Show(e.Message, "エクセル社員マスター読み込み", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //        }
        //        finally
        //        {
        //            // ウィンドウを非表示にする
        //            oXls.Visible = false;

        //            // 保存処理
        //            oXls.DisplayAlerts = false;

        //            // Bookをクローズ
        //            oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

        //            // Excelを終了
        //            oXls.Quit();

        //            // COM オブジェクトの参照カウントを解放する 
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);
        //            oXls = null;
        //            oXlsBook = null;
        //            oxlsSheet = null;
        //            GC.Collect();
        //        }

        //        ComboShain cmb1;

        //        tempObj.Items.Clear();
        //        tempObj.DisplayMember = "DisplayName";
        //        tempObj.ValueMember = "code";

        //        // 配列をソートします
        //        al.Sort();

        //        string alCode = string.Empty;

        //        foreach (var item in al)
        //        {
        //            string[] d = item.ToString().Split(',');

        //            // 重複社員はネグる
        //            if (alCode != string.Empty && alCode.Substring(0, 5) == d[0])
        //            {
        //                continue;
        //            }

        //            // コンボボックスにセット
        //            cmb1 = new ComboShain();
        //            cmb1.ID = 0;
        //            cmb1.DisplayName = item.ToString().Replace(',', ' ');

        //            string[] cn = item.ToString().Split(',');
        //            cmb1.Name = cn[1] + "";
        //            cmb1.code = cn[0] + "";

        //            if (szStatus == global.flgOn)
        //            {
        //                cmb1.BumonCode = cn[2] + "";
        //                cmb1.BumonName = cn[3] + "";
        //            }

        //            tempObj.Items.Add(cmb1);

        //            alCode = item.ToString();
        //        }
        //    }

        //    ///------------------------------------------------------------------------
        //    /// <summary>
        //    ///     常陽コンピュータサービスエクセル社員マスター読み込み </summary>
        //    /// <param name="fName">
        //    ///     エクセルファイル名</param>
        //    /// <param name="sheetNum">
        //    ///     シート名</param>
        //    /// <param name="xS">
        //    ///     読み込む配列</param>
        //    ///------------------------------------------------------------------------
        //    public static void xlsArrayLoad(string fName, string sheetNum, ref xlsShain [] xS)
        //    {
        //        string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

        //        Excel.Application oXls = new Excel.Application();

        //        Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(fName, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //                                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //                                           Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //                                           Type.Missing, Type.Missing));

        //        Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[sheetNum];

        //        Excel.Range dRg;
        //        Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

        //        xS = null;
        //        const int C_BCODE = 7;
        //        const int C_BNAME = 8;
        //        const int C_SCODE = 11;
        //        const int C_SEI = 24;
        //        const int C_MEI = 25;

        //        int iX = 0;

        //        try
        //        {
        //            int frmRow = 21;  // 開始行
        //            int toRow = oxlsSheet.UsedRange.Rows.Count;

        //            for (int i = frmRow; i <= toRow; i++)
        //            {
        //                // 社員番号
        //                dRg = (Excel.Range)oxlsSheet.Cells[i, C_SCODE];

        //                // 社員番号に有効値があること
        //                string sc = dRg.Text.ToString().Trim();
        //                if (Utility.StrtoInt(sc) == 0)
        //                {
        //                    continue;
        //                }

        //                // 配列を加算
        //                Array.Resize(ref xS, iX + 1);
        //                xS[iX] = new xlsShain();

        //                // 社員番号
        //                xS[iX].sCode = Utility.StrtoInt(sc);

        //                // 組織コード
        //                dRg = (Excel.Range)oxlsSheet.Cells[i, C_BCODE];
        //                xS[iX].bCode = Utility.StrtoInt(dRg.Text.ToString().Trim());

        //                // 組織名称
        //                dRg = (Excel.Range)oxlsSheet.Cells[i, C_BNAME];
        //                xS[iX].bName = dRg.Text.ToString().Trim();

        //                // 社員姓
        //                dRg = (Excel.Range)oxlsSheet.Cells[i, C_SEI];
        //                string sei = dRg.Text.ToString().Trim();

        //                // 社員名
        //                dRg = (Excel.Range)oxlsSheet.Cells[i, C_MEI];
        //                string mei = dRg.Text.ToString().Trim();

        //                xS[iX].sName = sei + " " + mei;

        //                iX++;
        //            }
        //        }
        //        catch(Exception e)
        //        {
        //            MessageBox.Show(e.Message, "エクセル社員マスター読み込み", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //        }
        //        finally
        //        {
        //            // ウィンドウを非表示にする
        //            oXls.Visible = false;

        //            // 保存処理
        //            oXls.DisplayAlerts = false;

        //            // Bookをクローズ
        //            oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

        //            // Excelを終了
        //            oXls.Quit();

        //            // COM オブジェクトの参照カウントを解放する 
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);
        //        }

        //    }

        //    ///------------------------------------------------------------------
        //    /// <summary>
        //    ///     社員名簿配列から社員情報を取得する </summary>
        //    /// <param name="x">
        //    ///     社員名簿配列</param>
        //    /// <param name="sCode">
        //    ///     社員番号</param>
        //    /// <param name="zCode">
        //    ///     勤務先コード</param>
        //    /// <param name="zName">
        //    ///     勤務先名</param>
        //    /// <returns>
        //    ///     社員名</returns>
        //    ///------------------------------------------------------------------
        //    public static string getXlsSname(xlsShain[] x, int sCode, out string zCode, out string zName)
        //    {
        //        string rVal = string.Empty;
        //        zCode = string.Empty;
        //        zName = string.Empty;

        //        foreach (var t in x.Where(a => a.sCode == sCode))
        //        {
        //            rVal = t.sName;
        //            zCode = t.bCode.ToString();
        //            zName = t.bName.ToString();
        //        }

        //        return rVal;
        //    }

        //    ///------------------------------------------------------------------
        //    /// <summary>
        //    ///     社員名簿配列から部門名を取得する </summary>
        //    /// <param name="x">
        //    ///     社員名簿配列</param>
        //    /// <param name="zCode">
        //    ///     部門コード</param>
        //    /// <returns>
        //    ///     部門名</returns>
        //    ///------------------------------------------------------------------
        //    public static string getXlSzName(xlsShain[] x, int zCode)
        //    {
        //        string rVal = string.Empty;

        //        foreach (var t in x.Where(a => a.bCode == zCode))
        //        {
        //            rVal = t.bName;
        //        }

        //        return rVal;
        //    }
        //    ///------------------------------------------------------------------
        //    /// <summary>
        //    ///     社員名簿配列に指定の社員番号が存在するか調べる </summary>
        //    /// <param name="x">
        //    ///     社員名簿配列</param>
        //    /// <param name="sCode">
        //    ///     社員番号</param>
        //    /// <returns>
        //    ///     true:あり, false:なし</returns>
        //    ///------------------------------------------------------------------

        //    public static bool isXlsCode(xlsShain[] x, int sCode)
        //    {
        //        bool rVal = false;

        //        foreach (var t in x.Where(a => a.sCode == sCode))
        //        {
        //            rVal = true;
        //        }

        //        return rVal;
        //    }

        //    ///------------------------------------------------------------------
        //    /// <summary>
        //    ///     社員名簿配列に指定の部門コードが存在するか調べる </summary>
        //    /// <param name="x">
        //    ///     社員名簿配列</param>
        //    /// <param name="sCode">
        //    ///     部門コード</param>
        //    /// <returns>
        //    ///     true:あり, false:なし</returns>
        //    ///------------------------------------------------------------------
        //    public static bool isXlSzCode(xlsShain[] x, int zCode)
        //    {
        //        bool rVal = false;

        //        foreach (var t in x.Where(a => a.bCode == zCode))
        //        {
        //            rVal = true;
        //        }

        //        return rVal;
        //    }

        //    // パートタイマーロード
        //    public static void loadPart(ComboBox tempObj, string dbName)
        //    {
        //        try
        //        {
        //            ComboShain cmb1;
        //            string sqlSTRING = string.Empty;
        //            dbControl.DataControl dCon = new dbControl.DataControl(dbName);
        //            OleDbDataReader dR;
        //            sqlSTRING += "select Bumon.Code as bumoncode,Bumon.Name as bumonname,Shain.Id as shainid,";
        //            sqlSTRING += "Shain.Code as shaincode,Shain.Sei,Shain.Mei, Shain.YakushokuType ";
        //            sqlSTRING += "from Shain left join Bumon ";
        //            sqlSTRING += "on Shain.BumonId = Bumon.Id ";
        //            sqlSTRING += "where Shurojokyo = 1 and YakushokuType = 1 ";
        //            sqlSTRING += "order by Shain.Code";
                    
        //            //sqlSTRING += "select Id,Code, Sei, Mei, YakushokuType from Shain ";
        //            //sqlSTRING += "where Shurojokyo = 1 and YakushokuType = 1 ";
        //            //sqlSTRING += "order by Code";

        //            //データリーダーを取得する
        //            dR = dCon.FreeReader(sqlSTRING);

        //            tempObj.Items.Clear();
        //            tempObj.DisplayMember = "DisplayName";
        //            tempObj.ValueMember = "code";

        //            while (dR.Read())
        //            {
        //                cmb1 = new ComboShain();
        //                cmb1.ID = int.Parse(dR["shainid"].ToString());
        //                cmb1.DisplayName = dR["shaincode"].ToString().Trim() + " " + dR["Sei"].ToString().Trim() + "　" + dR["Mei"].ToString().Trim();
        //                cmb1.Name = dR["Sei"].ToString().Trim() + "　" + dR["Mei"].ToString().Trim();
        //                cmb1.code = (dR["shaincode"].ToString() + "").Trim();
        //                cmb1.YakushokuType = int.Parse(dR["YakushokuType"].ToString());
        //                cmb1.BumonCode = dR["bumoncode"].ToString().PadLeft(3, '0');
        //                cmb1.BumonName = dR["bumonname"].ToString();
        //                tempObj.Items.Add(cmb1);
        //            }

        //            dR.Close();
        //            dCon.Close();
        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show(ex.Message, "社員コンボボックスロード");
        //        }

        //    }
        //}


        //// データ領域コンボボックスクラス
        //public class ComboDataArea
        //{
        //    public string ID { get; set; }
        //    public string DisplayName { get; set; }
        //    public string Name { get; set; }
        //    public string code { get; set; }

        //    // データ領域ロード
        //    public static void load(ComboBox tempObj)
        //    {
        //        dbControl.DataControl dcon = new dbControl.DataControl(Properties.Settings.Default.SQLDataBase);
        //        OleDbDataReader dR = null;

        //        try
        //        {
        //            ComboDataArea cmb;

        //            // データリーダー取得
        //            string mySql = string.Empty;
        //            mySql += "SELECT * FROM Common_Unit_DataAreaInfo ";
        //            mySql += "where CompanyTerm = " + DateTime.Today.Year.ToString();
        //            dR = dcon.FreeReader(mySql);

        //            //会社情報がないとき
        //            if (!dR.HasRows)
        //            {
        //                MessageBox.Show("会社領域情報が存在しません", "会社領域選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //                return;
        //            }

        //            // コンボボックスにアイテムを追加します
        //            tempObj.Items.Clear();
        //            tempObj.DisplayMember = "DisplayName";

        //            while (dR.Read())
        //            {
        //                cmb = new ComboDataArea();
        //                // "CompanyCode"が数字のレコードを対象とする
        //                if (Utility.NumericCheck(dR["CompanyCode"].ToString()))
        //                {
        //                    cmb.DisplayName = dR["CompanyName"].ToString().Trim();
        //                    cmb.ID = dR["Name"].ToString().Trim();
        //                    cmb.code = dR["CompanyCode"].ToString().Trim();
        //                    tempObj.Items.Add(cmb);
        //                }
        //            }
        //        }
        //        catch (Exception e)
        //        {
        //            MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
        //        }
        //        finally
        //        {
        //            if (!dR.IsClosed) dR.Close();
        //            dcon.Close();
        //        }

        //    }
        //}


        /////--------------------------------------------------------
        ///// <summary>
        ///// 会社情報より部門コード桁数、社員コード桁数を取得
        ///// </summary>
        ///// -------------------------------------------------------
        //public class BumonShainKetasu
        //{
        //    public string ID { get; set; }
        //    public string DisplayName { get; set; }
        //    public string Name { get; set; }
        //    public string code { get; set; }

        //    // 会社情報取得
        //    public static void GetKetasu(string dbName)
        //    {
        //        dbControl.DataControl dcon = new dbControl.DataControl(dbName);
        //        OleDbDataReader dR = null;

        //        try
        //        {
        //            // データリーダー取得
        //            string mySql = string.Empty;
        //            mySql += "SELECT BumonCodeKeta,ShainCodeKeta FROM Kaisha ";
        //            dR = dcon.FreeReader(mySql);

        //            // 部門コード桁数、社員コード桁数を取得
        //            while (dR.Read())
        //            {
        //                global.ShozokuLength = int.Parse(dR["BumonCodeKeta"].ToString());
        //                global.ShainLength = int.Parse(dR["ShainCodeKeta"].ToString());
        //            }
        //        }
        //        catch (Exception e)
        //        {
        //            MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
        //        }
        //        finally
        //        {
        //            if (!dR.IsClosed) dR.Close();
        //            dcon.Close();
        //        }

        //    }
        //}


        ///------------------------------------------------------------------
        /// <summary>
        ///     ファイル選択ダイアログボックスの表示 </summary>
        /// <param name="sTitle">
        ///     タイトル文字列</param>
        /// <param name="sFilter">
        ///     ファイルのフィルター</param>
        /// <returns>
        ///     選択したファイル名</returns>
        ///------------------------------------------------------------------
        public static string userFileSelect(string sTitle, string sFilter)
        {
            DialogResult ret;

            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            //ダイアログボックスの初期設定
            openFileDialog1.Title = sTitle;
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = sFilter;
            //openFileDialog1.Filter = "CSVファイル(*.CSV)|*.csv|全てのファイル(*.*)|*.*";

            //ダイアログボックスの表示
            ret = openFileDialog1.ShowDialog();
            if (ret == System.Windows.Forms.DialogResult.Cancel)
            {
                return string.Empty;
            }

            if (MessageBox.Show(openFileDialog1.FileName + Environment.NewLine + " が選択されました。よろしいですか?", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return string.Empty;
            }

            return openFileDialog1.FileName;
        }

        public class frmMode
        {
            public int ID { get; set; }

            public int Mode { get; set; }

            public int rowIndex { get; set; }
        }

        public class xlsShain
        {
            public int sCode { get; set; }
            public string sName { get; set; }
            public int bCode { get; set; }
            public string bName { get; set; }
        }
        
        ///----------------------------------------------------------------------------
        /// <summary>
        ///     CSVファイルを追加モードで出力する</summary>
        /// <param name="sPath">
        ///     出力するパス</param>
        /// <param name="arrayData">
        ///     書き込む配列データ</param>
        /// <param name="sFileName">
        ///     CSVファイル名</param>
        ///----------------------------------------------------------------------------
        public static void csvFileWrite(string sPath, string[] arrayData, string sFileName)
        {
            //// ファイル名（タイムスタンプ）
            //string timeStamp = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') +
            //                     DateTime.Now.Day.ToString().PadLeft(2, '0') + DateTime.Now.Hour.ToString().PadLeft(2, '0') +
            //                     DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0');

            //// ファイル名
            //string outFileName = sPath + timeStamp + ".csv";

            //// 出力ファイルが存在するとき
            //if (System.IO.File.Exists(outFileName))
            //{
            //    // 既存のファイルを削除
            //    System.IO.File.Delete(outFileName);
            //}

            // CSVファイル出力
            //System.IO.File.WriteAllLines(outFileName, arrayData, System.Text.Encoding.GetEncoding("shift-jis"));
            System.IO.File.AppendAllLines(sPath, arrayData, Encoding.GetEncoding("shift-Jis"));
        }
        
        ///---------------------------------------------------------------------
        /// <summary>
        ///     任意のディレクトリのファイルを削除する </summary>
        /// <param name="sPath">
        ///     指定するディレクトリ</param>
        /// <param name="sFileType">
        ///     ファイル名及び形式</param>
        /// --------------------------------------------------------------------
        public static void FileDelete(string sPath, string sFileType)
        {
            //sFileTypeワイルドカード"*"は、すべてのファイルを意味する
            foreach (string files in System.IO.Directory.GetFiles(sPath, sFileType))
            {
                // ファイルを削除する
                System.IO.File.Delete(files);
            }
        }


        // 勘定奉行データベース接続
        public class SQLDBConnect
        {
            SqlConnection cn = new SqlConnection();

            public SqlConnection Cn
            {
                get
                {
                    return cn;
                }
            }

            /// <summary>
            /// SQLServerへ接続
            /// </summary>
            /// <param name="sConnect">接続文字列</param>
            public SQLDBConnect(string sConnect)
            {
                try
                {
                    // データベース接続文字列
                    cn.ConnectionString = sConnect;
                    cn.Open();
                }

                catch (Exception e)
                {
                    throw e;
                }
            }
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     文字列を指定文字数をＭＡＸとして返します</summary>
        /// <param name="s">
        ///     文字列</param>
        /// <param name="n">
        ///     文字数</param>
        /// <returns>
        ///     文字数範囲内の文字列</returns>
        /// --------------------------------------------------------------------
        public static string GetStringSubMax(string s, int n)
        {
            string val = string.Empty;

            // 文字間のスペースを除去 2015/03/10
            s = s.Replace(" ", "");

            if (s.Length > n) val = s.Substring(0, n);
            else val = s;

            return val;
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     ライン・部門・製品群コード配列取得   </summary>
        /// <returns>
        ///     ID,コード配列</returns>
        ///-------------------------------------------------------------------
        public static string[] getCategoryArray(string dbName)
        {
            // 接続文字列取得
            string sc = sqlControl.obcConnectSting.get(dbName);
            sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            StringBuilder sb = new StringBuilder();
            sb.Append("select DivisionID, CategoryID, CategoryCode from tbHistoryDivisionCategory");
            SqlDataReader dr = sdCon.free_dsReader(sb.ToString());

            int iX = 0;
            string[] hArray = new string[1];

            while (dr.Read())
            {
                if (iX > 0)
                {
                    Array.Resize(ref hArray, iX + 1);
                }

                hArray[iX] = dr["CategoryID"].ToString() + "," + dr["CategoryCode"].ToString() + "," + dr["DivisionID"].ToString();
                iX++;
            }

            dr.Close();
            sdCon.Close();

            return hArray;
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     社員情報抽出ＳＱＬ作成 </summary>
        /// <param name="bCode">
        ///     社員コード</param>
        /// <returns>
        ///     ＳＱＬ文字列</returns>
        ///---------------------------------------------------------------
        public static string getEmployee(string bCode)
        {
            string dt = DateTime.Today.ToShortDateString();

            // 社員情報抽出ＳＱＬ
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT tbEmployeeBase.EmployeeID, tbHR_DivisionCategory.CategoryCode as zaisekikbn,");
            sb.Append("tbEmployeeBase.EmployeeNo, tbEmployeeBase.NameKana, tbEmployeeBase.Name,");
            sb.Append("tbDepartment.DepartmentID, right(replace(tbDepartment.DepartmentCode, ' ', ''), 5) as DepartmentCode, tbDepartment.DepartmentName,");
            sb.Append("tbEmployeeBase.RetireCorpScheduleDate, d.JobTypeID, d.DutyID, d.QualificationGradeID ");

            sb.Append("from(((tbEmployeeBase inner join ");
            sb.Append("(select tbEmployeeMainDutyPersonnelChange.EmployeeID, tbEmployeeMainDutyPersonnelChange.AnnounceDate,");
            sb.Append("tbEmployeeMainDutyPersonnelChange.BelongID, tbEmployeeMainDutyPersonnelChange.DutyID,");
            sb.Append("tbEmployeeMainDutyPersonnelChange.JobTypeID, tbEmployeeMainDutyPersonnelChange.QualificationGradeID ");

            sb.Append("from tbEmployeeMainDutyPersonnelChange inner join ");

            sb.Append("(select EmployeeID, max(AnnounceDate) as AnnounceDate from tbEmployeeMainDutyPersonnelChange ");
            sb.Append("where AnnounceDate <= '").Append(DateTime.Today.ToShortDateString()).Append("' ");
            sb.Append("group by EmployeeID) as a ");
            sb.Append("on(tbEmployeeMainDutyPersonnelChange.EmployeeID = a.EmployeeID) and ");
            sb.Append("(tbEmployeeMainDutyPersonnelChange.AnnounceDate = a.AnnounceDate) ");
            sb.Append(") as d ");
            sb.Append("on tbEmployeeBase.EmployeeID = d.EmployeeID) ");

            sb.Append("inner join tbDepartment on d.BelongID = tbDepartment.DepartmentID) ");
            sb.Append("inner join tbHR_DivisionCategory on tbEmployeeBase.BeOnTheRegisterDivisionID = tbHR_DivisionCategory.CategoryID) ");

            //sb.Append("where EmployeeNo = '" + bCode + "' and tbHR_DivisionCategory.CategoryCode <> 2 "); // 2017/05/08 

            // 在籍区分 <> 2 を外した : 2017/09/28　
            sb.Append("where EmployeeNo = '" + bCode + "' ");
            sb.Append("ORDER BY DepartmentCode,tbEmployeeBase.EmployeeNo");

            return sb.ToString();
        }

        ///-----------------------------------------------------------------------
        /// <summary>
        ///     ライン・部門・製品群コード取得　</summary>
        /// <param name="hArray">
        ///     配列</param>
        /// <param name="sCode">
        ///     CategoryID</param>
        /// <returns>
        ///     CategoryCode</returns>
        ///-----------------------------------------------------------------------
        public static string getHisCategory(string[] hArray, string sCode)
        {
            string rtnCode = "";

            foreach (var t in hArray)
            {
                string[] n = t.Split(',');

                if (n[0].ToString() == sCode)
                {
                    rtnCode = n[1];
                    break;
                }
            }

            return rtnCode.Trim();
        }

        public static bool getHisCategory(string[] hArray, string sCode, string divID)
        {
            bool rtn = false;

            foreach (var t in hArray)
            {
                string[] n = t.Split(',');

                if (n[1].ToString() == sCode && n[2].ToString() == divID)
                {
                    rtn = true;
                    break;
                }
            }

            return rtn;
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
        public static bool chkJiyu(string s, string _dbName)
        {
            bool dm = false;

            // 奉行SQLServer接続文字列取得
            string sc = sqlControl.obcConnectSting.get(_dbName);
            sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

            // 登録済み事由コード検証
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select LaborReasonCode from tbLaborReason ");
            sb.Append("where IsValid = 1 and LaborReasonCode = '" + s.PadLeft(2, '0') + "'");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                dm = true;
                break;
            }

            dR.Close();
            sdCon.Close();

            return dm;
        }

        /////------------------------------------------------------------
        ///// <summary>
        /////     「終日」事由と他の事由併記チェック </summary>
        ///// <param name="s">
        /////     事由配列</param>
        ///// <param name="_dbName">
        /////     データベース名</param>
        ///// <returns>
        /////     true:エラーなし、false:エラー</returns>
        /////------------------------------------------------------------
        //public static bool chkJiyu(string [] s, string _dbName)
        //{
        //    bool dm = false;

        //    // 奉行SQLServer接続文字列取得
        //    string sc = sqlControl.obcConnectSting.get(_dbName);
        //    sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

        //    // 登録済み事由コード検証
        //    StringBuilder sb = new StringBuilder();

        //    for (int i = 0; i < 3; i++)
        //    {
        //        if (s[i].Trim() == string.Empty)
        //        {
        //            continue;
        //        }

        //        sb.Clear();
        //        sb.Append("select LaborReasonCode,AcquireUnit from tbLaborReason ");
        //        sb.Append("where IsValid = 1 and LaborReasonCode = '" + s[i].PadLeft(2, '0') + "'");

        //        SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

        //        while (dR.Read())
        //        {
        //            // 取得区分
        //            if (Utility.NulltoStr(dR["AcquireUnit"]) == global.FLGOFF)
        //            {
        //                // 終日「０
        //                dm = true;  // 終日あり
        //            }

        //            break;
        //        }

        //        dR.Close();
        //    }

        //    sdCon.Close();

        //    // 終日事由があり、他の事由が併記されているときはエラ―
        //    if (!dm)
        //    {
        //        // 終日がない場合戻る
        //        return true;
        //    }
        //    else
        //    {
        //        int cnt = 0;
        //        for (int i = 0; i < 3; i++)
        //        {
        //            if (s[i] != string.Empty)
        //            {
        //                // 事由記入あり
        //                cnt++;
        //            }
        //        }

        //        if (cnt > 1)
        //        {
        //            return false;
        //        }
        //        else
        //        {
        //            return true;
        //        }
        //    }
        //}

        /////------------------------------------------------------------
        ///// <summary>
        /////     「終日」事由と休出シフトの記入チェック </summary>
        ///// <param name="s">
        /////     事由配列</param>
        ///// <param name="_dbName">
        /////     データベース名</param>
        ///// <param name="r">
        /////     DataSet1.勤務票ヘッダRow </param>
        ///// <param name="m">
        /////     DataSet1.勤務票明細Row</param>
        ///// <returns>
        /////     true:エラーなし、false:エラー</returns>
        /////------------------------------------------------------------
        //public static bool chkJiyu(string[] s, string _dbName, DataSet1.勤務票ヘッダRow r, DataSet1.勤務票明細Row m)
        //{
        //    bool dm = false;

        //    // 奉行SQLServer接続文字列取得
        //    string sc = sqlControl.obcConnectSting.get(_dbName);
        //    sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

        //    // 事由コード取得
        //    StringBuilder sb = new StringBuilder();

        //    for (int i = 0; i < 3; i++)
        //    {
        //        if (s[i].Trim() == string.Empty)
        //        {
        //            continue;
        //        }

        //        sb.Clear();
        //        sb.Append("select LaborReasonCode,AcquireUnit from tbLaborReason ");
        //        sb.Append("where IsValid = 1 and LaborReasonCode = '" + s[i].PadLeft(2, '0') + "'");

        //        SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

        //        while (dR.Read())
        //        {
        //            // 取得区分
        //            if (Utility.NulltoStr(dR["AcquireUnit"]) == global.FLGOFF)
        //            {
        //                // 終日「０」
        //                dm = true;  // 終日あり
        //            }

        //            break;
        //        }

        //        dR.Close();
        //    }

        //    sdCon.Close();

        //    if (!dm)
        //    {
        //        // 終日がない場合戻る
        //        return true;
        //    }
        //    else
        //    {
        //        // 休日出勤のときエラー
        //        if (Utility.StrtoInt(m.シフトコード) == global.SFT_KYUSHUTSU || r.シフトコード == global.SFT_KYUSHUTSU)
        //        {
        //            return false;
        //        }
        //        else
        //        {
        //            return true;
        //        }
        //    }
        //}

        /////------------------------------------------------------------
        ///// <summary>
        /////     「終日」事由とシフト以外の記入チェック </summary>
        ///// <param name="s">
        /////     事由配列</param>
        ///// <param name="_dbName">
        /////     データベース名</param>
        ///// <param name="m">
        /////     DataSet1.勤務票明細Row</param>
        ///// <returns>
        /////     true:エラーなし、false:エラー</returns>
        /////------------------------------------------------------------
        //public static bool chkJiyu(string[] s, string _dbName, DataSet1.勤務票明細Row m, out int eNum)
        //{
        //    bool dm = false;
        //    eNum = 0; 

        //    // 奉行SQLServer接続文字列取得
        //    string sc = sqlControl.obcConnectSting.get(_dbName);
        //    sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

        //    // 登録済み事由コード検証
        //    StringBuilder sb = new StringBuilder();

        //    for (int i = 0; i < 3; i++)
        //    {
        //        if (s[i].Trim() == string.Empty)
        //        {
        //            continue;
        //        }

        //        sb.Clear();
        //        sb.Append("select LaborReasonCode,AcquireUnit from tbLaborReason ");
        //        sb.Append("where IsValid = 1 and LaborReasonCode = '" + s[i].PadLeft(2, '0') + "'");

        //        SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

        //        while (dR.Read())
        //        {
        //            // 取得区分
        //            if (Utility.NulltoStr(dR["AcquireUnit"]) == global.FLGOFF)
        //            {
        //                // 終日「０」
        //                dm = true;  // 終日あり
        //            }

        //            break;
        //        }

        //        dR.Close();
        //    }

        //    sdCon.Close();

        //    if (!dm)
        //    {
        //        // 終日がない場合戻る
        //        return true;
        //    }
        //    else
        //    {
        //        // 終日でシフトコード以外記入があるとき
        //        if (m.出勤時 != string.Empty)
        //        {
        //            eNum = 1;
        //            return false;
        //        }
        //        if (m.出勤分 != string.Empty)
        //        {
        //            eNum = 2;
        //            return false;
        //        }

        //        if (m.退勤時 != string.Empty)
        //        {
        //            eNum = 3;
        //            return false;
        //        }

        //        if (m.退勤分 != string.Empty)
        //        {
        //            eNum = 4;
        //            return false;
        //        }

        //        if (m.残業理由1 != string.Empty)
        //        {
        //            eNum = 5;
        //            return false;
        //        }

        //        if (m.残業時1 != string.Empty)
        //        {
        //            eNum = 6;
        //            return false;
        //        }

        //        if (m.残業分1 != string.Empty)
        //        {
        //            eNum = 7;
        //            return false;
        //        }

        //        if (m.残業理由2 != string.Empty)
        //        {
        //            eNum = 8;
        //            return false;
        //        }

        //        if (m.残業時2 != string.Empty)
        //        {
        //            eNum = 9;
        //            return false;
        //        }

        //        if (m.残業分2 != string.Empty)
        //        {
        //            eNum = 10;
        //            return false;
        //        }

        //        if (m.応援 == global.FLGON)
        //        {
        //            eNum = 11;
        //            return false;
        //        }

        //        if (m.シフトコード != string.Empty)
        //        {
        //            eNum = 12;
        //            return false;
        //        }

        //        return true;


        //        //// 終日でシフトコード以外記入があるとき
        //        //if (m.出勤時 != string.Empty || m.出勤分 != string.Empty || m.退勤時 != string.Empty || m.退勤分 != string.Empty ||
        //        //    m.残業理由1 != string.Empty || m.残業時1 != string.Empty || m.残業分1 != string.Empty ||
        //        //    m.残業理由2 != string.Empty || m.残業時2 != string.Empty || m.残業分2 != string.Empty ||
        //        //    m.応援 == global.FLGON)
        //        //{
        //        //    return false;
        //        //}
        //        //else
        //        //{
        //        //    return true;
        //        //}
        //    }
        //}

        /////------------------------------------------------------------
        ///// <summary>
        /////     取得単位「半日」事由の取得区分の重複記入チェック </summary>
        ///// <param name="s">
        /////     事由配列</param>
        ///// <param name="_dbName">
        /////     データベース名</param>
        ///// <param name="m">
        /////     DataSet1.勤務票明細Row</param>
        ///// <returns>
        /////     true:エラーなし、false:エラー</returns>
        /////------------------------------------------------------------
        //public static bool chkJiyuDiv(string[] s, string _dbName, out int dCnt)
        //{
        //    dCnt = 0;
        //    bool dm = false;
        //    string[] div = new string[3];

        //    // 奉行SQLServer接続文字列取得
        //    string sc = sqlControl.obcConnectSting.get(_dbName);
        //    sqlControl.DataControl sdCon = new sqlControl.DataControl(sc);

        //    // 登録済み事由コード検証
        //    StringBuilder sb = new StringBuilder();

        //    for (int i = 0; i < 3; i++)
        //    {
        //        if (s[i].Trim() == string.Empty)
        //        {
        //            div[i] = string.Empty;
        //            continue;
        //        }

        //        sb.Clear();
        //        sb.Append("select LaborReasonCode,AcquireUnit,AcquireDivision from tbLaborReason ");
        //        sb.Append("where IsValid = 1 and LaborReasonCode = '" + s[i].PadLeft(2, '0') + "'");

        //        SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

        //        while (dR.Read())
        //        {
        //            // 取得単位
        //            if (Utility.NulltoStr(dR["AcquireUnit"]) == global.FLGON)
        //            {
        //                // 半日「1」
        //                dm = true;  // 終日あり
        //                div[i] = Utility.NulltoStr(dR["AcquireDivision"]); // 取得区分
        //            }

        //            break;
        //        }

        //        dR.Close();
        //    }

        //    sdCon.Close();

        //    if (!dm)
        //    {
        //        // 半日がない場合戻る
        //        return true;
        //    }
        //    else
        //    {
        //        int kbn = 0;

        //        for (int i = 0; i < 3; i++)
        //        {
        //            if (div[i] == string.Empty)
        //            {
        //                continue;
        //            }

        //            kbn += Utility.StrtoInt(div[i]);
        //            dCnt++;
        //        }

        //        if (dCnt == 3)
        //        {
        //            // 半日事由が3つ記入されている
        //            return false;
        //        }
        //        else if (dCnt == 2)
        //        {
        //            // 半日事由が2つ記入されている
        //            if (kbn != 1)
        //            {
        //                // 前半(0)・前半(0)、または後半(1)・後半(1)の組み合わせになっている
        //                return false;
        //            }
        //            else
        //            {
        //                return true;
        //            }
        //        }
        //        else
        //        {
        //            return true;
        //        }
        //    }
        //}

        ///------------------------------------------------------------------------------------
        /// <summary>
        ///     時間記入範囲チェック 0～23の数値 </summary>
        /// <param name="h">
        ///     記入値</param>
        /// <returns>
        ///     正常:true, エラー:false</returns>
        ///------------------------------------------------------------------------------------
        public static bool checkHourSpan(string h)
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
        public static bool checkMinSpan(string m, int tani)
        {
            if (!Utility.NumericCheck(m)) return false;
            else if (int.Parse(m) < 0 || int.Parse(m) > 59) return false;
            else if (int.Parse(m) % tani != 0) return false;
            else return true;
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
        public static bool chkZangyoRe(string zanRe, string zH, string zM)
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
        public static bool chkZangyoRe2(string zanRe, string zH, string zM)
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

        ///------------------------------------------------------------
        /// <summary>
        ///     部署名取得 </summary>
        /// <param name="sdCon">
        ///     sqlControl.DataControl オブジェクト </param>
        /// <param name="s">
        ///     部署コード</param>
        /// <returns>
        ///     部署名</returns>
        ///------------------------------------------------------------
        public static string getDepartmentName(sqlControl.DataControl sdCon, string s)
        {
            string dName = string.Empty;

            // 登録済み事由コード検証
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select DepartmentName from tbDepartment ");
            sb.Append("where DepartmentCode = '" + s + "'");

            SqlDataReader dR = sdCon.free_dsReader(sb.ToString());

            while (dR.Read())
            {
                dName = dR["DepartmentName"].ToString();
                break;
            }

            dR.Close();

            return dName;
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     社員情報抽出ＳＱＬ作成 </summary>
        /// <param name="bCode">
        ///     部署コード</param>
        /// <param name="sDt">
        ///     基準年月日 : 2017/09/28</param>
        /// <returns>
        ///     ＳＱＬ文字列</returns>
        ///---------------------------------------------------------------
        public static string getEmployeeCount(string bCode, DateTime sDt)
        {
            string dt = DateTime.Today.ToShortDateString();

            // 社員情報抽出ＳＱＬ
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT count(tbEmployeeBase.EmployeeID) as cnt ");

            sb.Append("from(((tbEmployeeBase inner join ");
            sb.Append("(select tbEmployeeMainDutyPersonnelChange.EmployeeID, tbEmployeeMainDutyPersonnelChange.AnnounceDate,");
            sb.Append("tbEmployeeMainDutyPersonnelChange.BelongID, tbEmployeeMainDutyPersonnelChange.DutyID,");
            sb.Append("tbEmployeeMainDutyPersonnelChange.JobTypeID, tbEmployeeMainDutyPersonnelChange.QualificationGradeID ");

            sb.Append("from tbEmployeeMainDutyPersonnelChange inner join ");

            sb.Append("(select EmployeeID, max(AnnounceDate) as AnnounceDate from tbEmployeeMainDutyPersonnelChange ");
            sb.Append("where AnnounceDate <= '").Append(sDt.ToShortDateString()).Append("' ");
            sb.Append("group by EmployeeID) as a ");
            sb.Append("on(tbEmployeeMainDutyPersonnelChange.EmployeeID = a.EmployeeID) and ");
            sb.Append("(tbEmployeeMainDutyPersonnelChange.AnnounceDate = a.AnnounceDate) ");
            sb.Append(") as d ");
            sb.Append("on tbEmployeeBase.EmployeeID = d.EmployeeID) ");

            sb.Append("inner join tbDepartment on d.BelongID = tbDepartment.DepartmentID) ");
            sb.Append("inner join tbHR_DivisionCategory on tbEmployeeBase.BeOnTheRegisterDivisionID = tbHR_DivisionCategory.CategoryID) ");
           
            //sb.Append("where DepartmentCode = '" + bCode + "' and tbHR_DivisionCategory.CategoryCode <> 2"); // 2017/05/08 
            
            // 在籍区分 <> 2 を外し、入社年月日と退職年月日で判断 : 2017/09/28　
            sb.Append("where DepartmentCode = '" + bCode + "' ");
            sb.Append("and EnterCorpDate <= '" + sDt.ToShortDateString() + "' and RetireCorpDate >= '" + sDt.ToShortDateString() + "' ");
            
            return sb.ToString();
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     検索用DepartmentCodeを取得する </summary>
        /// <returns>
        ///     DepartmentCode</returns>
        ///----------------------------------------------------------
        public static string getDepartmentCode(string bCode)
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

        ///---------------------------------------------------------------
        /// <summary>
        ///     社員情報抽出ＳＱＬ作成 </summary>
        /// <param name="bCode">
        ///     部署コード</param>
        /// <param name="sDt">
        ///     基準年月日</param>
        /// <returns>
        ///     ＳＱＬ文字列</returns>
        ///---------------------------------------------------------------
        public static string getEmployeeOrder(string bCode, DateTime sDt)
        {
            string dt = DateTime.Today.ToShortDateString();

            // 社員情報抽出ＳＱＬ
            StringBuilder sb = new StringBuilder();

            //sb.Append("SELECT tbEmployeeBase.EmployeeID, tbHR_DivisionCategory.CategoryCode as zaisekikbn,");
            //sb.Append("tbEmployeeBase.EmployeeNo, tbEmployeeBase.NameKana, tbEmployeeBase.Name,");
            //sb.Append("tbDepartment.DepartmentID, right(replace(tbDepartment.DepartmentCode, ' ', ''), 5) as DepartmentCode, tbDepartment.DepartmentName,");
            //sb.Append("tbEmployeeBase.RetireCorpScheduleDate, d.JobTypeID, d.DutyID, d.QualificationGradeID ");

            //sb.Append("from(((tbEmployeeBase inner join ");
            //sb.Append("(select tbEmployeeMainDutyPersonnelChange.EmployeeID, tbEmployeeMainDutyPersonnelChange.AnnounceDate,");
            //sb.Append("tbEmployeeMainDutyPersonnelChange.BelongID, tbEmployeeMainDutyPersonnelChange.DutyID,");
            //sb.Append("tbEmployeeMainDutyPersonnelChange.JobTypeID, tbEmployeeMainDutyPersonnelChange.QualificationGradeID ");

            //sb.Append("from tbEmployeeMainDutyPersonnelChange inner join ");

            //sb.Append("(select EmployeeID, max(AnnounceDate) as AnnounceDate from tbEmployeeMainDutyPersonnelChange ");
            //sb.Append("where AnnounceDate <= '").Append(DateTime.Today.ToShortDateString()).Append("' ");
            //sb.Append("group by EmployeeID) as a ");
            //sb.Append("on(tbEmployeeMainDutyPersonnelChange.EmployeeID = a.EmployeeID) and ");
            //sb.Append("(tbEmployeeMainDutyPersonnelChange.AnnounceDate = a.AnnounceDate) ");
            //sb.Append(") as d ");
            //sb.Append("on tbEmployeeBase.EmployeeID = d.EmployeeID) ");

            //sb.Append("inner join tbDepartment on d.BelongID = tbDepartment.DepartmentID) ");
            //sb.Append("inner join tbHR_DivisionCategory on tbEmployeeBase.BeOnTheRegisterDivisionID = tbHR_DivisionCategory.CategoryID) ");
            //sb.Append("where DepartmentCode = '" + bCode + "' and tbHR_DivisionCategory.CategoryCode <> 2 "); // 2017/05/08 
            //sb.Append("ORDER BY DepartmentCode,tbEmployeeBase.EmployeeNo");


            // 印字順：役職→ライン別→雇用区分別→性別→社員番号　2017/05/26
            // 印字順：役職→ライン別→雇用区分別→性別→連番　2017/09/20
            sb.Append("SELECT tbEmployeeBase.EmployeeID, tbHR_DivisionCategory.CategoryCode as zaisekikbn,");
            sb.Append("tbEmployeeBase.EmployeeNo, tbEmployeeBase.NameKana, tbEmployeeBase.Name,");
            sb.Append("tbDepartment.DepartmentID, right(replace(tbDepartment.DepartmentCode, ' ', ''), 5) as DepartmentCode, tbDepartment.DepartmentName,");
            sb.Append("tbEmployeeBase.RetireCorpScheduleDate, dc.CategoryCode as ManagerCode, d.JobTypeID, dc2.CategoryCode as LineCode, hrdc.CategoryCode as koyou, ");
            sb.Append("d.DutyID, d.QualificationGradeID, d.ManagerialPostID, dc3.CategoryCode ");

            sb.Append("from((((((((tbEmployeeBase inner join ");
            sb.Append("(select tbEmployeeMainDutyPersonnelChange.EmployeeID, tbEmployeeMainDutyPersonnelChange.AnnounceDate,");
            sb.Append("tbEmployeeMainDutyPersonnelChange.BelongID, tbEmployeeMainDutyPersonnelChange.DutyID,");
            sb.Append("tbEmployeeMainDutyPersonnelChange.JobTypeID, tbEmployeeMainDutyPersonnelChange.QualificationGradeID, ");
            sb.Append("tbEmployeeMainDutyPersonnelChange.ManagerialPostID ");

            sb.Append("from tbEmployeeMainDutyPersonnelChange inner join ");

            sb.Append("(select EmployeeID, max(AnnounceDate) as AnnounceDate from tbEmployeeMainDutyPersonnelChange ");
            //sb.Append("where AnnounceDate <= '").Append(DateTime.Today.ToShortDateString()).Append("' ");
            sb.Append("where AnnounceDate <= '").Append(sDt.ToShortDateString()).Append("' ");
            sb.Append("group by EmployeeID) as a ");
            sb.Append("on(tbEmployeeMainDutyPersonnelChange.EmployeeID = a.EmployeeID) and ");
            sb.Append("(tbEmployeeMainDutyPersonnelChange.AnnounceDate = a.AnnounceDate) ");
            sb.Append(") as d ");
            sb.Append("on tbEmployeeBase.EmployeeID = d.EmployeeID) ");

            sb.Append("inner join tbDepartment on d.BelongID = tbDepartment.DepartmentID) ");
            sb.Append("inner join tbHR_DivisionCategory on tbEmployeeBase.BeOnTheRegisterDivisionID = tbHR_DivisionCategory.CategoryID) ");
            sb.Append("inner join tbHistoryDivisionCategory as dc on d.ManagerialPostID = dc.CategoryID) ");
            sb.Append("inner join tbHistoryDivisionCategory as dc2 on d.JobTypeID = dc2.CategoryID) ");
            sb.Append("inner join tbHR_DivisionCategory as hrdc on tbEmployeeBase.EmploymentDivisionID = hrdc.CategoryID) ");
            sb.Append("inner join tbEmployeeDivision on tbEmployeeBase.EmployeeID = tbEmployeeDivision.EmployeeID) ");   // 2017/09/20 
            sb.Append("inner join tbHR_DivisionCategory as dc3 on tbEmployeeDivision.Division04ID =  dc3.CategoryID) ");     // 2017/09/20 

            //sb.Append("where DepartmentCode = '" + bCode + "' and tbHR_DivisionCategory.CategoryCode <> 2 "); // 2017/05/08

            // 在籍区分 <> 2 を外し、入社年月日と退職年月日で判断 : 2017/09/28　
            sb.Append("where DepartmentCode = '" + bCode + "' ");
            sb.Append("and EnterCorpDate <= '" + sDt.ToShortDateString() + "' and RetireCorpDate >= '" + sDt.ToShortDateString() + "' "); 
            sb.Append("ORDER BY DepartmentCode, ManagerCode, LineCode, koyou, tbEmployeeBase.SexID, dc3.CategoryCode");     // 2017/09/20:ソートの最後を社員番号→連番

            return sb.ToString();
        }



        ///---------------------------------------------------------------
        /// <summary>
        ///     社員情報抽出ＳＱＬ作成 : 勤怠表</summary>
        /// <param name="bCode">
        ///     部署コード</param>
        /// <param name="sDt">
        ///     基準年月日</param>
        /// <param name="rDt">
        ///     退職基準年月日</param>
        /// <returns>
        ///     ＳＱＬ文字列</returns>
        ///---------------------------------------------------------------
        public static string getEmployeeKintaiRep(string bCode, DateTime sDt, DateTime rDt)
        {

            string dt = DateTime.Today.ToShortDateString();

            // 社員情報抽出ＳＱＬ
            StringBuilder sb = new StringBuilder();
            // 印字順：役職→ライン別→雇用区分別→性別→社員番号　2017/05/26
            // 印字順：役職→ライン別→雇用区分別→性別→連番　2017/09/20
            sb.Append("SELECT tbEmployeeBase.EmployeeID, tbHR_DivisionCategory.CategoryCode as zaisekikbn,");
            sb.Append("tbEmployeeBase.EmployeeNo, tbEmployeeBase.NameKana, tbEmployeeBase.Name,");
            sb.Append("tbDepartment.DepartmentID, right(replace(tbDepartment.DepartmentCode, ' ', ''), 5) as DepartmentCode, tbDepartment.DepartmentName,");
            sb.Append("tbEmployeeBase.RetireCorpScheduleDate, dc.CategoryCode as ManagerCode, d.JobTypeID, dc2.CategoryCode as LineCode, hrdc.CategoryCode as koyou, ");
            sb.Append("d.DutyID, d.QualificationGradeID, d.ManagerialPostID, dc3.CategoryCode ");

            sb.Append("from((((((((tbEmployeeBase inner join ");
            sb.Append("(select tbEmployeeMainDutyPersonnelChange.EmployeeID, tbEmployeeMainDutyPersonnelChange.AnnounceDate,");
            sb.Append("tbEmployeeMainDutyPersonnelChange.BelongID, tbEmployeeMainDutyPersonnelChange.DutyID,");
            sb.Append("tbEmployeeMainDutyPersonnelChange.JobTypeID, tbEmployeeMainDutyPersonnelChange.QualificationGradeID, ");
            sb.Append("tbEmployeeMainDutyPersonnelChange.ManagerialPostID ");

            sb.Append("from tbEmployeeMainDutyPersonnelChange inner join ");

            sb.Append("(select EmployeeID, max(AnnounceDate) as AnnounceDate from tbEmployeeMainDutyPersonnelChange ");
            //sb.Append("where AnnounceDate <= '").Append(DateTime.Today.ToShortDateString()).Append("' ");
            sb.Append("where AnnounceDate <= '").Append(sDt.ToShortDateString()).Append("' ");
            sb.Append("group by EmployeeID) as a ");
            sb.Append("on(tbEmployeeMainDutyPersonnelChange.EmployeeID = a.EmployeeID) and ");
            sb.Append("(tbEmployeeMainDutyPersonnelChange.AnnounceDate = a.AnnounceDate) ");
            sb.Append(") as d ");
            sb.Append("on tbEmployeeBase.EmployeeID = d.EmployeeID) ");

            sb.Append("inner join tbDepartment on d.BelongID = tbDepartment.DepartmentID) ");
            sb.Append("inner join tbHR_DivisionCategory on tbEmployeeBase.BeOnTheRegisterDivisionID = tbHR_DivisionCategory.CategoryID) ");
            sb.Append("inner join tbHistoryDivisionCategory as dc on d.ManagerialPostID = dc.CategoryID) ");
            sb.Append("inner join tbHistoryDivisionCategory as dc2 on d.JobTypeID = dc2.CategoryID) ");
            sb.Append("inner join tbHR_DivisionCategory as hrdc on tbEmployeeBase.EmploymentDivisionID = hrdc.CategoryID) ");
            sb.Append("inner join tbEmployeeDivision on tbEmployeeBase.EmployeeID = tbEmployeeDivision.EmployeeID) ");   // 2017/09/20 
            sb.Append("inner join tbHR_DivisionCategory as dc3 on tbEmployeeDivision.Division04ID =  dc3.CategoryID) ");     // 2017/09/20 

            //sb.Append("where DepartmentCode = '" + bCode + "' and tbHR_DivisionCategory.CategoryCode <> 2 "); // 2017/05/08

            // 在籍区分 <> 2 を外し、入社年月日と退職年月日で判断 : 2017/09/28
　          // 退職基準日パラメータで判断    2018/02/19
            sb.Append("where DepartmentCode = '" + bCode + "' ");
            sb.Append("and EnterCorpDate <= '" + sDt.ToShortDateString() + "' and RetireCorpDate >= '" + rDt.ToShortDateString() + "' ");
            sb.Append("ORDER BY DepartmentCode, ManagerCode, LineCode, koyou, tbEmployeeBase.SexID, dc3.CategoryCode");     // 2017/09/20:ソートの最後を社員番号→連番

            return sb.ToString();
        }

        ///--------------------------------------------------------------------------------------
        /// <summary>
        ///     社員名なしの過去勤務票明細、過去応援移動票明細に社員名をセットする：2018/03/22</summary>
        /// <param name="dbName">
        ///     会社領域データベース名</param>
        ///--------------------------------------------------------------------------------------
        public static void getNoNameRecovery(string dbName)
        {
            DataSet1 dts = new DataSet1();

            // 接続文字列取得 2018/03/22
            string sc = sqlControl.obcConnectSting.get(dbName);
            sqlControl.DataControl sdCon = new Common.sqlControl.DataControl(sc);

            DataSet1TableAdapters.過去勤務票明細TableAdapter kAdp = new DataSet1TableAdapters.過去勤務票明細TableAdapter();
            DataSet1TableAdapters.過去応援移動票明細TableAdapter uAdp = new DataSet1TableAdapters.過去応援移動票明細TableAdapter();
            SqlDataReader dR = null;

            try
            {
                // 社員名なしの過去勤務票明細データに社員名をセットする 2018/03/22
                kAdp.FillByNoName(dts.過去勤務票明細);

                foreach (var nn in dts.過去勤務票明細)
                {
                    string bCode = nn.社員番号.PadLeft(10, '0');
                    dR = sdCon.free_dsReader(Utility.getEmployee(bCode));

                    while (dR.Read())
                    {
                        // 社員名セット 2018/03/22
                        nn.社員名 = dR["Name"].ToString().Trim();
                    }

                    kAdp.Update(dts.過去勤務票明細);
                    dR.Close();
                }

                // 社員名なしの過去応援移動票明細に社員名をセットする 2018/03/22
                uAdp.FillByNoName(dts.過去応援移動票明細);

                foreach (var uu in dts.過去応援移動票明細)
                {
                    string bCode = uu.社員番号.PadLeft(10, '0');
                    dR = sdCon.free_dsReader(Utility.getEmployee(bCode));

                    while (dR.Read())
                    {
                        // 社員名セット 2018/03/22
                        uu.社員名 = dR["Name"].ToString().Trim();
                    }

                    uAdp.Update(dts.過去応援移動票明細);
                    dR.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (dR != null && !dR.IsClosed)
                {
                    dR.Close();
                }

                if (sdCon.Cn.State == System.Data.ConnectionState.Open)
                {
                    sdCon.Close();
                }
            }
        }
    }
}
