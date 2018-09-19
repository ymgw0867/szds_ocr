using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace SZDS_TIMECARD.Common
{
    public class xlsData
    {
        // 部署別勤務体系配列
        public object[,] zArray = null;

        // 部署別残業理由配列
        public object[,] rArray = null;

        // 部署別残業計画配列
        public object[,] zpArray = null;
        
        // 部署別理由別残業計画配列
        public object[,] zrpArray = null;

        ///----------------------------------------------------------------------
        /// <summary>
        ///     部署別勤務体系シートよりデータを配列に取得する </summary>
        /// <param name="sPath">
        ///     任意指定した部署別勤務体系シートパス</param>
        /// <returns>
        ///     取得した配列：object</returns>
        ///----------------------------------------------------------------------
        public object[,] getShiftCode(string sPath)
        {
            // 2017/08/30
            string xlsPath = string.Empty;
            if (sPath == string.Empty)
            {
                xlsPath = Properties.Settings.Default.xlsBmnShift;
            }
            else
            {
                xlsPath = sPath;
            }

            object[,] rtnArray = null;

            string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

            Excel.Application oXls = new Excel.Application();

            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(xlsPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];
            oxlsSheet.Select(Type.Missing);

            Excel.Range rng = null;

            try
            {
                rng = oxlsSheet.Range[oxlsSheet.Cells[2, 1], oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 6]];
                rtnArray = (object[,])rng.Value2;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "部署別勤務体系シート取り込みエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                // Bookをクローズ
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                // Excelを終了
                oXls.Quit();

                // COM オブジェクトの参照カウントを解放する 
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                oXls = null;
                oXlsBook = null;
                oxlsSheet = null;

                GC.Collect();
            }

            return rtnArray;
        }

        ///--------------------------------------------------------------------------
        /// <summary>
        ///     部署別勤務体系配列より勤務体系（シフト）コード名を取得する </summary>
        /// <param name="sName">
        ///     取得する勤務体系（シフト）コード名</param>
        /// <param name="bCode">
        ///     部署コード</param>
        /// <param name="sCode">
        ///     勤務体系（シフト）コード</param>
        /// <param name="sHol">
        ///     休日区分</param>
        /// <returns>
        ///     true:該当あり、false:該当なし</returns>
        ///--------------------------------------------------------------------------
        public bool getBushoSft(out string sName, string bCode, string sCode, string sHol)
        {
            sName = string.Empty;
            bool rtn = false;

            if (setShiftCode(out sName, zArray, bCode, sCode, sHol))
            {
                rtn = true;
            }

            return rtn;
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     勤務体系（シフト）名を取得する </summary>
        /// <param name="sName">
        ///     勤務体系名</param>
        /// <param name="sArray">
        ///     勤務体系（シフト）配列</param>
        /// <param name="bCode">
        ///     部署コード</param>
        /// <param name="sCode">
        ///     勤務体系コード</param>
        /// <param name="hol">
        ///     休日区分</param>
        ///------------------------------------------------------------------
        private bool setShiftCode(out string sName, object[,] sArray, string bCode, string sCode, string sHol)
        {
            sName = string.Empty;
            bool rtn = false;

            for (int i = 1; i <= sArray.GetLength(0); i++)
            {
                // 休日の有無をチェック要因から撤廃：2018/05/31
                //if (sArray[i, 1].ToString() == bCode && sArray[i, 2].ToString() == sCode && sArray[i, 4].ToString() == sHol)
                if (sArray[i, 1].ToString() == bCode && sArray[i, 2].ToString() == sCode)
                {
                    // シフト（勤務体系）名
                    sName = sArray[i, 3].ToString();
                    rtn = true;

                    // ループから抜ける
                    break;
                }
            }

            return rtn;
        }


        ///----------------------------------------------------------------------
        /// <summary>
        ///     部署別残業理由シートよりデータを配列に取得する </summary>
        /// <returns>
        ///     取得した配列：object</returns>
        ///----------------------------------------------------------------------
        public object[,] getZanReason()
        {
            object[,] rtnArray = null;

            string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

            Excel.Application oXls = new Excel.Application();

            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(Properties.Settings.Default.xlsBmnZanReason, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];
            oxlsSheet.Select(Type.Missing);

            Excel.Range rng = null;

            try
            {
                rng = oxlsSheet.Range[oxlsSheet.Cells[2, 1], oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 3]];
                rtnArray = (object[,])rng.Value2;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "部署別残業理由シート取り込みエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                // Bookをクローズ
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                // Excelを終了
                oXls.Quit();

                // COM オブジェクトの参照カウントを解放する 
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                oXls = null;
                oXlsBook = null;
                oxlsSheet = null;

                GC.Collect();
            }

            return rtnArray;
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     部署別残業理由別計画シートよりデータを配列に取得する </summary>
        /// <returns>
        ///     取得した配列：object</returns>
        ///----------------------------------------------------------------------
        public object[,] getZanReasonPlan()
        {
            object[,] rtnArray = null;

            string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

            Excel.Application oXls = new Excel.Application();

            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(Properties.Settings.Default.xlsBmnReZanPlan, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];
            oxlsSheet.Select(Type.Missing);

            Excel.Range rng = null;

            try
            {
                rng = oxlsSheet.Range[oxlsSheet.Cells[2, 1], oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 6]];
                rtnArray = (object[,])rng.Value2;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "部署別理由別残業計画シート取り込みエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                // Bookをクローズ
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                // Excelを終了
                oXls.Quit();

                // COM オブジェクトの参照カウントを解放する 
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                oXls = null;
                oXlsBook = null;
                oxlsSheet = null;

                GC.Collect();
            }

            return rtnArray;
        }

        ///--------------------------------------------------------------------------
        /// <summary>
        ///     部署別残業理由配列より名称を取得する </summary>
        /// <param name="sName">
        ///     取得する部署別残業理由名称</param>
        /// <param name="bCode">
        ///     部署コード</param>
        /// <param name="sCode">
        ///     理由コード</param>
        /// <returns>
        ///     true:該当あり、false:該当なし</returns>
        ///--------------------------------------------------------------------------
        public bool getBushoZanRe(out string sName, string bCode, string sCode)
        {
            sName = string.Empty;
            bool rtn = false;

            if (setReCode(out sName, rArray, bCode, sCode))
            {
                rtn = true;
            }

            return rtn;
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     部署別残業理由名称を取得する </summary>
        /// <param name="sName">
        ///     部署別残業理由名称</param>
        /// <param name="sArray">
        ///     部署別残業理由配列</param>
        /// <param name="bCode">
        ///     部署コード</param>
        /// <param name="sCode">
        ///     理由コード</param>
        ///------------------------------------------------------------------
        private bool setReCode(out string sName, object[,] sArray, string bCode, string sCode)
        {
            sName = string.Empty;
            bool rtn = false;

            for (int i = 1; i <= sArray.GetLength(0); i++)
            {
                if (sArray[i, 1].ToString() == bCode && sArray[i, 2].ToString() == Utility.StrtoInt(sCode).ToString())
                {
                    // 部署別残業理由名称
                    sName = sArray[i, 3].ToString();
                    rtn = true;

                    // ループから抜ける
                    break;
                }
            }

            return rtn;
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     部署別残業計画シートよりデータを配列に取得する </summary>
        /// <returns>
        ///     取得した配列：object</returns>
        ///----------------------------------------------------------------------
        public object[,] getZanPlan()
        {
            object[,] rtnArray = null;

            string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

            Excel.Application oXls = new Excel.Application();

            Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(Properties.Settings.Default.xlsBmnZanPlan, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing));

            Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[1];
            oxlsSheet.Select(Type.Missing);

            Excel.Range rng = null;

            try
            {
                rng = oxlsSheet.Range[oxlsSheet.Cells[2, 1], oxlsSheet.Cells[oxlsSheet.UsedRange.Rows.Count, 8]];
                rtnArray = (object[,])rng.Value2;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "部署別残業計画シート取り込みエラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                // Bookをクローズ
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                // Excelを終了
                oXls.Quit();

                // COM オブジェクトの参照カウントを解放する 
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                oXls = null;
                oXlsBook = null;
                oxlsSheet = null;

                GC.Collect();
            }

            return rtnArray;
        }
    }
}
