using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Data;
using System.Data.SqlClient;
using SZDS_TIMECARD.Common;

namespace SZDS_TIMECARD.Common
{
    class dbControl
    {
        /// <summary>
        /// DataControlクラスの基本クラス
        /// </summary>
        public class BaseControl
        {
            private DBConnect DBConnect;
            protected OleDbConnection dbControlCn;

            // BaseControlのコンストラクタ。DBConnectクラスのインスタンスを作成します。
            public BaseControl(string dbName)
            {
                // データベースをオープンする
                DBConnect = new DBConnect(dbName);
            }

            // データベースに接続しコネクション情報を返す
            public OleDbConnection GetConnection()
            {
                dbControlCn = DBConnect.Cn;
                return DBConnect.Cn;
            }
        }

        public class DataControl : BaseControl
        {
            // データコントロールクラスのコンストラクタ
            public DataControl(string dbName):base(dbName)
            {
            }

            /// <summary>
            /// データベース接続解除
            /// </summary>
            public void Close()
            {
                if (dbControlCn.State == ConnectionState.Open)
                {
                    dbControlCn.Close();
                }
            }

            /// <summary>
            /// 任意のSQLを実行する
            /// </summary>
            /// <param name="tempSql">SQL文</param>
            /// <returns>成功 : true, 失敗 : false</returns>
            public bool FreeSql(string tempSql)
            {
                bool rValue = false;

                try
                {
                    OleDbCommand sCom = new OleDbCommand();
                    sCom.CommandText = tempSql;
                    sCom.Connection = GetConnection();

                    //SQLの実行
                    sCom.ExecuteNonQuery();
                    rValue = true;
                }
                catch (Exception ex)
                {
                    rValue = false;
                }

                return rValue;
            }

            /// <summary>
            /// データリーダーを取得する
            /// </summary>
            /// <param name="tempSQL">SQL文</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader FreeReader(string tempSQL)
            {
                OleDbCommand sCom = new OleDbCommand();
                sCom.CommandText = tempSQL;
                sCom.Connection = GetConnection();
                OleDbDataReader dR = sCom.ExecuteReader();

                return dR;
            }

            /// <summary>
            /// 社員情報を取得します
            /// </summary>
            /// <param name="sYY">基準年</param>
            /// <param name="sMM">基準月</param>
            /// <returns>データリーダー</returns>
            public OleDbDataReader GetEmployeeBase(string sYY, string sMM, string sDD, string sNo)
            {
                string tempDate;

                //基準年月日
                string sDate = sYY.ToString() + "/" + sMM + "/" + sDD;
                DateTime eDate;
                if (DateTime.TryParse(sDate, out eDate)) tempDate = eDate.ToShortDateString();   //日付を返す
                else tempDate = DateTime.Today.ToShortDateString();　　//当日日付を返す

                //// SQLServer接続
                ////dbControl.DataControl dCon = new dbControl.DataControl(_PCADBName);
                OleDbDataReader dRs;
                StringBuilder sb = new StringBuilder();
                string SqlStr = string.Empty;

                sb.Append("select tbDepartment.DepartmentID,tbDepartment.DepartmentCode,tbDepartment.DepartmentName,tbEmployeeBase.EmployeeNo,tbEmployeeBase.Name,tbHR_DivisionCategory.CategoryName ");
                sb.Append("from ((tbEmployeeBase inner join tbHR_DivisionCategory ");
                sb.Append("on EmploymentDivisionID = CategoryID) left join ");

                sb.Append("(select tbEmployeeMainDutyPersonnelChange.EmployeeID,tbEmployeeMainDutyPersonnelChange.BelongID from tbEmployeeMainDutyPersonnelChange inner join (");
                sb.Append("select EmployeeID,max(AnnounceDate) as AnnounceDate from tbEmployeeMainDutyPersonnelChange ");
                sb.Append("where AnnounceDate <= '" + tempDate + "'");
                sb.Append("group by EmployeeID) as a ");
                sb.Append("on (tbEmployeeMainDutyPersonnelChange.EmployeeID = a.EmployeeID) and ");
                sb.Append("(tbEmployeeMainDutyPersonnelChange.AnnounceDate = a.AnnounceDate) ");
                sb.Append("inner join tbDepartment ");
                sb.Append("on tbEmployeeMainDutyPersonnelChange.BelongID = tbDepartment.DepartmentID ");
                sb.Append(") as d ");

                sb.Append("on tbEmployeeBase.EmployeeID = d.EmployeeID) left join ");
                sb.Append("tbDepartment on d.BelongID = tbDepartment.DepartmentID ");

                sb.Append("where tbEmployeeBase.EmployeeNo = '" + string.Format("{0:0000000000}", int.Parse(sNo)) + "' ");
                sb.Append(" and BeOnTheRegisterDivisionID != 9");

                dRs = FreeReader(sb.ToString());
                return dRs;
            }
        }

        /// <summary>
        /// SQLServerデータベース接続クラス
        /// </summary>
        public class DBConnect
        {
            OleDbConnection cn = new OleDbConnection();

            public OleDbConnection Cn
            {
                get
                {
                    return cn;
                }
            }

            private string sServerName;
            private string sLogin;
            private string sPass;
            private string sDatabase;

            public DBConnect(string dbName)
            {
                try
                {
                    // MySeting項目の取得
                    sServerName = Properties.Settings.Default.SQLServerName;    // サーバ名
                    sLogin = Properties.Settings.Default.SQLLogin;              // ログイン名
                    sPass = Properties.Settings.Default.SQLPass;                // パスワード
                    sDatabase = dbName;                                         // データベース名

                    // データベース接続文字列
                    cn.ConnectionString = "";
                    cn.ConnectionString += "Provider=SQLOLEDB;";
                    cn.ConnectionString += "SERVER=" + sServerName + ";";
                    cn.ConnectionString += "DataBase=" + sDatabase + ";";
                    cn.ConnectionString += "UID=" + sLogin + ";";
                    cn.ConnectionString += "PWD=" + sPass + ";";
                    cn.ConnectionString += "WSID=";

                    cn.Open();
                }

                catch (Exception e)
                {
                    throw e;
                }
            }
        }        
    }
}
