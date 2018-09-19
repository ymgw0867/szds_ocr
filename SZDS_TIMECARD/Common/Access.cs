using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace SZDS_TIMECARD.Common
{
    public class Access
    {
        public class DataAccess
        {
            public DataAccess()
            {
            }

            // 条件付きデータリーダー取得インターフェイス
            public interface IFillBy
            {
                // 抽象メソッド
                SqlDataReader GetdsReader(SqlConnection tempConnection, string tempString);
            }

            // データリーダー取得クラス
            public class free_dsReader : IFillBy
            {
                private SqlCommand SCom = new SqlCommand();
                private String mySql;
                private SqlDataReader dR;

                ///----------------------------------------------------------------------
                /// <summary>
                ///     データリーダー取得 </summary>
                /// <param name="tempConnection">
                ///     データベース接続情報</param>
                /// <param name="tempString">
                ///     SQL文</param>
                /// <returns>
                ///     データリーダー</returns>
                ///----------------------------------------------------------------------
                public SqlDataReader GetdsReader(SqlConnection tempConnection, string tempString)
                {
                    //throw new Exception("The method or operation is not implemented.");

                    mySql = "";
                    mySql += tempString;
                    SCom.CommandText = mySql;
                    SCom.Connection = tempConnection;
                    dR = SCom.ExecuteReader();
                    return dR;
                }
            }
        }
    }
}
