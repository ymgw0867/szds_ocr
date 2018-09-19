using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SZDS_TIMECARD.sumData
{
    class clsZanSum
    {
        public string sSzCode { get; set; }         // 部署コード
        public int sDay { get; set; }               // 日付
        public double sZangyo { get; set; }         // 該当日残業時間
        public double sMonthPlan { get; set; }      // 月間計画値
        public double sPlanbyDay { get; set; }      // 日別計画値
        public string sZissekibyDay { get; set; }   // 実績線データ
        public int sYear { get; set; }              // 年
        public int sMonth { get; set; }             // 月
        public int sEndDay { get; set; }            // 月末日
        public int sHoliday { get; set; }           // 休日
    }
}
