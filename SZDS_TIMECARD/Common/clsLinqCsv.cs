using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using LINQtoCSV;

namespace SZDS_TIMECARD.Common
{
    public class clsLinqCsv
    {
        [CsvColumn(FieldIndex = 1)]
        public string buCode { get; set; }
        [CsvColumn(FieldIndex = 2)]
        //public int saNum { get; set; } // 2018/02/05
        public string saNum { get; set; }
        [CsvColumn(FieldIndex = 3)]
        public DateTime workDate { get; set; }
        [CsvColumn(FieldIndex = 4)]
        public decimal zHayade { get; set; }
        [CsvColumn(FieldIndex = 5)]
        public decimal zFutsuu { get; set; }
        [CsvColumn(FieldIndex = 6)]
        public decimal zShinya { get; set; }
        [CsvColumn(FieldIndex = 7)]
        public decimal zKyuushutsu { get; set; }
        [CsvColumn(FieldIndex = 8)]
        public decimal zKyuShinya { get; set; }
        [CsvColumn(FieldIndex = 9)]
        public string zMark { get; set; }  // 記号 2018/02/10
    }
    
    public class clsLinqZan
    {
        [CsvColumn(FieldIndex = 1)]
        public string buCode { get; set; }
        [CsvColumn(FieldIndex = 2)]
        public int zYear { get; set; }
        [CsvColumn(FieldIndex = 3)]
        public int zMonth { get; set; }
        [CsvColumn(FieldIndex = 4)]
        public int zNin { get; set; }
        [CsvColumn(FieldIndex = 5)]
        public double zZanPlan { get; set; }
        [CsvColumn(FieldIndex = 6)]
        public int zSeisan { get; set; }
    }
    
    public class clsLinqkaByreByZan
    {
        [CsvColumn(FieldIndex = 1)]
        public string buCode { get; set; }
        [CsvColumn(FieldIndex = 2)]
        public string zYear { get; set; }
        [CsvColumn(FieldIndex = 3)]
        public string zMonth { get; set; }
        [CsvColumn(FieldIndex = 4)]
        public string zReason { get; set; }
        [CsvColumn(FieldIndex = 5)]
        public string zReName { get; set; }
        [CsvColumn(FieldIndex = 6)]
        public double zZanPlan { get; set; }
    }
}
