using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using Audit;
using excel = Microsoft.Office.Interop.Excel;

namespace StatisticHelper
{
    public partial class StatisticHelper
    {
        public DataTable Get_dt_r学科事件类型套数c月份(bool writetosheet, excel.Workbook book)
        {
            DataView dt_sourceview = new DataView(dt_source);
            DataTable dt_r学科事件类型套数c月份 = new DataTable("dt_r学科事件类型套数c月份");
            dt_r学科事件类型套数c月份.Columns.Add("学科");
            dt_r学科事件类型套数c月份.Columns.Add("事件类型");
            dt_r学科事件类型套数c月份.PrimaryKey = new DataColumn[] { dt_r学科事件类型套数c月份.Columns["学科"], dt_r学科事件类型套数c月份.Columns["事件类型"] };
            foreach (string sci in GDef.sciencelist)
            {
                foreach (string abname in GDef.abtypelist)
                {
                    dt_r学科事件类型套数c月份.Rows.Add(sci, abname + "(套)");
                }
                dt_r学科事件类型套数c月份.Rows.Add(sci, sci + "总计(套)");
            }
            for (int i = 1; i <= 12; i++)
            {
                Console.WriteLine(i);
                dt_sourceview.RowFilter = string.Format("END_DATE >= '{0}' and START_DATE <= '{1}'", new DateTime(DATEBEGIN.Year, i, 1), new DateTime(DATEBEGIN.Year, i, 1).AddMonths(1).AddSeconds(-1));
                DataColumn tmpdc = new DataColumn(i + "月", typeof(int));
                tmpdc.DefaultValue = 0;
                dt_r学科事件类型套数c月份.Columns.Add(tmpdc);
                DataTable tmp = dt_sourceview.ToTable(true, "sp", "science", "ab_type_name");
                var q = from sp in tmp.AsEnumerable()
                        group sp by new { _sci = sp["science"], _ab = sp["ab_type_name"] } into g
                        select new { _sci = g.Key._sci, _ab = g.Key._ab, cnt = g.Count() };
                foreach (var r in q)
                {
                    dt_r学科事件类型套数c月份.Rows.Find(new object[] { r._sci, r._ab + "(套)" })[i + "月"] = r.cnt;
                }
                foreach (string sci in GDef.sciencelist)
                {
                    int sum = 0;
                    foreach (string ab in GDef.abtypelist)
                    {
                        sum += dt_r学科事件类型套数c月份.Rows.Find(new object[] { sci, ab + "(套)" }).Field<int>(i + "月");
                    }
                    dt_r学科事件类型套数c月份.Rows.Find(new object[] { sci, sci + "总计(套)" })[i + "月"] = sum;
                }

            }
            if (writetosheet)
                dth.DTToExcelSheet(dt_r学科事件类型套数c月份, book, "dt_r学科事件类型套数c月份");
            return dt_r学科事件类型套数c月份;
        }

    }
}