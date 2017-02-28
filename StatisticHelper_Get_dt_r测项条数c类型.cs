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
        public DataTable Get_dt_r测项条数c类型(bool writetosheet, excel.Workbook book)
        {
            DataTable dt_r测项条数c类型 = new DataTable("dt_r测项条数c类型");
            dt_r测项条数c类型.Columns.Add("学科");
            dt_r测项条数c类型.Columns.Add("测项");
            dt_r测项条数c类型.PrimaryKey = new DataColumn[] { dt_r测项条数c类型.Columns["测项"] };
            DataTable items = (new DataView(dt_sp)).ToTable(true, "science", "item");
            foreach (string sci in GDef.sciencelist)
            {
                foreach (DataRow r in items.Rows)
                {
                    if (r.Field<string>("science") == sci)
                    {
                        dt_r测项条数c类型.Rows.Add(r["science"], r["item"]);
                    }
                }
            }

            foreach (string abname in GDef.abtypelist)
            {
                DataColumn dc = new DataColumn(abname + "(条)", typeof(int));
                dc.DefaultValue = 0;
                dt_r测项条数c类型.Columns.Add(dc);
                var q = from lid in dt_source.AsEnumerable()
                        where lid.Field<string>("ab_type_name") == abname
                        group lid by lid.Field<string>("item") into g
                        select new { item = g.Key, cnt = g.Count() };

                foreach (var r in q)
                {
                    dt_r测项条数c类型.Rows.Find(r.item)[abname + "(条)"] = r.cnt;
                }
            }
            if (writetosheet)
                dth.DTToExcelSheet(dt_r测项条数c类型, book, "dt_r测项条数c类型");
            return dt_r测项条数c类型;
        }

    }
}