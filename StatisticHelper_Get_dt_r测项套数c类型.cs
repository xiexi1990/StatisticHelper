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
        public DataTable Get_dt_r测项套数c类型(bool writetosheet, excel.Workbook book)
        {
            DataView dt_sourceview = new DataView(dt_source);
            DataTable items = (new DataView(dt_sp)).ToTable(true, "science", "item");
            DataTable dt_r测项套数c类型 = new DataTable("dt_r测项套数c类型");
            dt_r测项套数c类型.Columns.Add("学科");
            dt_r测项套数c类型.Columns.Add("测项");
            dt_r测项套数c类型.PrimaryKey = new DataColumn[] { dt_r测项套数c类型.Columns["测项"] };
            dt_r测项套数c类型.Columns.Add("在运行仪器（套）", typeof(int));
            dt_r测项套数c类型.Columns["在运行仪器（套）"].DefaultValue = 0;
            foreach (string sci in GDef.sciencelist)
            {
                foreach (DataRow r in items.Rows)
                {
                    if (r.Field<string>("science") == sci)
                    {
                        dt_r测项套数c类型.Rows.Add(r["science"], r["item"]);
                    }
                }
            }
            {
                var q = from sp in dt_sp.AsEnumerable()
                        group sp by sp.Field<string>("item") into g
                        select new { item = g.Key, cnt = g.Count() };
                foreach (var r in q)
                {
                    dt_r测项套数c类型.Rows.Find(r.item)["在运行仪器（套）"] = r.cnt;
                }

            }

            foreach (string abname in GDef.abtypelist)
            {
                DataColumn dc = new DataColumn(abname + "受影响仪器套数", typeof(int));
                dc.DefaultValue = 0;
                dt_r测项套数c类型.Columns.Add(dc);
                dc = new DataColumn(abname + "受影响比例(%)");
                dc.DefaultValue = "0";
                dt_r测项套数c类型.Columns.Add(dc);

                dt_sourceview.RowFilter = "ab_type_name = '" + abname + "'";
                var q = from sp in dt_sourceview.ToTable(true, "sp", "item").AsEnumerable()
                        group sp by sp.Field<string>("item") into g
                        select new { item = g.Key, cnt = g.Count() };

                foreach (var r in q)
                {
                    dt_r测项套数c类型.Rows.Find(r.item)[abname + "受影响仪器套数"] = r.cnt;
                    if (dt_r测项套数c类型.Rows.Find(r.item).Field<int>("在运行仪器（套）") != 0)
                    {
                        dt_r测项套数c类型.Rows.Find(r.item)[abname + "受影响比例(%)"] = Math.Round(Convert.ToDouble(r.cnt) / dt_r测项套数c类型.Rows.Find(r.item).Field<int>("在运行仪器（套）") * 100.0, 1);
                    }
                    else
                    {
                        dt_r测项套数c类型.Rows.Find(r.item)[abname + "受影响比例(%)"] = "-";
                    }
                }
            }

            if (writetosheet)
                dth.DTToExcelSheet(dt_r测项套数c类型, book, "dt_r测项套数c类型");
            return dt_r测项套数c类型;
        }

    }
}