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
        public DataSet Get_ds_r影响因素c学科(bool showtoeapp)
        {
            DataSet ds = new DataSet("ds_r影响因素c学科");
            DataView dt_sourceview = new DataView(dt_source);
            excel.Application eapp = null;
            excel.Workbook book_dt_r影响因素c学科 = null;
            if (showtoeapp)
            {
                eapp = new excel.Application();
                book_dt_r影响因素c学科 = eapp.Workbooks.Add();
            }
            foreach (string ab in GDef.abtypelist)
            {
                Console.WriteLine("doing dt_r影响因素c学科 " + ab + " ...");
                DataTable dt_r影响因素c学科 = new DataTable(ab);
                string abfilter = "ab_type_name = '" + ab + "'";
                dt_sourceview.RowFilter = abfilter;
                DataTable type2list = dt_sourceview.ToTable(true, "type2_name");
                dt_r影响因素c学科.Columns.Add("影响因素");
                dt_r影响因素c学科.PrimaryKey = new DataColumn[] { dt_r影响因素c学科.Columns["影响因素"] };
                dt_r影响因素c学科.Columns.Add("总计", typeof(int));
                foreach (DataRow r in type2list.Rows)
                {
                    dt_r影响因素c学科.Rows.Add(r["type2_name"], 0);
                }
                DataTable tmps = dt_sourceview.ToTable();
                foreach (string sci in GDef.sciencelist)
                {
                    DataColumn dc = new DataColumn(sci, typeof(int));
                    dc.DefaultValue = 0;
                    dt_r影响因素c学科.Columns.Add(dc);

                    var q = from logid in tmps.AsEnumerable()
                            where logid.Field<string>("science") == sci
                            group logid by logid.Field<string>("type2_name") into g
                            select new { t2 = g.Key, cnt = g.Count() };
                    foreach (var r in q)
                    {
                        dt_r影响因素c学科.Rows.Find(r.t2)[sci] = r.cnt;
                        dt_r影响因素c学科.Rows.Find(r.t2)["总计"] = dt_r影响因素c学科.Rows.Find(r.t2).Field<int>("总计") + r.cnt;
                    }

                }
                if (showtoeapp)
                    dth.DTToExcelSheet(dt_r影响因素c学科, book_dt_r影响因素c学科, ab);
                ds.Tables.Add(dt_r影响因素c学科);
            }
            if (showtoeapp)
                eapp.Visible = true;
            return ds;
        }

    }
}