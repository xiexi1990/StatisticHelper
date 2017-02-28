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
        public DataSet Get_ds_r测项套数c影响因素(bool showtoeapp)
        {
            DataSet ds = new DataSet("ds_r测项套数c影响因素");
            DataView dt_sourceview = new DataView(dt_source);
            DataTable items = (new DataView(dt_sp)).ToTable(true, "science", "item");
            excel.Application eapp = null;
            excel.Workbook book_dt_r测项套数c影响因素 = null;
            if (showtoeapp)
            {
                eapp = new excel.Application();
                book_dt_r测项套数c影响因素 = eapp.Workbooks.Add();
            }
            foreach (string ab in GDef.abtypelist)
            {
                Console.WriteLine("doing dt_r测项套数c影响因素 " + ab + " ...");
                DataTable dt_r测项套数c影响因素 = new DataTable(ab);
                dt_r测项套数c影响因素.Columns.Add("学科");
                dt_r测项套数c影响因素.Columns.Add("测项");
                dt_r测项套数c影响因素.PrimaryKey = new DataColumn[] { dt_r测项套数c影响因素.Columns["测项"] };

                foreach (string sci in GDef.sciencelist)
                {
                    foreach (DataRow r in items.Rows)
                    {
                        if (r.Field<string>("science") == sci)
                        {
                            dt_r测项套数c影响因素.Rows.Add(r["science"], r["item"]);
                        }
                    }
                }
                string abfilter = "ab_type_name = '" + ab + "'";
                dt_sourceview.RowFilter = abfilter;
                DataTable type2list = dt_sourceview.ToTable(true, "type2_name");

                DataTable tmps = dt_sourceview.ToTable(true, "sp", "item", "type2_name");
                foreach (DataRow rtype2 in type2list.Rows)
                {
                    string type2 = rtype2[0].ToString();
                    DataColumn dc = new DataColumn(type2 + "(套)", typeof(int));
                    dc.DefaultValue = 0;
                    dt_r测项套数c影响因素.Columns.Add(dc);

                    var q = from sp in tmps.AsEnumerable()
                            where sp.Field<string>("type2_name") == type2
                            group sp by sp.Field<string>("item") into g
                            select new { item = g.Key, cnt = g.Count() };
                    foreach (var r in q)
                    {
                        dt_r测项套数c影响因素.Rows.Find(r.item)[type2 + "(套)"] = r.cnt;
                    }
                }
                {
                    DataColumn dc = new DataColumn("总计(套)", typeof(int));
                    dc.DefaultValue = 0;
                    dt_r测项套数c影响因素.Columns.Add(dc);
                    DataTable tmps2 = dt_sourceview.ToTable(true, "sp", "item");
                    var q = from sp in tmps2.AsEnumerable()
                            group sp by sp.Field<string>("item") into g
                            select new { item = g.Key, cnt = g.Count() };
                    foreach (var r in q)
                    {
                        dt_r测项套数c影响因素.Rows.Find(r.item)["总计(套)"] = r.cnt;
                    }
                }
                if (showtoeapp)
                    dth.DTToExcelSheet(dt_r测项套数c影响因素, book_dt_r测项套数c影响因素, ab);
                ds.Tables.Add(dt_r测项套数c影响因素);
            }
            if (showtoeapp)
                eapp.Visible = true;
            return ds;
        }

    }
}