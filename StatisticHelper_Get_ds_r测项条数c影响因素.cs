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
        public DataSet Get_ds_r测项条数c影响因素(bool showtoeapp)
        {
            DataSet ds = new DataSet("ds_r测项条数c影响因素");

            DataView dt_sourceview = new DataView(dt_source);
            DataTable items = (new DataView(dt_sp)).ToTable(true, "science", "item");
            excel.Application eapp = null;
            excel.Workbook book_dt_r测项条数c影响因素 = null;
            if (showtoeapp)
            {
                eapp = new excel.Application();
                book_dt_r测项条数c影响因素 = eapp.Workbooks.Add();
            }
            foreach (string ab in GDef.abtypelist)
            {
                Console.WriteLine("doing dt_r测项条数c影响因素 " + ab + " ...");
                DataTable dt_r测项条数c影响因素 = new DataTable(ab);
                dt_r测项条数c影响因素.Columns.Add("学科");
                dt_r测项条数c影响因素.Columns.Add("测项");
                dt_r测项条数c影响因素.PrimaryKey = new DataColumn[] { dt_r测项条数c影响因素.Columns["测项"] };
                foreach (string sci in GDef.sciencelist)
                {
                    foreach (DataRow r in items.Rows)
                    {
                        if (r.Field<string>("science") == sci)
                        {
                            dt_r测项条数c影响因素.Rows.Add(r["science"], r["item"]);
                        }
                    }
                }
                string abfilter = "ab_type_name = '" + ab + "'";
                dt_sourceview.RowFilter = abfilter;
                DataTable type2list = dt_sourceview.ToTable(true, "type2_name");

                {
                    DataColumn dc = new DataColumn("总计(条)", typeof(int));
                    dc.DefaultValue = 0;
                    dt_r测项条数c影响因素.Columns.Add(dc);
                }

                DataTable tmps = dt_sourceview.ToTable();
                foreach (DataRow rtype2 in type2list.Rows)
                {
                    string type2 = rtype2[0].ToString();
                    DataColumn dc = new DataColumn(type2 + "(条)", typeof(int));
                    dc.DefaultValue = 0;
                    dt_r测项条数c影响因素.Columns.Add(dc);

                    var q = from lid in tmps.AsEnumerable()
                            where lid.Field<string>("type2_name") == type2 && lid.Field<string>("ab_type_name") == ab
                            group lid by lid.Field<string>("item") into g
                            select new { item = g.Key, cnt = g.Count() };
                    foreach (var r in q)
                    {
                        dt_r测项条数c影响因素.Rows.Find(r.item)[type2 + "(条)"] = r.cnt;
                        dt_r测项条数c影响因素.Rows.Find(r.item)["总计(条)"] = dt_r测项条数c影响因素.Rows.Find(r.item).Field<int>("总计(条)") + r.cnt;
                    }
                }
                dt_r测项条数c影响因素.Columns["总计(条)"].SetOrdinal(dt_r测项条数c影响因素.Columns.Count - 1);
                if (showtoeapp)
                    dth.DTToExcelSheet(dt_r测项条数c影响因素, book_dt_r测项条数c影响因素, ab);
                ds.Tables.Add(dt_r测项条数c影响因素);
            }
            if (showtoeapp)
                eapp.Visible = true;
            return ds;
        }

    }
}