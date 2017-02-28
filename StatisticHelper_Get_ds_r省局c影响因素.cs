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
        public DataSet Get_ds_r省局c影响因素(bool showtoeapp)
        {
            DataSet ds = new DataSet("ds_r省局c影响因素");
            DataView dt_sourceview = new DataView(dt_source);
            excel.Application eapp = null;
            excel.Workbook book_dt_r省局c影响因素 = null;
            if (showtoeapp)
            {
                eapp = new excel.Application();
                book_dt_r省局c影响因素 = eapp.Workbooks.Add();
            }
            foreach (string ab in GDef.abtypelist)
            {
                Console.WriteLine("doing dt_r省局c影响因素 " + ab + " ...");
                DataTable dt_r省局c影响因素 = new DataTable(ab);
                string abfilter = "ab_type_name = '" + ab + "'";
                dt_sourceview.RowFilter = abfilter;
                DataTable type2list = dt_sourceview.ToTable(true, "type2_name");
                dt_r省局c影响因素.Columns.Add("省局");
                dt_r省局c影响因素.PrimaryKey = new DataColumn[] { dt_r省局c影响因素.Columns["省局"] };
                dt_r省局c影响因素.Columns.Add("总计", typeof(int));
                foreach (string u in GDef.unitnamelist)
                {
                    dt_r省局c影响因素.Rows.Add(u, 0);
                }
                DataTable tmps = dt_sourceview.ToTable();
                foreach (DataRow rtype2 in type2list.Rows)
                {
                    string type2 = rtype2[0].ToString();
                    DataColumn dc = new DataColumn(type2, typeof(int));
                    dc.DefaultValue = 0;
                    dt_r省局c影响因素.Columns.Add(dc);

                    var q = from logid in tmps.AsEnumerable()
                            where logid.Field<string>("type2_name") == type2
                            group logid by logid.Field<string>("unitname") into g
                            select new { u = g.Key, cnt = g.Count() };
                    foreach (var r in q)
                    {
                        dt_r省局c影响因素.Rows.Find(r.u)[type2] = r.cnt;
                        dt_r省局c影响因素.Rows.Find(r.u)["总计"] = dt_r省局c影响因素.Rows.Find(r.u).Field<int>("总计") + r.cnt;
                    }

                }
                if (showtoeapp)
                    dth.DTToExcelSheet(dt_r省局c影响因素, book_dt_r省局c影响因素, ab);
                ds.Tables.Add(dt_r省局c影响因素);
            }
            if (showtoeapp)
                eapp.Visible = true;
            return ds;
        }

    }
}