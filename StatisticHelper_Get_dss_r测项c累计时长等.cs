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
        public void Get_dss_r测项c累计时长等(bool showtoeapp = true)
        {
            excel.Application eapp = new excel.Application();
            DataView dt_sourceview = new DataView(dt_source);
            DataView dt_spview = new DataView(dt_sp);
            DataTable items = dt_spview.ToTable(true, "science", "item");

            foreach (string ab in GDef.abtypelist)
            {
                Console.WriteLine("doing dt_r测项c累计时长等 " + ab + " ...");
                excel.Workbook abbook = eapp.Workbooks.Add();

                string abfilter = "ab_type_name = '" + ab + "'";
                dt_sourceview.RowFilter = abfilter;
                DataTable type2list = dt_sourceview.ToTable(true, "type2_name");

                foreach (DataRow dr in type2list.Rows)
                {
                    string type2 = dr[0].ToString();
                    Console.Write("\rdoing dt_r测项c累计时长等 " + ab + " " + type2 + " ...");

                    dt_sourceview.RowFilter = abfilter + "and type2_name = '" + type2 + "'";
                    DataTable dt_ab_lid = dt_sourceview.ToTable();
                    DataTable dt_ab_sp_cnt = dt_sourceview.ToTable(true, "item", "science", "sp");

                    DataTable dt_ab_type2_out = new DataTable();
                    dt_ab_type2_out.Columns.Add("学科");
                    dt_ab_type2_out.Columns.Add(type2 + "_总事件数", typeof(int));

                    DataColumn tmpdc = new DataColumn("测项");
                    dt_ab_type2_out.Columns.Add(tmpdc);
                    dt_ab_type2_out.PrimaryKey = new DataColumn[] { tmpdc };

                    tmpdc = new DataColumn("事件数", typeof(int));
                    tmpdc.DefaultValue = 0;
                    dt_ab_type2_out.Columns.Add(tmpdc);
                    tmpdc = new DataColumn("累计时长(天)", typeof(double));
                    tmpdc.DefaultValue = 0;
                    dt_ab_type2_out.Columns.Add(tmpdc);
                    tmpdc = new DataColumn("平均时长(天)", typeof(double));
                    tmpdc.DefaultValue = 0;
                    dt_ab_type2_out.Columns.Add(tmpdc);
                    tmpdc = new DataColumn("受影响仪器套数", typeof(int));
                    tmpdc.DefaultValue = 0;
                    dt_ab_type2_out.Columns.Add(tmpdc);
                    tmpdc = new DataColumn("仪器总数", typeof(int));
                    tmpdc.DefaultValue = 0;
                    dt_ab_type2_out.Columns.Add(tmpdc);
                    tmpdc = new DataColumn("比例(%)");
                    tmpdc.DefaultValue = "0";
                    dt_ab_type2_out.Columns.Add(tmpdc);

                    foreach (string sci in GDef.sciencelist)
                    {
                        bool first = true;
                        foreach (DataRow r in items.Rows)
                        {
                            if (r.Field<string>("science") == sci)
                            {
                                DataRow tr = dt_ab_type2_out.NewRow();
                                tr["测项"] = r["item"];
                                if (first)
                                {
                                    first = false;
                                    tr["学科"] = sci;
                                    tr[type2 + "_总事件数"] = dt_ab_lid.Compute("count(log_id)", "science = '" + sci + "'");
                                }
                                dt_ab_type2_out.Rows.Add(tr);
                            }
                        }
                        {
                            var q = from lid in dt_ab_lid.AsEnumerable()
                                    where lid.Field<string>("science") == sci
                                    group lid by new { item = lid["item"] } into g
                                    select new { item = g.Key.item, tlen = Math.Round(g.Sum(s => s.Field<decimal>("len")), 1), cnt = g.Count(), alen = Math.Round(g.Sum(s => s.Field<decimal>("len")) / g.Count(), 1) };
                            foreach (var r in q)
                            {
                                dt_ab_type2_out.Rows.Find(r.item)["事件数"] = r.cnt;
                                dt_ab_type2_out.Rows.Find(r.item)["累计时长(天)"] = r.tlen;
                                dt_ab_type2_out.Rows.Find(r.item)["平均时长(天)"] = r.alen;
                            }
                        }
                        {
                            var q = from sp in dt_ab_sp_cnt.AsEnumerable()
                                    where sp.Field<string>("science") == sci
                                    group sp by new { item = sp["item"] } into g
                                    select new { item = g.Key.item, cnt = g.Count() };
                            foreach (var r in q)
                            {
                                dt_ab_type2_out.Rows.Find(r.item)["受影响仪器套数"] = r.cnt;
                            }
                        }
                        {
                            var q = from sp in dt_sp.AsEnumerable()
                                    where sp.Field<string>("science") == sci
                                    group sp by new { item = sp["item"] } into g
                                    select new { item = g.Key.item, cnt = g.Count() };
                            foreach (var r in q)
                            {
                                dt_ab_type2_out.Rows.Find(r.item)["仪器总数"] = r.cnt;
                                dt_ab_type2_out.Rows.Find(r.item)["比例(%)"] = Math.Round(dt_ab_type2_out.Rows.Find(r.item).Field<int>("受影响仪器套数") / (decimal)r.cnt * 100, 1);
                            }
                        }
                    }

                    var shname = (ab + "_" + type2).ToCharArray();
                    for (int i = 0; i < shname.Length; i++)
                    {
                        if (i == 20)
                        {
                            shname[20] = (char)0;
                            break;
                        }
                        if (":\\/?*[]".Contains(shname[i]))
                            shname[i] = '_';
                    }

                    dth.DTToExcelSheet(dt_ab_type2_out, abbook, new string(shname));

                }

                abbook.SaveAs("c:\\qztemp\\" + ab + "_测项_累计时长等");

            }
            eapp.Visible = true;
        }

    }
}