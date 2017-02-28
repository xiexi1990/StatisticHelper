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
        public void Get_dss_r测项突出测点c累计时长等(bool showtoeapp = true)
        {
            excel.Application eapp = new excel.Application();
            DataView dt_sourceview = new DataView(dt_source);
            DataView dt_spview = new DataView(dt_sp);
            DataTable items = dt_spview.ToTable(true, "science", "item");

            foreach (string ab in GDef.abtypelist)
            {
                Console.WriteLine("doing dt_r测项突出测点c累计时长等 " + ab + " ...");
                excel.Workbook abbook = eapp.Workbooks.Add();

                string abfilter = "ab_type_name = '" + ab + "'";
                dt_sourceview.RowFilter = abfilter;
                DataTable type2list = dt_sourceview.ToTable(true, "type2_name");

                foreach (DataRow dr in type2list.Rows)
                {
                    string type2 = dr[0].ToString();
                    Console.Write("\rdoing dt_r测项突出测点c累计时长等 " + ab + " " + type2 + " ...");

                    dt_sourceview.RowFilter = abfilter + "and type2_name = '" + type2 + "'";
                    DataTable dt_ab_lid = dt_sourceview.ToTable();
                    //          DataTable dt_ab_sp_cnt = dt_sourceview.ToTable(true, "item", "science", "sp");

                    DataTable dt_ab_type2_out = new DataTable();
                    dt_ab_type2_out.Columns.Add("学科");
                    dt_ab_type2_out.Columns.Add("测项");
                    dt_ab_type2_out.Columns.Add(type2 + "_总事件数", typeof(int));
                    dt_ab_type2_out.Columns.Add("台站");
                    dt_ab_type2_out.Columns.Add("突出测点");
                    dt_ab_type2_out.Columns.Add("事件次数", typeof(int));
                    dt_ab_type2_out.Columns.Add("累计时长(天)", typeof(double));
                    dt_ab_type2_out.Columns.Add("平均时长(天)", typeof(double));

                    foreach (string sci in GDef.sciencelist)
                    {
                        bool scifirst = true;
                        foreach (DataRow r in items.Rows)
                        {
                            if (r.Field<string>("science") == sci)
                            {
                                var q = from lid in dt_ab_lid.AsEnumerable()
                                        where lid.Field<string>("science") == sci && lid.Field<string>("item") == r.Field<string>("item")
                                        group lid by new { item = lid["item"], station = lid["stationname"], sp = lid["sp"], instrname = lid["instr"] } into g
                                        orderby g.Sum(s => s.Field<decimal>("len")) descending
                                        select new { item = g.Key.item, station = g.Key.station, sp = g.Key.sp, intrname = g.Key.instrname, tlen = Math.Round(g.Sum(s => s.Field<decimal>("len")), 1), cnt = g.Count(), alen = Math.Round(g.Sum(s => s.Field<decimal>("len")) / g.Count(), 1) };
                                bool itemfirst = true;
                                foreach (var qr in q.Take(5))
                                {
                                    DataRow tr = dt_ab_type2_out.NewRow();
                                    if (itemfirst)
                                    {
                                        itemfirst = false;
                                        tr["测项"] = r["item"];
                                        tr[type2 + "_总事件数"] = dt_ab_lid.Compute("count(log_id)", string.Format("science = '{0}' and item = '{1}'", sci, r["item"]));
                                    }
                                    if (scifirst)
                                    {
                                        scifirst = false;
                                        tr["学科"] = sci;
                                    }
                                    tr["台站"] = qr.station;
                                    tr["突出测点"] = qr.intrname + "(" + qr.sp + ")";
                                    tr["事件次数"] = qr.cnt;
                                    tr["累计时长(天)"] = qr.tlen;
                                    tr["平均时长(天)"] = qr.alen;
                                    dt_ab_type2_out.Rows.Add(tr);
                                }
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

                abbook.SaveAs("c:\\qztemp\\" + ab + "_测项突出测点_累计时长等");

            }
            eapp.Visible = true;
        }

    }
}