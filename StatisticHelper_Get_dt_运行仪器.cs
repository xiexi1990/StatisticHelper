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
        public DataTable Get_dt_运行仪器(bool writetosheet, excel.Workbook book)
        {
            DataTable dt_运行仪器 = new DataTable("dt_运行仪器");
            dt_运行仪器.Columns.Add("学科");
            dt_运行仪器.Columns.Add("在运行仪器");
            dt_运行仪器.Columns.Add("分析完整率超95%仪器");
            dt_运行仪器.Columns.Add("应分析仪器");
            dt_运行仪器.Columns.Add("分析完整率超95%的应分析仪器");
            dt_运行仪器.Columns.Add("分析完整率不足50%仪器");
            dt_运行仪器.Columns.Add("分析完整率超95%的增加分析仪器");
            foreach (string sci in GDef.sciencelist)
            {
                DataRow r = dt_运行仪器.NewRow();
                r["学科"] = sci;
                r["在运行仪器"] = dt_eval_instrid.Compute("count(sp)", "science = '" + sci + "'");
                r["分析完整率超95%仪器"] = dt_eval_instrid_comp.Compute("count(sp)", "science = '" + sci + "'");
                r["应分析仪器"] = dt_eval_y_instrid.Compute("count(sp)", "science = '" + sci + "'");
                r["分析完整率超95%的应分析仪器"] = dt_eval_y_instrid_comp.Compute("count(sp)", "science = '" + sci + "'");
                {
                    string str = "";
                    var q = from sp in dt_eval_y_instrid.AsEnumerable()
                            join sta in dt_stationlist.AsEnumerable() on sp.Field<string>("stationid") equals sta.Field<string>("stationid")
                            where sp.Field<double>("comp") < 0.5 && sp.Field<string>("science") == sci
                            select new
                            {
                                uname = GDef.unitcode2abbrunitname(sp.Field<string>("unitcode")),
                                staname = sta.Field<string>("stationname"),
                                pnt = sp.Field<string>("pointid"),
                                instrname = sp.Field<string>("fullinstrname"),
                                comp = Math.Round(sp.Field<double>("comp") * 100, 1)
                            };
                    foreach (var ins in q)
                    {
                        str += ins.uname + ins.staname + "[" + ins.pnt + "]" + ins.instrname + "(" + ins.comp + "%)、";
                    }
                    r["分析完整率不足50%仪器"] = str == "" ? "" : str.Substring(0, str.Length - 1);
                }
                r["分析完整率超95%的增加分析仪器"] = dt_eval_not_y_instrid_comp.Compute("count(sp)", "science = '" + sci + "'");
                dt_运行仪器.Rows.Add(r);
            }
            if (writetosheet)
                dth.DTToExcelSheet(dt_运行仪器, book, "dt_运行仪器");
            return dt_运行仪器;
        }

    }
}