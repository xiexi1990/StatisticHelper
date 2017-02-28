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
        public DataTable Get_dt_总表(bool writetosheet, excel.Workbook book)
        {
            DataView dt_sourceview = new DataView(dt_source);
            DataView dt_spview = new DataView(dt_sp);
            DataTable dt_总表 = new DataTable("dt_总表");
            dt_总表.Columns.Add("学科");
            foreach (string s in GDef.sciencelist)
            {
                dt_总表.Columns.Add(s);
            }
            dt_总表.Columns.Add("总计");

            foreach (string abtype in GDef.abtypelist)
            {
                DataRow r = dt_总表.NewRow();
                r[0] = abtype + "（条）";
                int t = 0;
                foreach (string sci in GDef.sciencelist)
                {
                    dt_sourceview.RowFilter = string.Format("science = '{0}' and ab_type_name = '{1}'", sci, abtype);
                    r[sci] = dt_sourceview.Count;
                    t += Convert.ToInt32(r[sci]);
                }
                r["总计"] = t;
                dt_总表.Rows.Add(r);
            }
            foreach (string abtype in GDef.abtypelist)
            {
                DataRow r = dt_总表.NewRow();
                r[0] = abtype + "（套）";
                int t = 0;
                foreach (string sci in GDef.sciencelist)
                {
                    dt_sourceview.RowFilter = string.Format("science = '{0}' and ab_type_name = '{1}'", sci, abtype);
                    r[sci] = dt_sourceview.ToTable(true, "sp").Rows.Count;
                    t += Convert.ToInt32(r[sci]);
                }
                r["总计"] = t;
                dt_总表.Rows.Add(r);
            }
            {
                DataRow r = dt_总表.NewRow();
                r[0] = "总事件数（条）";
                int t = 0;
                foreach (string sci in GDef.sciencelist)
                {
                    dt_sourceview.RowFilter = string.Format("science = '{0}'", sci);
                    r[sci] = dt_sourceview.Count;
                    t += Convert.ToInt32(r[sci]);
                }
                r["总计"] = t;
                dt_总表.Rows.Add(r);
            }
            {
                DataRow r = dt_总表.NewRow();
                r[0] = "仪器总数（套）";
                int t = 0;
                foreach (string sci in GDef.sciencelist)
                {
                    dt_spview.RowFilter = string.Format("science = '{0}'", sci);
                    r[sci] = dt_spview.Count;
                    t += Convert.ToInt32(r[sci]);
                }
                r["总计"] = t;
                dt_总表.Rows.Add(r);
            }
            dt_总表.PrimaryKey = new DataColumn[] { dt_总表.Columns["学科"] };
            {
                DataRow r = dt_总表.NewRow();
                r[0] = "平均每套产出事件数";
                foreach (string sci in GDef.sciencelist)
                {
                    r[sci] = Math.Round(Convert.ToDouble(dt_总表.Rows.Find("总事件数（条）")[sci]) / Convert.ToDouble(dt_总表.Rows.Find("仪器总数（套）")[sci]), 2);
                }
                r["总计"] = Math.Round(Convert.ToDouble(dt_总表.Rows.Find("总事件数（条）")["总计"]) / Convert.ToDouble(dt_总表.Rows.Find("仪器总数（套）")["总计"]), 2);
                dt_总表.Rows.Add(r);
            }
            DataTable re = dth.Transpose(dt_总表);
            if (writetosheet)
                dth.DTToExcelSheet(re, book, "dt_总表");
            return re;
        }
    }
}