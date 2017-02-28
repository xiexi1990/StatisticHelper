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


        public void PrepareTabs()
        {
            dt_stationlist = oh.GetOrCacheTab("dt_stationlist", "select distinct stationid, stationname from qzdata.qz_dict_stations");

            dt_source = oh.GetOrCacheTab("dt_source", sg.GenSTabSql(DATEBEGIN, DATEEND));
            dt_sp = oh.GetOrCacheTab("dt_sp", @"select distinct b.stationid||'_'||b.pointid sp, decode(substr(b.itemid,0,3), '411','水位','431','水温', d.item) item, b.science, D.INSTRCODE, D.INSTRTYPE||D.INSTRNAME instr, F.UNITNAME, E.STATIONNAME from qzdata.qz_pj_evalist1 b, qzdata.qz_dict_stationinstruments c, qzdata.qz_abnormity_instrinfo d, qzdata.qz_dict_stations e, qz_abnormity_units f where b.science <> '辅助'  and b.stationid = C.stationid and b.pointid = C.pointid and c.instrcode = d.instrcode  and b.runstatus = '在运行' and b.stationid = E.STATIONID and E.UNITCODE = F.UNIT_CODE");
            string log_string = @"select distinct a.log_id, b.unitcode, a.stationid, b.science, a.stationid || '_' || a.pointid sp, a.start_date, decode(a.end_date, null, to_date('20990101 000000', 'yyyymmdd hh24miss'), a.end_date) end_date, a.ab_id from qzdata.qz_abnormity_log a, qzdata.qz_pj_evalist1 b where a.ab_id >= 1 and a.ab_id <= 7 and a.stationid = b.stationid and a.pointid = b.pointid  and b.runstatus = '在运行' and b.science <> '辅助' and (STDDATE)";
            dt_log = oh.GetOrCacheTab("dt_log", log_string.Replace("STDDATE", GDef.std_datestr(DATEBEGIN, DATEEND)));
            dt_no_check = oh.GetOrCacheTab("dt_no_check", string.Format(@"select unitcode, count(log_id) cnt from(select distinct a.log_id, b.unitcode from qzdata.qz_abnormity_log a, qzdata.qz_pj_evalist1 b where ({0}) and a.stationid = b.stationid and a.pointid = b.pointid and B.SCIENCE <> '辅助' and a.ab_id <> 1 and b.runstatus = '在运行') where log_id not in (select distinct log_id from qzdata.qz_abnormity_check) group by unitcode", GDef.std_datestr(DATEBEGIN, DATEEND)));
            dt_ncheck = oh.GetOrCacheTab("dt_ncheck", string.Format(@"select log_id,flag_5 sgroup, flag_3 stime, flag_2 slog, is_agree sgraph,     reason, check_date from qzdata.qz_abnormity_ncheck where check_date >= to_date('{0}', 'yyyymmdd hh24miss') and check_date <= to_date('{1}', 'yyyymmdd hh24miss')", DATEBEGIN.ToString(GDef.date_tostring_format), DATEEND.ToString(GDef.date_tostring_format)));

            {
                string filename = GDef.store_qztemp + GDef.store_dtprefix + "dt_evaltab" + DATEBEGIN.ToString(GDef.date_tostring_format) + DATEEND.ToString(GDef.date_tostring_format);
                if (File.Exists(filename + "dt"))
                {
                    dt_evaltab = new DataTable();
                    dt_evaltab.ReadXmlSchema(filename + "sch");
                    dt_evaltab.ReadXml(filename + "dt");
                }
                else
                {
                    dt_evaltab = oh.GetDataTable("select distinct b.stationid, b.pointid, b.stationid||'_'||b.pointid as sp, d.instrname, d.instrtype || d.instrname fullinstrname, b.unitcode, b.ab_flag, b.science from qzdata.qz_pj_evalist1 b, qzdata.qz_dict_stationinstruments c, qzdata.qz_abnormity_instrinfo d where b.runstatus = '在运行' and b.science <> '辅助' and b.stationid = c.stationid and b.pointid = c.pointid and c.instrcode = d.instrcode");
                    dt_evaltab.Columns.Add("comp", typeof(double));
                    Console.WriteLine("calculating instr comp ...");
                    DataTable tmpsp = dt_evaltab.DefaultView.ToTable(true, "sp");
                    //              dth.DTToExcel(tmpsp, "tmp111", true, false);
                    tmpsp.Columns.Add("comp", typeof(double));
                    tmpsp.CaseSensitive = true;
                    tmpsp.PrimaryKey = new DataColumn[] { tmpsp.Columns["sp"] };

                    int i = 0;
                    foreach (DataRow r in tmpsp.Rows)
                    {
                        r["comp"] = CalCompInstr(dt_log, r["sp"].ToString(), DATEBEGIN, DATEEND, prectyp);
                        i++;
                        if (i % 1 == 0)
                            Console.Write("\rfinished " + i);
                    }
                    foreach (DataRow r in dt_evaltab.Rows)
                    {
                        r["comp"] = tmpsp.Rows.Find(r["sp"])["comp"];
                    }
                    Console.WriteLine();
                    dt_evaltab.TableName = "dt_evaltab";
                    dt_evaltab.WriteXmlSchema(filename + "sch");
                    dt_evaltab.WriteXml(filename + "dt");
                }
            }
            DataView dt_evaltabview = new DataView(dt_evaltab);
            dt_eval_instrid = dt_evaltabview.ToTable(true, new string[] { "stationid", "pointid", "sp", "instrname", "fullinstrname", "unitcode", "science", "comp" });
            Console.WriteLine("dt_eval_instrid count = " + dt_eval_instrid.Rows.Count);

            dt_evaltabview.RowFilter = "ab_flag = 'Y'";
            dt_eval_y_instrid = dt_evaltabview.ToTable(true, new string[] { "stationid", "pointid", "sp", "instrname", "fullinstrname", "unitcode", "science", "comp" });
            Console.WriteLine("dt_eval_y_instrid count = " + dt_eval_y_instrid.Rows.Count);

            dt_evaltabview.RowFilter = "science <> '辅助'";
            dt_eval_noaux_instrid = dt_evaltabview.ToTable(true, new string[] { "stationid", "pointid", "sp", "instrname", "fullinstrname", "unitcode", "science", "comp" });
            IEnumerable<DataRow> query1 = dt_eval_noaux_instrid.AsEnumerable().Except(dt_eval_y_instrid.AsEnumerable(), DataRowComparer.Default);
            dt_eval_not_y_instrid = query1.CopyToDataTable();
            Console.WriteLine("dt_eval_not_y_instrid count = " + dt_eval_not_y_instrid.Rows.Count);
            string comp_req = "comp >= " + comp_require;
            dt_eval_instrid.DefaultView.RowFilter = dt_eval_y_instrid.DefaultView.RowFilter = dt_eval_not_y_instrid.DefaultView.RowFilter = comp_req;
            dt_eval_instrid_comp = dt_eval_instrid.DefaultView.ToTable();
            dt_eval_y_instrid_comp = dt_eval_y_instrid.DefaultView.ToTable();
            dt_eval_not_y_instrid_comp = dt_eval_not_y_instrid.DefaultView.ToTable();
            dt_loginstr = dt_log.DefaultView.ToTable(true, "sp");
            Console.WriteLine("dt_loginstr count = " + dt_loginstr.Rows.Count);
            dt_loginstr.Columns.Add("unitcode");

            foreach (DataRow r in dt_loginstr.Rows)
            {
                object[] info = dth.ExtractRowByLeftFirstCol_WithoutKey(dt_evaltab, r["sp"].ToString());
                if (info == null)
                {
                    continue;
                }
                r["unitcode"] = info[0];
            }

        }
    }
}