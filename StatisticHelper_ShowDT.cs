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
        public void ShowDT(bool showtoeapp = true)
        {
            excel.Application eapp = new excel.Application();
            excel.Workbook book = eapp.Workbooks.Add();
            dth.DTToExcelSheet(dt_eval_instrid.AsEnumerable().Take(3).CopyToDataTable(), book, "dt_eval_instrid");
            dth.DTToExcelSheet(dt_eval_instrid_comp.AsEnumerable().Take(3).CopyToDataTable(), book, "dt_eval_instrid_comp");
            dth.DTToExcelSheet(dt_eval_y_instrid_comp.AsEnumerable().Take(3).CopyToDataTable(), book, "dt_eval_y_instrid_comp");
            dth.DTToExcelSheet(dt_evaltab.AsEnumerable().Take(3).CopyToDataTable(), book, "dt_evaltab");
            dth.DTToExcelSheet(dt_log.AsEnumerable().Take(3).CopyToDataTable(), book, "dt_log");
            dth.DTToExcelSheet(dt_loginstr.AsEnumerable().Take(3).CopyToDataTable(), book, "dt_loginstr");
            dth.DTToExcelSheet(dt_no_check.AsEnumerable().Take(3).CopyToDataTable(), book, "dt_no_check");
            dth.DTToExcelSheet(dt_source.AsEnumerable().Take(3).CopyToDataTable(), book, "dt_source");
            dth.DTToExcelSheet(dt_sp.AsEnumerable().Take(3).CopyToDataTable(), book, "dt_sp");
            dth.DTToExcelSheet(dt_ncheck.AsEnumerable().Take(3).CopyToDataTable(), book, "dt_ncheck");
            eapp.Visible = true;
        }
    }
}