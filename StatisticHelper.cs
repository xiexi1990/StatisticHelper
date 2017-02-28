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
        DateTime DATEBEGIN, DATEEND;
        OraHelper oh;
        SqlGenerator sg;
        DataTableHelper dth;
        double comp_require = 0.95;
        DateMergerPrecision prectyp = DateMergerPrecision.Day;
        DataTable 
            dt_stationlist, 
            dt_source, 
            dt_sp, 
            dt_log, 
            dt_no_check, 
            dt_evaltab, 
            dt_eval_instrid, 
            dt_eval_y_instrid, 
            dt_eval_noaux_instrid, 
            dt_eval_not_y_instrid, 
            dt_eval_instrid_comp, 
            dt_eval_y_instrid_comp, 
            dt_eval_not_y_instrid_comp, 
            dt_loginstr, 
            dt_ncheck;
        
        public StatisticHelper(string ohparam)
        {
            DATEBEGIN = new DateTime(2016, 1, 1, 0, 0, 0);
            DATEEND = new DateTime(2016, 12, 31, 23, 59, 59);
            sg = new SqlGenerator();
            dth = new DataTableHelper();
            dth.outcnt = 0;
            oh = new OraHelper(ohparam, true);
            oh.Set_pdbqz_qzdata();
            oh.cache_dt_datebegin = DATEBEGIN;
            oh.cache_dt_dateend = DATEEND;        
        }
        


    }
}
